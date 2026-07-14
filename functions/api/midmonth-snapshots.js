import { json, requireAuth } from "../_lib/auth.js";
import { ensureSchema } from "../_lib/schema.js";

const VALID_ROUTES = new Set(["homeplus", "electroland"]);

export async function onRequestGet(context) {
  const denied = await requireAuth(context);
  if (denied) return denied;
  await ensureSchema(context.env.DB);

  const url = new URL(context.request.url);
  const route = String(url.searchParams.get("route") || "");
  const month = String(url.searchParams.get("month") || "");
  if (!VALID_ROUTES.has(route) || !/^\d{4}-\d{2}$/.test(month)) {
    return json({ error: "경로와 대상 월을 확인해 주세요." }, 400);
  }

  const result = await context.env.DB.prepare(`
    SELECT id, route, month, cutoff_date AS cutoffDate, snapshot_name AS snapshotName,
           employee_count AS employeeCount, daily_count AS dailyCount,
           created_at AS createdAt, updated_at AS updatedAt
    FROM attendance_midmonth_snapshots
    WHERE route = ? AND month = ? AND status = 'active'
    ORDER BY cutoff_date DESC, created_at DESC
  `).bind(route, month).all();

  return json({ items: result.results || [] });
}

export async function onRequestPost(context) {
  const denied = await requireAuth(context);
  if (denied) return denied;
  await ensureSchema(context.env.DB);

  const body = await context.request.json().catch(() => null);
  const route = String(body?.route || "");
  const month = String(body?.month || "");
  const cutoffDate = String(body?.cutoffDate || "");
  const employeeFacts = Array.isArray(body?.employeeFacts) ? body.employeeFacts.slice(0, 5000) : [];
  if (!body || !VALID_ROUTES.has(route) || !/^\d{4}-\d{2}$/.test(month) || !/^\d{4}-\d{2}-\d{2}$/.test(cutoffDate) || !cutoffDate.startsWith(month)) {
    return json({ error: "경로, 대상 월, 확정 기준일을 확인해 주세요." }, 400);
  }
  if (!employeeFacts.length) return json({ error: "저장할 상담사근태 확정값이 없습니다." }, 400);

  const old = await context.env.DB.prepare(`
    SELECT id FROM attendance_midmonth_snapshots
    WHERE route = ? AND month = ? AND cutoff_date = ? AND status = 'active'
    LIMIT 1
  `).bind(route, month, cutoffDate).first();
  const replaced = Boolean(old?.id);
  const snapshotId = old?.id || crypto.randomUUID();
  const rows = [];
  const employeeIds = new Set();
  for (const fact of employeeFacts) {
    const employeeId = String(fact?.employeeId || "").trim();
    if (!employeeId) continue;
    employeeIds.add(employeeId);
    const dailyStatuses = Array.isArray(fact?.dailyStatuses) ? fact.dailyStatuses : [];
    for (const item of dailyStatuses) {
      const date = String(item?.date || "");
      if (!/^\d{4}-\d{2}-\d{2}$/.test(date) || !date.startsWith(month) || date > cutoffDate) continue;
      rows.push({
        employeeId,
        employeeName: String(fact?.name || item?.name || ""),
        store: String(fact?.store || item?.store || ""),
        date,
        planStatus: String(item?.planStatus || "공백"),
        hasClockIn: item?.hasClockIn ? 1 : 0,
        actualStatus: String(item?.actualStatus || ""),
        actualIn: String(item?.actualIn || ""),
        changedIn: String(item?.changedIn || ""),
        displayStatus: String(item?.displayStatus || ""),
        evidenced: item?.evidenced ? 1 : 0,
        source: String(item?.source || "midmonth"),
      });
    }
  }
  if (!rows.length) return json({ error: `${cutoffDate} 이전에 저장할 날짜별 확정값이 없습니다.` }, 400);

  const snapshotName = String(body?.snapshotName || `${cutoffDate} 중간 확인`).slice(0, 80);
  try {
    const statements = [
      context.env.DB.prepare("DELETE FROM attendance_midmonth_snapshot_items WHERE snapshot_id = ?").bind(snapshotId),
      context.env.DB.prepare(`
        INSERT INTO attendance_midmonth_snapshots
        (id, route, month, cutoff_date, snapshot_name, status, employee_count, daily_count, created_at, updated_at)
        VALUES (?, ?, ?, ?, ?, 'active', ?, ?, datetime('now'), datetime('now'))
        ON CONFLICT(route, month, cutoff_date) DO UPDATE SET
          snapshot_name = excluded.snapshot_name,
          status = 'active',
          employee_count = excluded.employee_count,
          daily_count = excluded.daily_count,
          updated_at = datetime('now')
      `).bind(snapshotId, route, month, cutoffDate, snapshotName, employeeIds.size, rows.length),
    ];
    await context.env.DB.batch(statements);

    const itemStatements = rows.map((row) => context.env.DB.prepare(`
      INSERT OR REPLACE INTO attendance_midmonth_snapshot_items
      (snapshot_id, route, month, employee_id, employee_name, store, status_date,
       plan_status, has_clock_in, actual_status, actual_in, changed_in, display_status, evidenced, source)
      VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    `).bind(
      snapshotId, route, month, row.employeeId, row.employeeName, row.store, row.date,
      row.planStatus, row.hasClockIn, row.actualStatus, row.actualIn,
      row.changedIn, row.displayStatus, row.evidenced, row.source
    ));
    await runBatch(context.env.DB, itemStatements);

    return json({ ok: true, id: snapshotId, replaced, route, month, cutoffDate, employeeCount: employeeIds.size, dailyCount: rows.length }, 201);
  } catch (error) {
    return json({ error: `중간 확인 저장 중 오류가 발생했습니다: ${error.message || "알 수 없는 오류"}` }, 500);
  }
}

export async function onRequestDelete(context) {
  const denied = await requireAuth(context);
  if (denied) return denied;
  await ensureSchema(context.env.DB);

  const url = new URL(context.request.url);
  const id = String(url.searchParams.get("id") || "");
  if (!id) return json({ error: "삭제할 중간 저장본 ID가 없습니다." }, 400);

  const existing = await context.env.DB.prepare(`
    SELECT id, route, month, cutoff_date FROM attendance_midmonth_snapshots
    WHERE id = ? AND status = 'active'
    LIMIT 1
  `).bind(id).first();
  if (!existing?.id) return json({ error: "삭제할 중간 저장본을 찾지 못했습니다." }, 404);

  await context.env.DB.batch([
    context.env.DB.prepare("DELETE FROM attendance_midmonth_snapshot_items WHERE snapshot_id = ?").bind(id),
    context.env.DB.prepare("DELETE FROM attendance_midmonth_snapshots WHERE id = ?").bind(id),
  ]);
  return json({ ok: true, deleted: existing });
}

async function runBatch(db, statements) {
  for (let index = 0; index < statements.length; index += 40) {
    await db.batch(statements.slice(index, index + 40));
  }
}
