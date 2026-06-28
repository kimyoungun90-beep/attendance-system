import { json, requireAuth } from "../_lib/auth.js";
import { recalculateRoute } from "../_lib/recalculate.js";
import { ensureSchema } from "../_lib/schema.js";

const VALID_ROUTES = new Set(["homeplus", "electroland"]);
const D1_FILE_LIMIT = 20 * 1024 * 1024;
const D1_CHUNK_SIZE = 450_000;

export async function onRequestGet(context) {
  const denied = await requireAuth(context);
  if (denied) return denied;
  await ensureSchema(context.env.DB);

  const url = new URL(context.request.url);
  const month = String(url.searchParams.get("month") || "");
  const route = String(url.searchParams.get("route") || "");
  if (route && !VALID_ROUTES.has(route)) return json({ error: "경로 구분이 올바르지 않습니다." }, 400);

  if (month) {
    if (!/^\d{4}-\d{2}$/.test(month)) return json({ error: "적용 월이 올바르지 않습니다." }, 400);
    const upload = await context.env.DB.prepare(`
      SELECT id, month, file_name, content_type, size_bytes, storage_type,
             employee_count, electroland_count, homeplus_count, portal_count,
             created_at, updated_at
      FROM attendance_workforce_uploads WHERE month = ? LIMIT 1
    `).bind(month).first();
    if (!upload) return json({ upload: null, members: [] });

    let statement = context.env.DB.prepare(`
      SELECT route, regional_manager, manager, region1, region2, store_code, store_name,
             portal_id, employee_id, employee_name, hire_date, group_hire_date, note
      FROM attendance_workforce_members
      WHERE month = ? ${route ? "AND route = ?" : ""}
      ORDER BY route, regional_manager, manager, region2, store_code, employee_name
    `);
    statement = route ? statement.bind(month, route) : statement.bind(month);
    const members = await statement.all();
    return json({ upload, members: (members.results || []).map(toClientMember) });
  }

  const result = await context.env.DB.prepare(`
    SELECT id, month, file_name, content_type, size_bytes, storage_type,
           employee_count, electroland_count, homeplus_count, portal_count,
           created_at, updated_at
    FROM attendance_workforce_uploads
    ORDER BY month DESC, updated_at DESC
    LIMIT 120
  `).all();
  return json({ items: result.results || [], d1FileLimit: D1_FILE_LIMIT, r2Configured: Boolean(context.env.FILES) });
}

export async function onRequestPost(context) {
  const denied = await requireAuth(context);
  if (denied) return denied;
  await ensureSchema(context.env.DB);

  const form = await context.request.formData().catch(() => null);
  const file = form?.get("file");
  const month = String(form?.get("month") || "");
  const membersRaw = parseJson(form?.get("members"), []);
  const portalMappingsRaw = parseJson(form?.get("portalMappings"), []);

  if (!(file instanceof File) || file.size <= 0) return json({ error: "인력 및 매장매칭 엑셀을 선택해 주세요." }, 400);
  if (!/^\d{4}-\d{2}$/.test(month)) return json({ error: "적용 월을 확인해 주세요." }, 400);
  if (!/\.(xlsx|xls|xlsb)$/i.test(file.name)) return json({ error: "엑셀 파일(xlsx, xls, xlsb)만 등록할 수 있습니다." }, 400);
  if (!context.env.FILES && file.size > D1_FILE_LIMIT) return json({ error: "현재 D1 파일 저장 한도는 20MB입니다." }, 413);

  const members = normalizeMembers(membersRaw, month);
  if (!members.length) return json({ error: "저장할 직원 매칭 정보가 없습니다." }, 400);
  const portalMappings = normalizePortalMappings(portalMappingsRaw);
  const old = await context.env.DB.prepare(`SELECT * FROM attendance_workforce_uploads WHERE month = ? LIMIT 1`).bind(month).first();
  const oldRoutesResult = old?.id
    ? await context.env.DB.prepare(`SELECT DISTINCT route FROM attendance_workforce_members WHERE upload_id = ?`).bind(old.id).all()
    : { results: [] };
  const replaced = Boolean(old?.id);
  const id = crypto.randomUUID();
  const contentType = file.type || "application/octet-stream";
  let storageType = "d1";
  let objectKey = null;

  try {
    if (context.env.FILES) {
      storageType = "r2";
      objectKey = `workforce/${month}/${id}-${sanitizeFileName(file.name)}`;
      await context.env.FILES.put(objectKey, file.stream(), {
        httpMetadata: { contentType },
        customMetadata: { originalName: file.name, month },
      });
    }

    for (const mapping of portalMappings) {
      await context.env.DB.prepare(`
        INSERT INTO attendance_portal_ids (employee_id, portal_id, employee_name, updated_at)
        VALUES (?, ?, ?, datetime('now'))
        ON CONFLICT(employee_id) DO UPDATE SET
          portal_id = excluded.portal_id,
          employee_name = COALESCE(NULLIF(excluded.employee_name, ''), attendance_portal_ids.employee_name),
          updated_at = datetime('now')
      `).bind(mapping.employeeId, mapping.portalId, mapping.employeeName || "").run();
    }

    const savedPortalsResult = await context.env.DB.prepare(`SELECT employee_id, portal_id FROM attendance_portal_ids`).all();
    const savedPortals = new Map((savedPortalsResult.results || []).map((row) => [normalizeId(row.employee_id), String(row.portal_id || "")]));
    const finalMembers = members.map((member) => ({
      ...member,
      portalId: member.portalId || savedPortals.get(member.employeeId) || "",
    }));
    const electrolandCount = finalMembers.filter((row) => row.route === "electroland").length;
    const homeplusCount = finalMembers.filter((row) => row.route === "homeplus").length;
    const portalCount = finalMembers.filter((row) => row.portalId).length;

    if (old?.id) {
      await context.env.DB.batch([
        context.env.DB.prepare(`DELETE FROM attendance_workforce_members WHERE upload_id = ?`).bind(old.id),
        context.env.DB.prepare(`DELETE FROM attendance_workforce_file_chunks WHERE upload_id = ?`).bind(old.id),
        context.env.DB.prepare(`DELETE FROM attendance_workforce_uploads WHERE id = ?`).bind(old.id),
      ]);
    }

    await context.env.DB.prepare(`
      INSERT INTO attendance_workforce_uploads
      (id, month, file_name, content_type, size_bytes, storage_type, object_key,
       employee_count, electroland_count, homeplus_count, portal_count, created_at, updated_at)
      VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, datetime('now'), datetime('now'))
    `).bind(
      id, month, file.name, contentType, file.size, storageType, objectKey,
      finalMembers.length, electrolandCount, homeplusCount, portalCount
    ).run();

    if (storageType === "d1") {
      const buffer = await file.arrayBuffer();
      const chunkStatements = [];
      for (let offset = 0, chunkIndex = 0; offset < buffer.byteLength; offset += D1_CHUNK_SIZE, chunkIndex += 1) {
        const chunk = buffer.slice(offset, Math.min(offset + D1_CHUNK_SIZE, buffer.byteLength));
        chunkStatements.push(context.env.DB.prepare(`
          INSERT INTO attendance_workforce_file_chunks (upload_id, chunk_index, chunk_blob)
          VALUES (?, ?, ?)
        `).bind(id, chunkIndex, chunk));
      }
      await runBatch(context.env.DB, chunkStatements);
    }

    const memberStatements = finalMembers.map((member) => context.env.DB.prepare(`
      INSERT INTO attendance_workforce_members
      (upload_id, month, route, regional_manager, manager, region1, region2,
       store_code, store_name, portal_id, employee_id, employee_name,
       hire_date, group_hire_date, note)
      VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    `).bind(
      id, month, member.route, member.regionalManager, member.manager, member.region1, member.region2,
      member.storeCode, member.storeName, member.portalId, member.employeeId, member.employeeName,
      member.hireDate || null, member.groupHireDate || null, member.note
    ));
    await runBatch(context.env.DB, memberStatements);

    if (old?.storage_type === "r2" && old.object_key && context.env.FILES) {
      await context.env.FILES.delete(old.object_key).catch(() => {});
    }

    const affectedRoutes = new Set([
      ...(oldRoutesResult.results || []).map((row) => row.route),
      ...finalMembers.map((row) => row.route),
    ].filter((value) => VALID_ROUTES.has(value)));
    let recalculatedMonths = 0;
    for (const affectedRoute of affectedRoutes) {
      try {
        const recalculated = await recalculateRoute(context.env.DB, affectedRoute);
        recalculatedMonths += Number(recalculated.affectedMonths || 0);
      } catch (error) {
        console.error("인력자료 저장 후 휴가 잔여 재계산 실패", affectedRoute, error);
      }
    }

    return json({
      ok: true,
      id,
      replaced,
      employeeCount: finalMembers.length,
      electrolandCount,
      homeplusCount,
      portalCount,
      storageType,
      recalculatedMonths,
    }, replaced ? 200 : 201);
  } catch (error) {
    if (storageType === "r2" && objectKey && context.env.FILES) await context.env.FILES.delete(objectKey).catch(() => {});
    await context.env.DB.prepare(`DELETE FROM attendance_workforce_members WHERE upload_id = ?`).bind(id).run().catch(() => {});
    await context.env.DB.prepare(`DELETE FROM attendance_workforce_file_chunks WHERE upload_id = ?`).bind(id).run().catch(() => {});
    await context.env.DB.prepare(`DELETE FROM attendance_workforce_uploads WHERE id = ?`).bind(id).run().catch(() => {});
    return json({ error: `인력·매장매칭 저장 중 오류가 발생했습니다: ${error.message || "알 수 없는 오류"}` }, 500);
  }
}

function normalizeMembers(items, month) {
  const map = new Map();
  for (const item of Array.isArray(items) ? items : []) {
    const route = String(item?.route || "");
    const employeeId = normalizeId(item?.employeeId || item?.employee_id);
    const employeeName = clean(item?.employeeName || item?.employee_name);
    const storeCode = clean(item?.storeCode || item?.store_code).replace(/\.0+$/, "");
    if (!VALID_ROUTES.has(route) || !employeeId || !employeeName) continue;
    const member = {
      month,
      route,
      regionalManager: clean(item?.regionalManager || item?.regional_manager),
      manager: clean(item?.manager),
      region1: clean(item?.region1),
      region2: clean(item?.region2),
      storeCode,
      storeName: clean(item?.storeName || item?.store_name),
      portalId: clean(item?.portalId || item?.portal_id),
      employeeId,
      employeeName,
      hireDate: normalizeDate(item?.hireDate || item?.hire_date),
      groupHireDate: normalizeDate(item?.groupHireDate || item?.group_hire_date),
      note: clean(item?.note),
    };
    map.set(`${route}|${employeeId}|${storeCode}`, member);
  }
  return [...map.values()];
}

function normalizePortalMappings(items) {
  const map = new Map();
  for (const item of Array.isArray(items) ? items : []) {
    const employeeId = normalizeId(item?.employeeId || item?.employee_id);
    const portalId = clean(item?.portalId || item?.portal_id);
    if (!employeeId || !portalId) continue;
    map.set(employeeId, { employeeId, portalId, employeeName: clean(item?.employeeName || item?.employee_name) });
  }
  return [...map.values()];
}

function toClientMember(row) {
  return {
    route: row.route,
    regionalManager: row.regional_manager || "",
    manager: row.manager || "",
    region1: row.region1 || "",
    region2: row.region2 || "",
    storeCode: row.store_code || "",
    storeName: row.store_name || "",
    portalId: row.portal_id || "",
    employeeId: row.employee_id || "",
    employeeName: row.employee_name || "",
    hireDate: row.hire_date || "",
    groupHireDate: row.group_hire_date || "",
    note: row.note || "",
  };
}

function parseJson(value, fallback) {
  try {
    const parsed = JSON.parse(String(value || ""));
    return parsed ?? fallback;
  } catch {
    return fallback;
  }
}

function normalizeId(value) {
  return String(value || "").trim().toUpperCase().replace(/\.0+$/, "").replace(/[\s\u00A0-]+/g, "").replace(/[^0-9A-Z가-힣]/g, "");
}

function normalizeDate(value) {
  const raw = clean(value);
  if (!raw) return "";
  const match = raw.match(/^(20\d{2})-(\d{2})-(\d{2})$/);
  return match ? raw : raw.slice(0, 30);
}

function clean(value) {
  return String(value ?? "").trim().slice(0, 500);
}

function sanitizeFileName(name) {
  return String(name || "workforce.xlsx").replace(/[^0-9A-Za-z가-힣._-]+/g, "_").slice(-120);
}

async function runBatch(db, statements) {
  for (let index = 0; index < statements.length; index += 40) {
    const chunk = statements.slice(index, index + 40);
    if (chunk.length) await db.batch(chunk);
  }
}
