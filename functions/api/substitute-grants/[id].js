import { json, requireAuth } from "../../_lib/auth.js";
import { recalculateRoute } from "../../_lib/recalculate.js";
import { ensureSchema } from "../../_lib/schema.js";

export async function onRequestDelete(context) {
  const denied = await requireAuth(context);
  if (denied) return denied;
  await ensureSchema(context.env.DB);

  const id = context.params.id;
  const grant = await context.env.DB.prepare(
    "SELECT route FROM attendance_leave_grants_v5 WHERE id = ?"
  ).bind(id).first();
  if (!grant) return json({ error: "휴가 부여 설정을 찾지 못했습니다." }, 404);

  await context.env.DB.batch([
    context.env.DB.prepare("DELETE FROM attendance_leave_allocations_v5 WHERE grant_id = ?").bind(id),
    context.env.DB.prepare("DELETE FROM attendance_leave_grants_v5 WHERE id = ?").bind(id),
    // 구버전 공통 대체휴무에서 자동 이관된 기록도 함께 지워 재생성되지 않게 합니다.
    context.env.DB.prepare("DELETE FROM route_substitute_grants WHERE id = ?").bind(id),
  ]);
  const recalculated = await recalculateRoute(context.env.DB, grant.route);
  return json({ ok: true, affectedMonths: recalculated.affectedMonths });
}
