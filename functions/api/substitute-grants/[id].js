import { json, requireAuth } from "../../_lib/auth.js";
import { recalculateRoute } from "../../_lib/recalculate.js";

export async function onRequestDelete(context) {
  const denied = await requireAuth(context);
  if (denied) return denied;

  const id = context.params.id;
  const grant = await context.env.DB.prepare(
    "SELECT route FROM route_substitute_grants WHERE id = ?"
  ).bind(id).first();
  if (!grant) return json({ error: "공통 대체휴무 설정을 찾지 못했습니다." }, 404);

  await context.env.DB.prepare("DELETE FROM route_substitute_grants WHERE id = ?").bind(id).run();
  const recalculated = await recalculateRoute(context.env.DB, grant.route);
  return json({ ok: true, affectedMonths: recalculated.affectedMonths });
}
