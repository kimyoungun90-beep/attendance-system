import { json, requireAuth } from "../../_lib/auth.js";

export async function onRequestDelete(context) {
  const denied = await requireAuth(context); if (denied) return denied;
  const id = context.params.id;
  const allocation = await context.env.DB.prepare("SELECT COUNT(*) AS count FROM substitute_dayoff_allocations WHERE grant_id = ?").bind(id).first();
  if (Number(allocation?.count || 0) > 0) return json({ error: "이미 월 마감에 사용된 대체휴무는 삭제할 수 없습니다." }, 409);
  await context.env.DB.prepare("DELETE FROM substitute_dayoff_grants WHERE id = ?").bind(id).run();
  return json({ ok: true });
}
