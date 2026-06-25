import { json, requireAuth } from "../../_lib/auth.js";

export async function onRequestGet(context) {
  const denied = await requireAuth(context); if (denied) return denied;
  const id = context.params.id;
  const closure = await context.env.DB.prepare("SELECT * FROM attendance_closures WHERE id = ?").bind(id).first();
  if (!closure) return json({ error: "기록을 찾지 못했습니다." }, 404);
  const items = await context.env.DB.prepare("SELECT * FROM attendance_missing_items WHERE closure_id = ? ORDER BY missing_date, store, employee_name").bind(id).all();
  return json({ closure, items: items.results || [] });
}

export async function onRequestDelete(context) {
  const denied = await requireAuth(context); if (denied) return denied;
  await context.env.DB.batch([
    context.env.DB.prepare("DELETE FROM attendance_missing_items WHERE closure_id = ?").bind(context.params.id),
    context.env.DB.prepare("DELETE FROM attendance_closures WHERE id = ?").bind(context.params.id),
  ]);
  return json({ ok: true });
}
