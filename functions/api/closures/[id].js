import { json, requireAuth } from "../../_lib/auth.js";

export async function onRequestGet(context) {
  const denied = await requireAuth(context); if (denied) return denied;
  const id = context.params.id;
  const closure = await context.env.DB.prepare("SELECT * FROM attendance_closures WHERE id = ?").bind(id).first();
  if (!closure) return json({ error: "기록을 찾지 못했습니다." }, 404);
  const [issues, summaries, allocations] = await Promise.all([
    context.env.DB.prepare("SELECT * FROM attendance_issue_items WHERE closure_id = ? ORDER BY issue_type, issue_date, store, employee_name").bind(id).all(),
    context.env.DB.prepare("SELECT * FROM attendance_employee_summaries WHERE closure_id = ? ORDER BY shortage DESC, base_excess DESC, store, employee_name").bind(id).all(),
    context.env.DB.prepare(`
      SELECT a.*, g.employee_name, g.store, g.valid_from, g.valid_to, g.reason
      FROM substitute_dayoff_allocations a
      JOIN substitute_dayoff_grants g ON g.id = a.grant_id
      WHERE a.closure_id = ? ORDER BY g.valid_to, g.employee_name
    `).bind(id).all(),
  ]);
  return json({ closure, issues: issues.results || [], summaries: summaries.results || [], allocations: allocations.results || [] });
}

export async function onRequestDelete(context) {
  const denied = await requireAuth(context); if (denied) return denied;
  const id = context.params.id;
  await context.env.DB.batch([
    context.env.DB.prepare("DELETE FROM substitute_dayoff_allocations WHERE closure_id = ?").bind(id),
    context.env.DB.prepare("DELETE FROM attendance_employee_summaries WHERE closure_id = ?").bind(id),
    context.env.DB.prepare("DELETE FROM attendance_issue_items WHERE closure_id = ?").bind(id),
    context.env.DB.prepare("DELETE FROM attendance_missing_items WHERE closure_id = ?").bind(id),
    context.env.DB.prepare("DELETE FROM attendance_closures WHERE id = ?").bind(id),
  ]);
  return json({ ok: true });
}
