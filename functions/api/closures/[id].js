import { json, requireAuth } from "../../_lib/auth.js";
import { recalculateRoute } from "../../_lib/recalculate.js";
import { ensureSchema } from "../../_lib/schema.js";

export async function onRequestGet(context) {
  const denied = await requireAuth(context);
  if (denied) return denied;
  await ensureSchema(context.env.DB);

  const id = context.params.id;
  const closure = await context.env.DB.prepare(`
    SELECT *, company AS route
    FROM attendance_closures
    WHERE id = ?
  `).bind(id).first();
  if (!closure) return json({ error: "월 마감 기록을 찾지 못했습니다." }, 404);

  const [issues, mismatches, facts, summaries, allocations] = await Promise.all([
    context.env.DB.prepare(`
      SELECT * FROM attendance_issue_items
      WHERE closure_id = ?
      ORDER BY issue_type, issue_date, store, employee_name
    `).bind(id).all(),
    context.env.DB.prepare(`
      SELECT * FROM attendance_mismatch_items
      WHERE closure_id = ?
      ORDER BY issue_date, store, employee_name
    `).bind(id).all(),
    context.env.DB.prepare(`
      SELECT * FROM attendance_monthly_employee_facts
      WHERE closure_id = ?
      ORDER BY store, employee_name
    `).bind(id).all(),
    context.env.DB.prepare(`
      SELECT * FROM attendance_monthly_summaries_v3
      WHERE closure_id = ?
      ORDER BY MAX(shortage, compensation_shortage) DESC, base_excess DESC, store, employee_name
    `).bind(id).all(),
    context.env.DB.prepare(`
      SELECT a.*, g.grant_month, g.valid_from, g.valid_to, g.reason, g.grant_type, g.grant_scope
      FROM attendance_leave_allocations_v5 a
      JOIN attendance_leave_grants_v5 g ON g.id = a.grant_id
      WHERE a.closure_id = ?
      ORDER BY g.valid_to, a.employee_id
    `).bind(id).all(),
  ]);

  return json({
    closure,
    issues: issues.results || [],
    mismatches: mismatches.results || [],
    facts: facts.results || [],
    summaries: summaries.results || [],
    allocations: allocations.results || [],
  });
}

export async function onRequestDelete(context) {
  const denied = await requireAuth(context);
  if (denied) return denied;
  await ensureSchema(context.env.DB);

  const id = context.params.id;
  const closure = await context.env.DB.prepare(
    "SELECT company AS route FROM attendance_closures WHERE id = ?"
  ).bind(id).first();
  if (!closure) return json({ error: "월 마감 기록을 찾지 못했습니다." }, 404);

  await context.env.DB.batch([
    context.env.DB.prepare("DELETE FROM attendance_leave_allocations_v5 WHERE closure_id = ?").bind(id),
    context.env.DB.prepare("DELETE FROM route_substitute_allocations WHERE closure_id = ?").bind(id),
    context.env.DB.prepare("DELETE FROM attendance_monthly_summaries_v3 WHERE closure_id = ?").bind(id),
    context.env.DB.prepare("DELETE FROM attendance_monthly_employee_facts WHERE closure_id = ?").bind(id),
    context.env.DB.prepare("DELETE FROM attendance_mismatch_items WHERE closure_id = ?").bind(id),
    context.env.DB.prepare("DELETE FROM attendance_issue_items WHERE closure_id = ?").bind(id),
    context.env.DB.prepare("DELETE FROM attendance_missing_items WHERE closure_id = ?").bind(id),
    context.env.DB.prepare("DELETE FROM attendance_closures WHERE id = ?").bind(id),
  ]);

  const recalculated = await recalculateRoute(context.env.DB, closure.route);
  return json({ ok: true, affectedMonths: recalculated.affectedMonths });
}
