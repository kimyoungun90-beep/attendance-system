import { json, requireAuth } from "../_lib/auth.js";

export async function onRequestGet(context) {
  const denied = await requireAuth(context); if (denied) return denied;
  const url = new URL(context.request.url);
  const company = url.searchParams.get("company");
  const month = url.searchParams.get("month");
  if (!["homeplus", "electroland"].includes(company || "") || !/^\d{4}-\d{2}$/.test(month || "")) {
    return json({ error: "회사와 대상 월을 확인해 주세요." }, 400);
  }
  const [year, monthNumber] = month.split("-").map(Number);
  const monthStart = `${month}-01`;
  const monthEnd = `${month}-${String(new Date(year, monthNumber, 0).getDate()).padStart(2, "0")}`;
  const result = await context.env.DB.prepare(`
    SELECT g.employee_id,
           SUM(MAX(0, g.granted_days - COALESCE(used.total_used, 0))) AS remaining_days
    FROM substitute_dayoff_grants g
    LEFT JOIN (
      SELECT grant_id, SUM(used_days) AS total_used
      FROM substitute_dayoff_allocations
      WHERE month <> ?
      GROUP BY grant_id
    ) used ON used.grant_id = g.id
    WHERE g.company = ? AND g.valid_from <= ? AND g.valid_to >= ?
    GROUP BY g.employee_id
  `).bind(month, company, monthEnd, monthStart).all();
  const balances = {};
  for (const item of result.results || []) balances[item.employee_id] = Math.round(Number(item.remaining_days || 0) * 2) / 2;
  return json({ balances });
}
