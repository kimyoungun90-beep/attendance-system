import { json, requireAuth } from "../../_lib/auth.js";
import { ensureSchema } from "../../_lib/schema.js";
import { recalculateRoute } from "../../_lib/recalculate.js";

export async function onRequestGet(context) {
  const denied = await requireAuth(context);
  if (denied) return denied;
  await ensureSchema(context.env.DB);
  const month = String(context.params.month || "");
  if (!/^\d{4}-\d{2}$/.test(month)) return json({ error: "적용 월이 올바르지 않습니다." }, 400);
  const item = await context.env.DB.prepare(`SELECT * FROM attendance_workforce_uploads WHERE month = ? LIMIT 1`).bind(month).first();
  if (!item) return json({ error: "인력·매장매칭 원본 파일을 찾지 못했습니다." }, 404);

  if (item.storage_type === "r2") {
    if (!context.env.FILES || !item.object_key) return json({ error: "R2 파일 저장소 연결을 확인해 주세요." }, 500);
    const object = await context.env.FILES.get(item.object_key);
    if (!object) return json({ error: "저장된 원본 파일을 찾지 못했습니다." }, 404);
    return new Response(object.body, { headers: downloadHeaders(item) });
  }

  const chunks = await context.env.DB.prepare(`
    SELECT chunk_blob FROM attendance_workforce_file_chunks
    WHERE upload_id = ? ORDER BY chunk_index
  `).bind(item.id).all();
  const arrays = (chunks.results || []).map((row) => new Uint8Array(row.chunk_blob));
  const total = arrays.reduce((sum, array) => sum + array.byteLength, 0);
  const merged = new Uint8Array(total);
  let offset = 0;
  for (const array of arrays) {
    merged.set(array, offset);
    offset += array.byteLength;
  }
  return new Response(merged, { headers: downloadHeaders(item) });
}

export async function onRequestDelete(context) {
  const denied = await requireAuth(context);
  if (denied) return denied;
  await ensureSchema(context.env.DB);
  const month = String(context.params.month || "");
  const item = await context.env.DB.prepare(`SELECT * FROM attendance_workforce_uploads WHERE month = ? LIMIT 1`).bind(month).first();
  if (!item) return json({ error: "삭제할 인력·매장매칭 자료를 찾지 못했습니다." }, 404);
  if (item.storage_type === "r2" && item.object_key && context.env.FILES) await context.env.FILES.delete(item.object_key).catch(() => {});
  await context.env.DB.batch([
    context.env.DB.prepare(`DELETE FROM attendance_workforce_members WHERE upload_id = ?`).bind(item.id),
    context.env.DB.prepare(`DELETE FROM attendance_workforce_file_chunks WHERE upload_id = ?`).bind(item.id),
    context.env.DB.prepare(`DELETE FROM attendance_workforce_uploads WHERE id = ?`).bind(item.id),
  ]);
  let affectedMonths = 0;
  for (const route of ["homeplus", "electroland"]) {
    const recalculated = await recalculateRoute(context.env.DB, route);
    affectedMonths += Number(recalculated.affectedMonths || 0);
  }
  return json({ ok: true, affectedMonths });
}

function downloadHeaders(item) {
  return {
    "content-type": item.content_type || "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    "content-length": String(item.size_bytes || ""),
    "content-disposition": `attachment; filename*=UTF-8''${encodeURIComponent(item.file_name || `${item.month}_workforce.xlsx`)}`,
    "cache-control": "private, no-store",
  };
}
