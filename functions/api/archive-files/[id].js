import { json, requireAuth } from "../../_lib/auth.js";
import { ensureSchema } from "../../_lib/schema.js";

export async function onRequestGet(context) {
  const denied = await requireAuth(context);
  if (denied) return denied;
  await ensureSchema(context.env.DB);

  const row = await context.env.DB.prepare(`
    SELECT id, file_name, content_type, size_bytes, storage_type, object_key, file_blob
    FROM attendance_archive_files WHERE id = ?
  `).bind(context.params.id).first();
  if (!row) return json({ error: "보관 파일을 찾지 못했습니다." }, 404);

  const headers = {
    "content-type": row.content_type || "application/octet-stream",
    "content-disposition": contentDisposition(row.file_name),
    "content-length": String(Number(row.size_bytes || 0)),
    "cache-control": "private, no-store",
  };

  if (row.storage_type === "r2") {
    if (!context.env.FILES) return json({ error: "R2 파일 저장소 바인딩(FILES)이 연결되지 않았습니다." }, 503);
    const object = await context.env.FILES.get(row.object_key);
    if (!object) return json({ error: "R2에서 파일을 찾지 못했습니다." }, 404);
    headers.etag = object.httpEtag;
    return new Response(object.body, { headers });
  }

  const legacyBytes = toUint8Array(row.file_blob);
  if (legacyBytes.byteLength) return new Response(legacyBytes, { headers });

  const countRow = await context.env.DB.prepare(`
    SELECT COUNT(*) AS chunk_count
    FROM attendance_archive_file_chunks
    WHERE file_id = ?
  `).bind(context.params.id).first();
  const chunkCount = Number(countRow?.chunk_count || 0);
  if (!chunkCount) return json({ error: "D1에 저장된 파일 데이터가 비어 있습니다." }, 404);

  let chunkIndex = 0;
  const stream = new ReadableStream({
    async pull(controller) {
      if (chunkIndex >= chunkCount) {
        controller.close();
        return;
      }
      try {
        const chunkRow = await context.env.DB.prepare(`
          SELECT chunk_blob
          FROM attendance_archive_file_chunks
          WHERE file_id = ? AND chunk_index = ?
          LIMIT 1
        `).bind(context.params.id, chunkIndex).first();
        const bytes = toUint8Array(chunkRow?.chunk_blob);
        if (!bytes.byteLength) throw new Error(`파일 조각 ${chunkIndex + 1}을 찾지 못했습니다.`);
        controller.enqueue(bytes);
        chunkIndex += 1;
      } catch (error) {
        controller.error(error);
      }
    },
  });
  return new Response(stream, { headers });
}

export async function onRequestDelete(context) {
  const denied = await requireAuth(context);
  if (denied) return denied;
  await ensureSchema(context.env.DB);

  const row = await context.env.DB.prepare(`
    SELECT storage_type, object_key FROM attendance_archive_files WHERE id = ?
  `).bind(context.params.id).first();
  if (!row) return json({ error: "보관 파일을 찾지 못했습니다." }, 404);

  if (row.storage_type === "r2" && row.object_key && context.env.FILES) {
    await context.env.FILES.delete(row.object_key).catch(() => {});
  }
  await context.env.DB.batch([
    context.env.DB.prepare("DELETE FROM attendance_archive_file_chunks WHERE file_id = ?").bind(context.params.id),
    context.env.DB.prepare("DELETE FROM attendance_archive_files WHERE id = ?").bind(context.params.id),
  ]);
  return json({ ok: true });
}

function toUint8Array(value) {
  if (value instanceof ArrayBuffer) return new Uint8Array(value);
  if (ArrayBuffer.isView(value)) return new Uint8Array(value.buffer, value.byteOffset, value.byteLength);
  if (Array.isArray(value)) return new Uint8Array(value);
  return new Uint8Array();
}

function contentDisposition(fileName) {
  const original = String(fileName || "download.xlsx").replace(/[\r\n]/g, "_");
  const extension = original.toLowerCase().endsWith(".xls") ? ".xls" : original.toLowerCase().endsWith(".xlsb") ? ".xlsb" : ".xlsx";
  return `attachment; filename="attendance-file${extension}"; filename*=UTF-8''${encodeURIComponent(original)}`;
}
