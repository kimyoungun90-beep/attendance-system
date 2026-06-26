import { json, requireAuth } from "../_lib/auth.js";
import { ensureSchema } from "../_lib/schema.js";

const VALID_ROUTES = new Set(["homeplus", "electroland"]);
const VALID_KINDS = new Set(["plan", "attendance", "result", "other"]);
const D1_FILE_LIMIT = 20 * 1024 * 1024;
const D1_CHUNK_SIZE = 450_000;

export async function onRequestGet(context) {
  const denied = await requireAuth(context);
  if (denied) return denied;
  await ensureSchema(context.env.DB);

  const url = new URL(context.request.url);
  const route = url.searchParams.get("route") || "";
  const month = url.searchParams.get("month") || "";
  const conditions = [];
  const bindings = [];
  if (route) {
    if (!VALID_ROUTES.has(route)) return json({ error: "경로 구분이 올바르지 않습니다." }, 400);
    conditions.push("route = ?");
    bindings.push(route);
  }
  if (month) {
    if (!/^\d{4}-\d{2}$/.test(month)) return json({ error: "대상 월이 올바르지 않습니다." }, 400);
    conditions.push("month = ?");
    bindings.push(month);
  }
  const where = conditions.length ? `WHERE ${conditions.join(" AND ")}` : "";
  let statement = context.env.DB.prepare(`
    SELECT id, route, month, file_kind, file_name, content_type, size_bytes,
           storage_type, source_type, closure_id, note, created_at
    FROM attendance_archive_files
    ${where}
    ORDER BY month DESC, route ASC, created_at DESC
    LIMIT 500
  `);
  if (bindings.length) statement = statement.bind(...bindings);
  const result = await statement.all();

  return json({
    items: result.results || [],
    r2Configured: Boolean(context.env.FILES),
    d1FileLimit: D1_FILE_LIMIT,
  });
}

export async function onRequestPost(context) {
  const denied = await requireAuth(context);
  if (denied) return denied;
  await ensureSchema(context.env.DB);

  const form = await context.request.formData().catch(() => null);
  const file = form?.get("file");
  const route = String(form?.get("route") || "");
  const month = String(form?.get("month") || "");
  const fileKind = String(form?.get("fileKind") || "other");
  const note = String(form?.get("note") || "").trim();
  const sourceType = form?.get("sourceType") === "closure" ? "closure" : "manual";
  const closureId = String(form?.get("closureId") || "").trim();
  const replace = String(form?.get("replace") || "") === "true";

  if (!(file instanceof File) || file.size <= 0) return json({ error: "보관할 엑셀 파일을 선택해 주세요." }, 400);
  if (!VALID_ROUTES.has(route) || !/^\d{4}-\d{2}$/.test(month) || !VALID_KINDS.has(fileKind)) {
    return json({ error: "경로·대상 월·파일 구분을 확인해 주세요." }, 400);
  }
  if (!/\.(xlsx|xls|xlsb)$/i.test(file.name)) return json({ error: "엑셀 파일(xlsx, xls, xlsb)만 보관할 수 있습니다." }, 400);
  if (!context.env.FILES && file.size > D1_FILE_LIMIT) {
    return json({
      error: `현재 D1 분할 보관 한도는 파일당 20MB입니다. 이 파일은 ${(file.size / 1024 / 1024).toFixed(1)}MB이므로 파일을 줄이거나 R2 바인딩(FILES)을 연결해 주세요.`,
      code: "R2_REQUIRED",
    }, 413);
  }

  const oldClosureFiles = replace && sourceType === "closure"
    ? await findMatchingClosureFiles(context, route, month, fileKind)
    : [];

  const id = crypto.randomUUID();
  const contentType = file.type || "application/octet-stream";
  let storageType = "d1";
  let objectKey = null;

  try {
    if (context.env.FILES) {
      storageType = "r2";
      objectKey = `${route}/${month}/${fileKind}/${id}-${sanitizeFileName(file.name)}`;
      await context.env.FILES.put(objectKey, file.stream(), {
        httpMetadata: { contentType },
        customMetadata: { originalName: file.name, route, month, fileKind },
      });
    }

    await context.env.DB.prepare(`
      INSERT INTO attendance_archive_files
      (id, route, month, file_kind, file_name, content_type, size_bytes,
       storage_type, object_key, file_blob, source_type, closure_id, note)
      VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, NULL, ?, ?, ?)
    `).bind(
      id, route, month, fileKind, file.name, contentType, file.size,
      storageType, objectKey, sourceType, closureId || null, note
    ).run();

    if (storageType === "d1") {
      const buffer = await file.arrayBuffer();
      const statements = [];
      for (let offset = 0, chunkIndex = 0; offset < buffer.byteLength; offset += D1_CHUNK_SIZE, chunkIndex += 1) {
        const chunk = buffer.slice(offset, Math.min(offset + D1_CHUNK_SIZE, buffer.byteLength));
        statements.push(context.env.DB.prepare(`
          INSERT INTO attendance_archive_file_chunks (file_id, chunk_index, chunk_blob)
          VALUES (?, ?, ?)
        `).bind(id, chunkIndex, chunk));
      }
      await runBatch(context.env.DB, statements);
    }
    await deleteArchivedFiles(context, oldClosureFiles);
  } catch (error) {
    if (storageType === "r2" && objectKey && context.env.FILES) await context.env.FILES.delete(objectKey).catch(() => {});
    await context.env.DB.prepare("DELETE FROM attendance_archive_file_chunks WHERE file_id = ?").bind(id).run().catch(() => {});
    await context.env.DB.prepare("DELETE FROM attendance_archive_files WHERE id = ?").bind(id).run().catch(() => {});
    return json({ error: `파일 보관 중 오류가 발생했습니다: ${error.message || "알 수 없는 오류"}` }, 500);
  }

  return json({ ok: true, id, storageType, r2Configured: Boolean(context.env.FILES) }, 201);
}

async function findMatchingClosureFiles(context, route, month, fileKind) {
  const old = await context.env.DB.prepare(`
    SELECT id, storage_type, object_key FROM attendance_archive_files
    WHERE route = ? AND month = ? AND file_kind = ? AND source_type = 'closure'
  `).bind(route, month, fileKind).all();
  return old.results || [];
}

async function deleteArchivedFiles(context, items) {
  for (const item of items || []) {
    if (item.storage_type === "r2" && item.object_key && context.env.FILES) {
      await context.env.FILES.delete(item.object_key).catch(() => {});
    }
    await context.env.DB.batch([
      context.env.DB.prepare("DELETE FROM attendance_archive_file_chunks WHERE file_id = ?").bind(item.id),
      context.env.DB.prepare("DELETE FROM attendance_archive_files WHERE id = ?").bind(item.id),
    ]);
  }
}

async function runBatch(db, statements) {
  for (let index = 0; index < statements.length; index += 40) {
    await db.batch(statements.slice(index, index + 40));
  }
}

function sanitizeFileName(name) {
  return String(name || "file.xlsx").replace(/[^0-9A-Za-z가-힣._-]+/g, "_").slice(-120);
}
