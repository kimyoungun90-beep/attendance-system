export async function onRequest(context) {
  if (context.request.method === "OPTIONS") return new Response(null, { status: 204, headers: corsHeaders() });
  if (!context.env.DB) return json({ error: "D1 DB binding is not configured." }, { status: 503 });
  await ensureSchema(context.env.DB);
  if (context.request.method === "GET") return handleGet(context);
  if (context.request.method === "POST") return handlePost(context);
  return json({ error: "Method not allowed." }, { status: 405 });
}

async function handleGet(context) {
  const url = new URL(context.request.url);
  const key = url.searchParams.get("key");
  if (!key) return json({ error: "key is required." }, { status: 400 });
  const row = await context.env.DB.prepare(`
    SELECT key, type, route, month, payload_json, updated_at
    FROM attendance_clean_store
    WHERE key = ?
  `).bind(key).first();
  if (!row) return json({ payload: null });
  return json({
    key: row.key,
    type: row.type,
    route: row.route,
    month: row.month,
    updatedAt: row.updated_at,
    payload: parseJson(row.payload_json),
  });
}

async function handlePost(context) {
  const body = await context.request.json().catch(() => null);
  if (!body?.key || !body?.type) return json({ error: "key and type are required." }, { status: 400 });
  const payloadJson = JSON.stringify(body.payload ?? {});
  await context.env.DB.prepare(`
    INSERT INTO attendance_clean_store (key, type, route, month, payload_json, updated_at)
    VALUES (?, ?, ?, ?, ?, datetime('now'))
    ON CONFLICT(key) DO UPDATE SET
      type = excluded.type,
      route = excluded.route,
      month = excluded.month,
      payload_json = excluded.payload_json,
      updated_at = datetime('now')
  `).bind(
    String(body.key),
    String(body.type),
    String(body.route || ""),
    String(body.month || ""),
    payloadJson,
  ).run();
  return json({ ok: true });
}

async function ensureSchema(db) {
  await db.prepare(`
    CREATE TABLE IF NOT EXISTS attendance_clean_store (
      key TEXT PRIMARY KEY,
      type TEXT NOT NULL,
      route TEXT,
      month TEXT,
      payload_json TEXT NOT NULL DEFAULT '{}',
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      updated_at TEXT NOT NULL DEFAULT (datetime('now'))
    )
  `).run();
}

function parseJson(value) {
  try {
    return JSON.parse(value || "{}");
  } catch {
    return {};
  }
}

function json(data, init = {}) {
  return new Response(JSON.stringify(data), {
    ...init,
    headers: {
      "content-type": "application/json; charset=utf-8",
      ...corsHeaders(),
      ...(init.headers || {}),
    },
  });
}

function corsHeaders() {
  return {
    "access-control-allow-origin": "*",
    "access-control-allow-methods": "GET, POST, OPTIONS",
    "access-control-allow-headers": "content-type",
  };
}
