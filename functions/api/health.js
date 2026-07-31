export async function onRequest(context) {
  return json({
    ok: Boolean(context.env.DB),
    storage: context.env.DB ? "d1" : "browser",
    version: "clean-v4",
  });
}

function json(data, init = {}) {
  return new Response(JSON.stringify(data), {
    ...init,
    headers: {
      "content-type": "application/json; charset=utf-8",
      ...(init.headers || {}),
    },
  });
}
