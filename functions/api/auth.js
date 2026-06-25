import { clearSessionCookie, createSessionCookie, isAuthenticated, json } from "../_lib/auth.js";

export async function onRequestGet(context) {
  const configured = Boolean(context.env.ADMIN_PASSWORD && context.env.SESSION_SECRET);
  const loggedIn = configured ? await isAuthenticated(context.request, context.env.SESSION_SECRET) : false;
  return json({ configured, loggedIn });
}

export async function onRequestPost(context) {
  if (!context.env.ADMIN_PASSWORD || !context.env.SESSION_SECRET) return json({ error: "ADMIN_PASSWORD와 SESSION_SECRET을 먼저 설정해 주세요." }, 503);
  const body = await context.request.json().catch(() => ({}));
  if (!body.password || body.password !== context.env.ADMIN_PASSWORD) return json({ error: "비밀번호가 일치하지 않습니다." }, 401);
  return json({ ok: true }, 200, { "set-cookie": await createSessionCookie(context.env.SESSION_SECRET) });
}

export async function onRequestDelete() {
  return json({ ok: true }, 200, { "set-cookie": clearSessionCookie() });
}
