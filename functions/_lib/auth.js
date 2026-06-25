const COOKIE_NAME = "attendance_session";
const MAX_AGE = 60 * 60 * 12;

function base64url(input) {
  const bytes = input instanceof Uint8Array ? input : new TextEncoder().encode(input);
  let binary = "";
  bytes.forEach((byte) => { binary += String.fromCharCode(byte); });
  return btoa(binary).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/g, "");
}

async function sign(value, secret) {
  const key = await crypto.subtle.importKey("raw", new TextEncoder().encode(secret), { name: "HMAC", hash: "SHA-256" }, false, ["sign"]);
  return base64url(new Uint8Array(await crypto.subtle.sign("HMAC", key, new TextEncoder().encode(value))));
}

export async function createSessionCookie(secret) {
  const payload = base64url(JSON.stringify({ exp: Math.floor(Date.now() / 1000) + MAX_AGE }));
  const signature = await sign(payload, secret);
  return `${COOKIE_NAME}=${payload}.${signature}; Path=/; HttpOnly; Secure; SameSite=Strict; Max-Age=${MAX_AGE}`;
}

export function clearSessionCookie() {
  return `${COOKIE_NAME}=; Path=/; HttpOnly; Secure; SameSite=Strict; Max-Age=0`;
}

export async function isAuthenticated(request, secret) {
  if (!secret) return false;
  const cookie = request.headers.get("cookie") || "";
  const token = cookie.split(";").map((item) => item.trim()).find((item) => item.startsWith(`${COOKIE_NAME}=`))?.split("=").slice(1).join("=");
  if (!token) return false;
  const [payload, signature] = token.split(".");
  if (!payload || !signature || (await sign(payload, secret)) !== signature) return false;
  try {
    const padded = payload.replace(/-/g, "+").replace(/_/g, "/") + "===".slice((payload.length + 3) % 4);
    const data = JSON.parse(atob(padded));
    return Number(data.exp) > Math.floor(Date.now() / 1000);
  } catch { return false; }
}

export async function requireAuth(context) {
  if (!context.env.ADMIN_PASSWORD || !context.env.SESSION_SECRET) return json({ error: "관리자 환경 변수가 설정되지 않았습니다." }, 503);
  if (!(await isAuthenticated(context.request, context.env.SESSION_SECRET))) return json({ error: "로그인이 필요합니다." }, 401);
  return null;
}

export function json(data, status = 200, headers = {}) {
  return new Response(JSON.stringify(data), { status, headers: { "content-type": "application/json; charset=utf-8", ...headers } });
}
