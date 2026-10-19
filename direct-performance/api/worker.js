/**
 * 飞书多维表格 → 销售直推业绩看板 数据代理
 * 运行环境：Cloudflare Worker（单文件，无依赖）
 *
 * 需要配置的环境变量（在 Cloudflare 控制台 Settings → Variables 中添加）：
 *   FEISHU_APP_ID        飞书自建应用 App ID
 *   FEISHU_APP_SECRET    飞书自建应用 App Secret
 *   BITABLE_APP_TOKEN    多维表格 app_token（表格 URL 中 /base/ 后面那段）
 *   ORDERS_TABLE_ID      订单池表 table_id
 *   STAFF_TABLE_ID       销售名单表 table_id
 *   LOGIN_USER           看板登录账号
 *   LOGIN_PASSWORD       看板登录密码
 *   SESSION_SECRET       随机长字符串，用于签发登录令牌
 *   ALLOWED_ORIGIN       允许的前端来源，如 https://1echo2026.github.io
 */

let tokenCache = { token: null, exp: 0 };

const enc = new TextEncoder();

/* ---------------- 工具 ---------------- */
function json(obj, status = 200, extra = {}) {
  return new Response(JSON.stringify(obj), {
    status,
    headers: { "Content-Type": "application/json; charset=utf-8", ...extra },
  });
}

function b64url(bytes) {
  let s = btoa(String.fromCharCode(...new Uint8Array(bytes)));
  return s.replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
}
function b64urlDecode(str) {
  str = String(str).replace(/-/g, "+").replace(/_/g, "/");
  while (str.length % 4) str += "=";
  const bin = atob(str);
  const bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  return bytes;
}

async function hmac(secret, data) {
  const key = await crypto.subtle.importKey(
    "raw", enc.encode(secret), { name: "HMAC", hash: "SHA-256" }, false, ["sign"]
  );
  return new Uint8Array(await crypto.subtle.sign("HMAC", key, enc.encode(data)));
}

async function signToken(payload, secret) {
  const body = b64url(enc.encode(JSON.stringify(payload)));
  const sig = b64url(await hmac(secret, body));
  return body + "." + sig;
}

async function verifyToken(token, secret) {
  if (!token || token.indexOf(".") < 0) return null;
  const parts = token.split(".");
  if (parts.length !== 2) return null;
  const [body, sig] = parts;
  const expect = b64url(await hmac(secret, body));
  if (sig.length !== expect.length) return null;
  let diff = 0;
  for (let i = 0; i < sig.length; i++) diff |= sig.charCodeAt(i) ^ expect.charCodeAt(i);
  if (diff !== 0) return null;
  try {
    const payload = JSON.parse(new TextDecoder().decode(b64urlDecode(body)));
    if (!payload.exp || payload.exp < Date.now()) return null;
    return payload;
  } catch (e) {
    return null;
  }
}

function corsHeaders(env, origin) {
  const allowed = String(env.ALLOWED_ORIGIN || "").split(",").map(s => s.trim()).filter(Boolean);
  let allow = "*";
  if (allowed.length) allow = allowed.includes(origin) ? origin : allowed[0];
  return {
    "Access-Control-Allow-Origin": allow,
    "Access-Control-Allow-Methods": "GET,POST,OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type,Authorization",
    "Access-Control-Max-Age": "86400",
    "Vary": "Origin",
  };
}

/* ---------------- 飞书 ---------------- */
async function getTenantToken(env) {
  const now = Date.now();
  if (tokenCache.token && tokenCache.exp > now + 60000) return tokenCache.token;

  const res = await fetch("https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal", {
    method: "POST",
    headers: { "Content-Type": "application/json; charset=utf-8" },
    body: JSON.stringify({ app_id: env.FEISHU_APP_ID, app_secret: env.FEISHU_APP_SECRET }),
  });
  const data = await res.json();
  if (data.code !== 0) throw new Error("飞书鉴权失败 " + data.code + ": " + data.msg);
  tokenCache = { token: data.tenant_access_token, exp: now + (data.expire || 7200) * 1000 };
  return tokenCache.token;
}

function normalizeValue(v) {
  if (v === null || v === undefined) return "";
  if (Array.isArray(v)) {
    return v.map(x => {
      if (x && typeof x === "object") return x.text ?? x.name ?? x.id ?? x.link ?? "";
      return x;
    }).join("");
  }
  if (typeof v === "object") return v.text ?? v.name ?? v.link ?? "";
  return v;
}

async function fetchTable(env, tableId) {
  if (!tableId) return [];
  const token = await getTenantToken(env);
  const base = "https://open.feishu.cn/open-apis/bitable/v1/apps/" +
    encodeURIComponent(env.BITABLE_APP_TOKEN) + "/tables/" + encodeURIComponent(tableId) + "/records";

  let items = [];
  let pageToken = "";
  for (let guard = 0; guard < 50; guard++) {
    const u = new URL(base);
    u.searchParams.set("page_size", "500");
    if (pageToken) u.searchParams.set("page_token", pageToken);
    const res = await fetch(u.toString(), {
      headers: { Authorization: "Bearer " + token },
    });
    const data = await res.json();
    if (data.code !== 0) throw new Error("飞书接口错误 " + data.code + ": " + data.msg);
    const page = data.data || {};
    items = items.concat(page.items || []);
    if (!page.has_more) break;
    pageToken = page.page_token || "";
    if (!pageToken) break;
  }

  return items.map(it => {
    const row = {};
    const fields = it.fields || {};
    for (const k of Object.keys(fields)) row[k] = normalizeValue(fields[k]);
    return row;
  });
}

/* ---------------- 路由 ---------------- */
async function handleLogin(request, env, cors) {
  let body = {};
  try { body = await request.json(); } catch (e) {}
  const u = String(body.username || "");
  const p = String(body.password || "");
  if (!env.LOGIN_USER || !env.LOGIN_PASSWORD) {
    return json({ error: "服务端未配置登录账号" }, 500, cors);
  }
  if (u !== String(env.LOGIN_USER) || p !== String(env.LOGIN_PASSWORD)) {
    return json({ error: "账号或密码错误" }, 401, cors);
  }
  const token = await signToken(
    { u, exp: Date.now() + 8 * 3600 * 1000 },
    env.SESSION_SECRET || "change-me"
  );
  return json({ token }, 200, cors);
}

async function handleData(request, env, cors) {
  const authz = request.headers.get("Authorization") || "";
  const token = authz.startsWith("Bearer ") ? authz.slice(7) : "";
  const payload = await verifyToken(token, env.SESSION_SECRET || "change-me");
  if (!payload) return json({ error: "未登录或登录已过期" }, 401, cors);

  const [orders, staff] = await Promise.all([
    fetchTable(env, env.ORDERS_TABLE_ID),
    fetchTable(env, env.STAFF_TABLE_ID),
  ]);
  return json({ orders, staff, syncedAt: new Date().toISOString() }, 200, cors);
}

export default {
  async fetch(request, env) {
    const url = new URL(request.url);
    const origin = request.headers.get("Origin") || "";
    const cors = corsHeaders(env, origin);

    if (request.method === "OPTIONS") {
      return new Response(null, { status: 204, headers: cors });
    }

    try {
      if (url.pathname === "/api/login" && request.method === "POST") {
        return await handleLogin(request, env, cors);
      }
      if (url.pathname === "/api/data" && request.method === "GET") {
        return await handleData(request, env, cors);
      }
      if (url.pathname === "/api/health") {
        return json({ ok: true, time: new Date().toISOString() }, 200, cors);
      }
      return json({ error: "not found" }, 404, cors);
    } catch (e) {
      return json({ error: String((e && e.message) || e) }, 500, cors);
    }
  },
};
