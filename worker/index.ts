interface Env {
  DB: D1Database;
  EXPLEE_API_KEY: string;
  MICROSOFT_CLIENT_ID: string;
  MICROSOFT_TENANT_ID: string;
  ALLOWED_ORIGIN: string;
}

type JwtHeader = { alg: string; kid: string };
type JwtClaims = { aud: string | string[]; exp: number; iss: string; oid?: string; sub: string; tid?: string; preferred_username?: string };
type Jwk = JsonWebKey & { kid: string };
type Lead = Record<string, unknown> & { person_id: string; campaign_id: number; became_hot_at?: string | null };

const EXPLEE_BASE = 'https://api.explee.com/public/api/v1';
const PROJECT_ID = 34448;
let cachedKeys: { expires: number; keys: Jwk[] } | null = null;

function base64UrlBytes(value: string) {
  const normalized = value.replace(/-/g, '+').replace(/_/g, '/').padEnd(Math.ceil(value.length / 4) * 4, '=');
  return Uint8Array.from(atob(normalized), (char) => char.charCodeAt(0));
}

function base64UrlJson<T>(value: string): T {
  return JSON.parse(new TextDecoder().decode(base64UrlBytes(value))) as T;
}

async function microsoftKeys(env: Env) {
  if (cachedKeys && cachedKeys.expires > Date.now()) return cachedKeys.keys;
  const response = await fetch(`https://login.microsoftonline.com/${env.MICROSOFT_TENANT_ID}/discovery/v2.0/keys`);
  if (!response.ok) throw new Error('Microsoft identity keys are unavailable.');
  const payload = await response.json<{ keys: Jwk[] }>();
  cachedKeys = { expires: Date.now() + 60 * 60 * 1000, keys: payload.keys };
  return payload.keys;
}

async function verifyUser(request: Request, env: Env) {
  const token = request.headers.get('Authorization')?.replace(/^Bearer\s+/i, '');
  if (!token) throw new Response('Sign in required.', { status: 401 });
  const parts = token.split('.');
  if (parts.length !== 3) throw new Response('Invalid sign-in token.', { status: 401 });
  const header = base64UrlJson<JwtHeader>(parts[0]);
  const claims = base64UrlJson<JwtClaims>(parts[1]);
  const expectedIssuer = `https://login.microsoftonline.com/${env.MICROSOFT_TENANT_ID}/v2.0`;
  const audienceOk = Array.isArray(claims.aud) ? claims.aud.includes(env.MICROSOFT_CLIENT_ID) : claims.aud === env.MICROSOFT_CLIENT_ID;
  if (header.alg !== 'RS256' || !audienceOk || claims.exp * 1000 <= Date.now() || claims.iss !== expectedIssuer) throw new Response('Sign-in token is not valid for this app.', { status: 401 });
  const jwk = (await microsoftKeys(env)).find((item) => item.kid === header.kid);
  if (!jwk) throw new Response('Sign-in key was not found.', { status: 401 });
  const key = await crypto.subtle.importKey('jwk', jwk, { name: 'RSASSA-PKCS1-v1_5', hash: 'SHA-256' }, false, ['verify']);
  const valid = await crypto.subtle.verify({ name: 'RSASSA-PKCS1-v1_5' }, key, base64UrlBytes(parts[2]), new TextEncoder().encode(`${parts[0]}.${parts[1]}`));
  if (!valid) throw new Response('Sign-in signature is invalid.', { status: 401 });
  return { id: claims.oid || claims.sub, email: claims.preferred_username || '' };
}

function cors(env: Env) {
  return { 'Access-Control-Allow-Origin': env.ALLOWED_ORIGIN, 'Access-Control-Allow-Headers': 'Authorization, Content-Type', 'Access-Control-Allow-Methods': 'GET, PUT, POST, OPTIONS', Vary: 'Origin' };
}

function json(data: unknown, env: Env, status = 200) {
  return Response.json(data, { status, headers: cors(env) });
}

async function explee<T>(env: Env, path: string): Promise<T> {
  const response = await fetch(`${EXPLEE_BASE}${path}`, { headers: { 'X-API-Key': env.EXPLEE_API_KEY } });
  if (!response.ok) throw new Error(`Explee request failed (${response.status}).`);
  return response.json<T>();
}

async function syncCampaign(env: Env, campaignId: number) {
  const key = `campaign:${campaignId}`;
  const state = await env.DB.prepare('SELECT cursor FROM sync_state WHERE key = ?').bind(key).first<{ cursor: string }>();
  let offset = 0, newest = state?.cursor || '';
  let hasMore = true;
  while (hasMore) {
    const query = new URLSearchParams({ campaign_id: String(campaignId), limit: '200', offset: String(offset) });
    if (state?.cursor) query.set('since', state.cursor);
    const page = await explee<{ leads: Lead[]; has_more: boolean; next_offset?: number | null }>(env, `/autogtm/hot-leads?${query}`);
    for (const lead of page.leads) {
      if (lead.became_hot_at && (!newest || lead.became_hot_at > newest)) newest = lead.became_hot_at;
      await env.DB.prepare(`INSERT INTO leads (campaign_id, person_id, name, email, phone, job_title, company_name, company_domain, linkedin_url, country, why_hot, became_hot_at, raw_json, updated_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, datetime('now'))
        ON CONFLICT(campaign_id, person_id) DO UPDATE SET name=excluded.name,email=excluded.email,phone=excluded.phone,job_title=excluded.job_title,company_name=excluded.company_name,company_domain=excluded.company_domain,linkedin_url=excluded.linkedin_url,country=excluded.country,why_hot=excluded.why_hot,became_hot_at=excluded.became_hot_at,raw_json=excluded.raw_json,updated_at=datetime('now')`)
        .bind(lead.campaign_id, lead.person_id, lead.name ?? null, lead.email ?? null, lead.phone ?? null, lead.job_title ?? null, lead.company_name ?? null, lead.company_domain ?? null, lead.linkedin_url ?? null, lead.country ?? null, lead.why_hot ?? null, lead.became_hot_at ?? null, JSON.stringify(lead)).run();
    }
    hasMore = page.has_more && page.next_offset != null;
    if (hasMore) offset = page.next_offset!;
  }
  if (newest) await env.DB.prepare(`INSERT INTO sync_state (key, cursor, updated_at) VALUES (?, ?, datetime('now')) ON CONFLICT(key) DO UPDATE SET cursor=excluded.cursor,updated_at=datetime('now')`).bind(key, newest).run();
}

async function syncAll(env: Env) {
  const result = await explee<{ campaigns: Array<{ id: number }> }>(env, `/autogtm/campaigns?project_id=${PROJECT_ID}`);
  for (const campaign of result.campaigns) await syncCampaign(env, campaign.id);
  return { campaigns: result.campaigns.length, synced_at: new Date().toISOString() };
}

async function handle(request: Request, env: Env) {
  if (request.method === 'OPTIONS') return new Response(null, { status: 204, headers: cors(env) });
  if (request.headers.get('Origin') && request.headers.get('Origin') !== env.ALLOWED_ORIGIN) return json({ error: 'Origin not allowed.' }, env, 403);
  const user = await verifyUser(request, env);
  const url = new URL(request.url);
  if (url.pathname === '/api/leads' && request.method === 'GET') {
    const { results } = await env.DB.prepare(`SELECT campaign_id,person_id,name,email,phone,job_title,company_name,company_domain,linkedin_url,country,why_hot,became_hot_at,status FROM leads ORDER BY became_hot_at DESC`).all();
    return json({ leads: results }, env);
  }
  if (url.pathname === '/api/settings' && request.method === 'GET') {
    const row = await env.DB.prepare('SELECT instruction, example, signature FROM settings WHERE user_id = ?').bind(user.id).first();
    return json(row || { instruction: 'Be warm, direct, and useful. Keep replies under 100 words.', example: '', signature: user.email.split('@')[0] || '' }, env);
  }
  if (url.pathname === '/api/settings' && request.method === 'PUT') {
    const body = await request.json<{ instruction?: string; example?: string; signature?: string }>();
    const instruction = String(body.instruction || '').slice(0, 4000), example = String(body.example || '').slice(0, 12000), signature = String(body.signature || '').slice(0, 200);
    await env.DB.prepare(`INSERT INTO settings (user_id,instruction,example,signature,updated_at) VALUES (?,?,?,?,datetime('now')) ON CONFLICT(user_id) DO UPDATE SET instruction=excluded.instruction,example=excluded.example,signature=excluded.signature,updated_at=datetime('now')`).bind(user.id, instruction, example, signature).run();
    return json({ saved: true }, env);
  }
  if (url.pathname === '/api/sync' && request.method === 'POST') return json(await syncAll(env), env);
  const status = url.pathname.match(/^\/api\/leads\/(\d+)\/([^/]+)\/status$/);
  if (status && request.method === 'PUT') {
    const body = await request.json<{ status?: string }>();
    const next = ['New', 'Drafted', 'Sent'].includes(body.status || '') ? body.status : null;
    if (!next) return json({ error: 'Invalid status.' }, env, 422);
    await env.DB.prepare("UPDATE leads SET status = ?, updated_at = datetime('now') WHERE campaign_id = ? AND person_id = ?").bind(next, Number(status[1]), decodeURIComponent(status[2])).run();
    return json({ saved: true }, env);
  }
  const thread = url.pathname.match(/^\/api\/leads\/(\d+)\/([^/]+)\/thread$/);
  if (thread && request.method === 'GET') return json(await explee(env, `/autogtm/campaigns/${thread[1]}/inbox/${encodeURIComponent(thread[2])}`), env);
  return json({ error: 'Not found.' }, env, 404);
}

const worker = {
  fetch(request: Request, env: Env) { return handle(request, env).catch((error) => error instanceof Response ? new Response(error.body, { status: error.status, headers: cors(env) }) : json({ error: error instanceof Error ? error.message : 'Unexpected error.' }, env, 500)); },
  scheduled(_controller: ScheduledController, env: Env, ctx: ExecutionContext) { ctx.waitUntil(syncAll(env)); },
};

export default worker;
