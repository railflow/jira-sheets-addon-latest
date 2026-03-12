/**
 * SheetSync Admin Panel — admin.sheetsync.dev
 * Cloudflare Worker: serves a self-contained SPA admin dashboard
 *
 * Secrets to set via `wrangler secret put`:
 *   ADMIN_USERNAME  (default: admin)
 *   ADMIN_PASSWORD  (default: M697~cUzqmQF)
 */

const MAX_ATTEMPTS = 5;
const LOCKOUT_MS = 3 * 60 * 1000; // 3 minutes

const PRICE_MAP = {
  'price_1T6S0JDhvP6DurKScZ69xlBd': { plan: 'Starter',      billing: 'Monthly', display: '$1/mo' },
  'price_1T334WDhvP6DurKS3ozSbzwW': { plan: 'Starter',      billing: 'Yearly',  display: '$9/yr' },
  'price_1T6PtuDhvP6DurKSeJveH6Jc': { plan: 'Professional', billing: 'Monthly', display: '$9/mo' },
  'price_1T6PtuDhvP6DurKSHIFeZaFV': { plan: 'Professional', billing: 'Yearly',  display: '$75/yr' },
  'price_1T6PvIDhvP6DurKSCrisUWHI': { plan: 'Enterprise',   billing: 'Monthly', display: '$345/mo' },
  'price_1T6PvVDhvP6DurKSHcVjbbC2': { plan: 'Enterprise',   billing: 'Yearly',  display: '$2,900/yr' },
};

// ─── Helpers ────────────────────────────────────────────────────────────────

function getCookie(request, name) {
  const header = request.headers.get('Cookie') || '';
  for (const part of header.split(';')) {
    const [k, ...v] = part.trim().split('=');
    if (k.trim() === name) return decodeURIComponent(v.join('='));
  }
  return null;
}

function getClientIP(request) {
  return request.headers.get('CF-Connecting-IP') ||
         (request.headers.get('X-Forwarded-For') || '').split(',')[0].trim() ||
         'unknown';
}

function generateToken() {
  const arr = new Uint8Array(32);
  crypto.getRandomValues(arr);
  return Array.from(arr).map(b => b.toString(16).padStart(2, '0')).join('');
}

function json(data, status = 200, extraHeaders = {}) {
  return new Response(JSON.stringify(data), {
    status,
    headers: { 'Content-Type': 'application/json', ...extraHeaders },
  });
}

// ─── Session ────────────────────────────────────────────────────────────────

async function getSession(request, env) {
  const token = getCookie(request, 'adm_sess');
  if (!token) return null;
  const raw = await env.LICENSES.get(`admin_session:${token}`);
  if (!raw) return null;
  const sess = JSON.parse(raw);
  if (sess.expiresAt && Date.now() > sess.expiresAt) {
    await env.LICENSES.delete(`admin_session:${token}`);
    return null;
  }
  return { token, ...sess };
}

// ─── Route Handlers ──────────────────────────────────────────────────────────

async function handleLogin(request, env) {
  const ip = getClientIP(request);
  const rlKey = `admin_ratelimit:${ip}`;

  const rlRaw = await env.LICENSES.get(rlKey);
  const rl = rlRaw ? JSON.parse(rlRaw) : { attempts: 0 };

  if (rl.lockedUntil && Date.now() < rl.lockedUntil) {
    const retryAfter = Math.ceil((rl.lockedUntil - Date.now()) / 1000);
    return json({ error: 'Too many failed attempts. Try again later.', retryAfter }, 429);
  }

  let body;
  try { body = await request.json(); } catch { return json({ error: 'Bad request' }, 400); }
  const { username, password, remember } = body;

  const expectedUser = env.ADMIN_USERNAME || 'admin';
  const expectedPass = env.ADMIN_PASSWORD || 'M697~cUzqmQF';

  if (username !== expectedUser || password !== expectedPass) {
    rl.attempts = (rl.attempts || 0) + 1;
    if (rl.attempts >= MAX_ATTEMPTS) {
      rl.lockedUntil = Date.now() + LOCKOUT_MS;
    }
    await env.LICENSES.put(rlKey, JSON.stringify(rl), { expirationTtl: 600 });
    return json({
      error: 'Invalid credentials.',
      attemptsRemaining: Math.max(0, MAX_ATTEMPTS - rl.attempts),
      locked: rl.lockedUntil ? true : false,
      retryAfter: rl.lockedUntil ? Math.ceil((rl.lockedUntil - Date.now()) / 1000) : null,
    }, 401);
  }

  // Success — clear rate limit
  await env.LICENSES.delete(rlKey);

  const token = generateToken();
  const days = remember ? 30 : 1;
  const expiresAt = Date.now() + days * 86400000;

  await env.LICENSES.put(
    `admin_session:${token}`,
    JSON.stringify({ createdAt: Date.now(), expiresAt, remember: !!remember }),
    { expirationTtl: days * 86400 }
  );

  const maxAge = remember ? `; Max-Age=${days * 86400}` : '';
  return json({ ok: true }, 200, {
    'Set-Cookie': `adm_sess=${token}; Path=/; HttpOnly; Secure; SameSite=Strict${maxAge}`,
  });
}

async function handleLogout(request, env) {
  const token = getCookie(request, 'adm_sess');
  if (token) await env.LICENSES.delete(`admin_session:${token}`);
  return json({ ok: true }, 200, {
    'Set-Cookie': 'adm_sess=; Path=/; HttpOnly; Secure; SameSite=Strict; Max-Age=0',
  });
}

async function handleMe(request, env) {
  const sess = await getSession(request, env);
  if (!sess) return json({ error: 'Unauthorized' }, 401);
  return json({ ok: true, expiresAt: sess.expiresAt });
}

async function handleData(request, env) {
  const sess = await getSession(request, env);
  if (!sess) return json({ error: 'Unauthorized' }, 401);

  const [usersRes, domainsRes, loginsRes, salesRes] = await Promise.all([
    env.DB.prepare('SELECT * FROM user_licenses   ORDER BY last_updated DESC').all(),
    env.DB.prepare('SELECT * FROM domain_licenses ORDER BY last_updated DESC').all(),
    env.DB.prepare('SELECT * FROM login_events    ORDER BY last_seen DESC').all(),
    env.DB.prepare('SELECT * FROM sales           ORDER BY created_at DESC').all(),
  ]);

  const normalizePlan = plan => {
    if (!plan || plan === 'free') return 'Starter';
    const p = plan.toLowerCase();
    if (p.includes('enterprise')) return 'Enterprise';
    if (p.includes('pro'))        return 'Professional';
    if (p.includes('starter'))    return 'Starter';
    return plan;
  };

  const users = (usersRes.results || []).map(u => ({
    email:          u.email,
    domain:         u.domain,
    plan:           u.plan,
    planLabel:      normalizePlan(u.plan),
    status:         u.status,
    allowed:        !!u.allowed,
    customerId:     u.customer_id,
    subscriptionId: u.subscription_id,
    priceId:        u.price_id,
    amount:         u.amount,
    renewsAt:       u.renews_at,
    lastUpdated:    u.last_updated,
    priceInfo:      u.price_id ? PRICE_MAP[u.price_id] : null,
  }));

  const domains = (domainsRes.results || []).map(d => ({
    domain:         d.domain,
    email:          d.email,
    plan:           d.plan,
    planLabel:      normalizePlan(d.plan),
    status:         d.status,
    allowed:        !!d.allowed,
    customerId:     d.customer_id,
    subscriptionId: d.subscription_id,
    priceId:        d.price_id,
    amount:         d.amount,
    renewsAt:       d.renews_at,
    seats:          d.seats,
    lastUpdated:    d.last_updated,
    priceInfo:      d.price_id ? PRICE_MAP[d.price_id] : null,
    users:          users.filter(u => u.email.split('@')[1] === d.domain),
  }));

  const domainSet   = new Set(domains.map(d => d.domain));
  const individuals = users.filter(u => !domainSet.has(u.email.split('@')[1]));

  const logins = (loginsRes.results || []).map(l => ({
    email:     l.email,
    firstSeen: l.first_seen,
    lastSeen:  l.last_seen,
    count:     l.visit_count,
    plan:      l.plan,
  }));

  const sales = (salesRes.results || []).map(s => ({
    date:         s.created_at,
    email:        s.email,
    plan:         s.plan,
    domain:       s.domain,
    customer:     s.customer_id,
    subscription: s.subscription_id,
    amount:       s.amount || 0,
  }));

  const totalRevenue = sales.reduce((n, s) => n + s.amount, 0);
  const activeLicenses =
    users.filter(u => u.status === 'active' || u.status === 'trialing').length +
    domains.filter(d => d.status === 'active' || d.status === 'trialing').length;

  return json({
    stats: { totalDomains: domains.length, totalUsers: users.length, activeLicenses, totalRevenue },
    domains,
    individuals,
    logins,
    sales,
  });
}

// ─── SPA HTML ────────────────────────────────────────────────────────────────

function serveSPA() {
  const html = /* html */`<!DOCTYPE html>
<html lang="en" class="h-full">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>SheetSync Admin</title>
  <script src="https://cdn.tailwindcss.com"></script>
  <style>
    body { background: #030712; color: #f1f5f9; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif; }
    .badge { display: inline-flex; align-items: center; padding: 2px 8px; border-radius: 9999px; font-size: 11px; font-weight: 600; letter-spacing: .02em; }
    .badge-green  { background: #052e16; color: #4ade80; border: 1px solid #166534; }
    .badge-yellow { background: #422006; color: #fbbf24; border: 1px solid #92400e; }
    .badge-red    { background: #450a0a; color: #f87171; border: 1px solid #7f1d1d; }
    .badge-blue   { background: #0c1a3d; color: #60a5fa; border: 1px solid #1e40af; }
    .badge-gray   { background: #111827; color: #9ca3af; border: 1px solid #374151; }
    .sidebar-link { display: flex; align-items: center; gap: 10px; padding: 8px 12px; border-radius: 8px; font-size: 14px; font-weight: 500; color: #94a3b8; cursor: pointer; transition: all .15s; }
    .sidebar-link:hover { background: #111827; color: #f1f5f9; }
    .sidebar-link.active { background: #1e3a8a22; color: #60a5fa; }
    .stat-card { background: #0f172a; border: 1px solid #1e293b; border-radius: 12px; padding: 20px 24px; }
    .tbl { width: 100%; border-collapse: collapse; font-size: 13px; }
    .tbl th { text-align: left; padding: 10px 14px; color: #64748b; font-weight: 600; font-size: 11px; text-transform: uppercase; letter-spacing: .06em; border-bottom: 1px solid #1e293b; }
    .tbl td { padding: 11px 14px; border-bottom: 1px solid #0f172a; vertical-align: middle; }
    .tbl tr:hover td { background: #0f172a; }
    .domain-card { background: #0a1628; border: 1px solid #1e293b; border-radius: 12px; overflow: hidden; margin-bottom: 16px; }
    .domain-header { display: flex; align-items: center; justify-content: space-between; padding: 14px 18px; cursor: pointer; }
    .domain-header:hover { background: #0f172a; }
    .panel { background: #0f172a; border: 1px solid #1e293b; border-radius: 12px; overflow: hidden; }
    input[type=text], input[type=password], input[type=search] {
      background: #0f172a; border: 1px solid #1e293b; border-radius: 8px;
      color: #f1f5f9; padding: 8px 12px; font-size: 14px; outline: none; transition: border-color .15s;
    }
    input[type=text]:focus, input[type=password]:focus, input[type=search]:focus { border-color: #3b82f6; }
    .btn-primary { background: #2563eb; color: white; border: none; border-radius: 8px; padding: 10px 18px; font-size: 14px; font-weight: 600; cursor: pointer; transition: background .15s; }
    .btn-primary:hover:not(:disabled) { background: #1d4ed8; }
    .btn-primary:disabled { background: #1e293b; color: #475569; cursor: not-allowed; }
    .spinner { width: 20px; height: 20px; border: 2px solid #1e293b; border-top-color: #3b82f6; border-radius: 50%; animation: spin .6s linear infinite; }
    @keyframes spin { to { transform: rotate(360deg); } }
    ::-webkit-scrollbar { width: 6px; } ::-webkit-scrollbar-track { background: transparent; }
    ::-webkit-scrollbar-thumb { background: #1e293b; border-radius: 3px; }
    .collapsible { transition: max-height .25s ease; overflow: hidden; }
    select { background: #0f172a; border: 1px solid #1e293b; border-radius: 8px; color: #f1f5f9; padding: 8px 12px; font-size: 14px; outline: none; transition: border-color .15s; }
    select:focus { border-color: #3b82f6; }
    .drawer { position:fixed; top:0; right:0; height:100vh; width:380px; background:#080f1e; border-left:1px solid #1e293b; display:flex; flex-direction:column; transform:translateX(100%); transition:transform .25s cubic-bezier(.4,0,.2,1); z-index:50; }
    .drawer.open { transform:translateX(0); }
    .drawer-overlay { position:fixed; inset:0; background:rgba(0,0,0,.4); z-index:49; opacity:0; pointer-events:none; transition:opacity .25s; }
    .drawer-overlay.open { opacity:1; pointer-events:auto; }
    .btn-danger { background:transparent; color:#f87171; border:1px solid #7f1d1d; border-radius:6px; padding:5px 10px; font-size:12px; cursor:pointer; transition:all .15s; }
    .btn-danger:hover { background:#450a0a; }
    .btn-ghost { background:transparent; color:#94a3b8; border:1px solid #1e293b; border-radius:6px; padding:5px 10px; font-size:12px; cursor:pointer; transition:all .15s; }
    .btn-ghost:hover { background:#1e293b; color:#f1f5f9; }
    .form-row { display:flex; flex-direction:column; gap:6px; }
    .form-label { font-size:11px; font-weight:600; color:#64748b; text-transform:uppercase; letter-spacing:.06em; }
    .form-input { width:100%; box-sizing:border-box; }
  </style>
</head>
<body class="h-full">
<div id="app">
  <div style="display:flex;height:100vh;align-items:center;justify-content:center;">
    <div class="spinner" style="width:32px;height:32px;"></div>
  </div>
</div>

<script>
// ── State ────────────────────────────────────────────────────────────────────
const S = {
  view: 'loading',      // 'login' | 'dashboard'
  tab: 'licenses',
  data: null,
  loginErr: null,
  attemptsLeft: 5,
  lockoutEnd: null,
  lockoutTimer: null,
  search: '',
  expandedDomains: new Set(),
  drawer: null,         // { mode:'add'|'edit', type:'user'|'domain', fields:{...} }
  drawerErr: null,
  drawerSaving: false,
  activitySearch: '',
  activitySort: { col: 'lastSeen', dir: 'desc' },
};

// ── Utils ────────────────────────────────────────────────────────────────────
const $ = id => document.getElementById(id);
const esc = s => String(s ?? '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');

function fmtDate(iso) {
  if (!iso) return '—';
  const d = new Date(iso);
  return isNaN(d) ? iso : d.toLocaleDateString('en-GB', { day:'2-digit', month:'short', year:'numeric' });
}

function fmtDateTime(iso) {
  if (!iso) return '—';
  const d = new Date(iso);
  return isNaN(d) ? iso : d.toLocaleString('en-GB', { day:'2-digit', month:'short', year:'numeric', hour:'2-digit', minute:'2-digit' });
}

function statusBadge(status) {
  const map = { active:'badge-green', trialing:'badge-yellow', canceled:'badge-red', none:'badge-gray', free:'badge-gray' };
  const cls = map[status] || 'badge-gray';
  return '<span class="badge ' + cls + '">' + esc(status || 'unknown') + '</span>';
}

function planBadge(plan) {
  const map = { 'Enterprise':'badge-blue', 'Professional':'badge-green', 'Starter':'badge-gray' };
  const cls = map[plan] || 'badge-gray';
  return '<span class="badge ' + cls + '">' + esc(plan) + '</span>';
}

function normPlan(plan) {
  if (!plan || plan === 'free') return 'Starter';
  const p = plan.toLowerCase();
  if (p.includes('enterprise')) return 'Enterprise';
  if (p.includes('pro')) return 'Professional';
  return plan;
}

function priceDisplay(row) {
  if (row.priceInfo) return esc(row.priceInfo.display) + ' <span style="color:#64748b;font-size:11px;">' + esc(row.priceInfo.billing) + '</span>';
  if (row.amount) return '$' + parseFloat(row.amount).toFixed(2);
  return '—';
}

// ── API ──────────────────────────────────────────────────────────────────────
async function apiFetch(path, opts) {
  const r = await fetch(path, { headers: { 'Content-Type': 'application/json' }, ...opts });
  const data = await r.json().catch(() => ({}));
  return { ok: r.ok, status: r.status, data };
}

// ── Init ─────────────────────────────────────────────────────────────────────
async function init() {
  const r = await apiFetch('/api/me');
  if (r.ok) {
    S.view = 'dashboard';
    await loadData();
  } else {
    S.view = 'login';
  }
  render();
}

async function loadData() {
  const r = await apiFetch('/api/data');
  if (r.ok) S.data = r.data;
}

// ── Auth ─────────────────────────────────────────────────────────────────────
async function doLogin() {
  const btn = $('loginBtn');
  if (btn.disabled) return;
  btn.disabled = true;
  btn.textContent = 'Signing in…';

  const username = $('loginUser').value.trim();
  const password = $('loginPass').value;
  const remember = $('rememberMe').checked;

  const r = await apiFetch('/api/auth/login', { method:'POST', body: JSON.stringify({ username, password, remember }) });
  if (r.ok) {
    S.loginErr = null;
    S.view = 'dashboard';
    await loadData();
    render();
  } else {
    S.loginErr = r.data.error || 'Login failed.';
    S.attemptsLeft = r.data.attemptsRemaining ?? S.attemptsLeft;
    if (r.data.retryAfter) startLockout(r.data.retryAfter);
    else render();
  }
}

function startLockout(sec) {
  S.lockoutEnd = Date.now() + sec * 1000;
  if (S.lockoutTimer) clearInterval(S.lockoutTimer);
  S.lockoutTimer = setInterval(() => {
    if (Date.now() >= S.lockoutEnd) { clearInterval(S.lockoutTimer); S.lockoutEnd = null; }
    render();
  }, 1000);
  render();
}

async function doLogout() {
  await apiFetch('/api/auth/logout', { method:'POST' });
  S.view = 'login'; S.data = null; S.tab = 'licenses'; render();
}

// ── Render ───────────────────────────────────────────────────────────────────
function render() {
  $('app').innerHTML = S.view === 'login' ? renderLogin() : S.view === 'dashboard' ? renderDashboard() : '';
  attachHandlers();
}

function renderLogin() {
  const locked = S.lockoutEnd && Date.now() < S.lockoutEnd;
  const rem = locked ? Math.ceil((S.lockoutEnd - Date.now()) / 1000) : 0;
  return \`
  <div style="min-height:100vh;display:flex;align-items:center;justify-content:center;padding:16px;">
    <div style="width:100%;max-width:380px;">
      <div style="text-align:center;margin-bottom:28px;">
        <div style="font-size:28px;font-weight:800;color:#f1f5f9;letter-spacing:-.02em;">⚡ SheetSync</div>
        <div style="color:#64748b;font-size:13px;margin-top:4px;">Admin Dashboard</div>
      </div>
      <div style="background:#0f172a;border:1px solid #1e293b;border-radius:16px;padding:32px;">
        \${S.loginErr ? \`<div style="background:#450a0a22;border:1px solid #7f1d1d;color:#fca5a5;font-size:13px;padding:10px 12px;border-radius:8px;margin-bottom:16px;">\${esc(S.loginErr)}\${locked ? \` — locked for \${rem}s\` : S.attemptsLeft > 0 ? \` (\${S.attemptsLeft} left)\` : ''}</div>\` : ''}
        <div style="display:flex;flex-direction:column;gap:14px;">
          <div>
            <label style="display:block;font-size:11px;font-weight:600;color:#64748b;text-transform:uppercase;letter-spacing:.06em;margin-bottom:6px;">Username</label>
            <input id="loginUser" type="text" placeholder="admin" autocomplete="username" style="width:100%;box-sizing:border-box;" />
          </div>
          <div>
            <label style="display:block;font-size:11px;font-weight:600;color:#64748b;text-transform:uppercase;letter-spacing:.06em;margin-bottom:6px;">Password</label>
            <input id="loginPass" type="password" placeholder="••••••••" autocomplete="current-password" style="width:100%;box-sizing:border-box;" />
          </div>
          <div style="display:flex;align-items:center;gap:8px;">
            <input id="rememberMe" type="checkbox" style="width:15px;height:15px;accent-color:#3b82f6;cursor:pointer;" />
            <label for="rememberMe" style="font-size:13px;color:#94a3b8;cursor:pointer;user-select:none;">Remember me for 30 days</label>
          </div>
          <button id="loginBtn" class="btn-primary" \${locked ? 'disabled' : ''} style="margin-top:4px;">
            \${locked ? \`Locked (\${rem}s)\` : 'Sign In'}
          </button>
        </div>
      </div>
    </div>
  </div>\`;
}

function renderDashboard() {
  if (!S.data) return \`
  <div style="display:flex;height:100vh;align-items:center;justify-content:center;gap:12px;color:#64748b;">
    <div class="spinner"></div><span>Loading data…</span>
  </div>\`;

  const { stats } = S.data;
  return \`
  \${renderDrawer()}
  <div class="drawer-overlay \${S.drawer ? 'open' : ''}" onclick="closeDrawer()"></div>
  <div style="display:flex;height:100vh;overflow:hidden;">
    <!-- Sidebar -->
    <aside style="width:220px;min-width:220px;background:#080f1e;border-right:1px solid #1e293b;display:flex;flex-direction:column;padding:20px 12px;">
      <div style="padding:8px 12px;margin-bottom:20px;">
        <div style="font-size:17px;font-weight:800;color:#f1f5f9;">⚡ SheetSync</div>
        <div style="font-size:11px;color:#475569;margin-top:2px;">Admin Panel</div>
      </div>
      <nav style="flex:1;display:flex;flex-direction:column;gap:2px;">
        <div class="sidebar-link \${S.tab==='licenses'?'active':''}" onclick="setTab('licenses')">
          <span style="font-size:16px;">🪪</span> Licenses
        </div>
        <div class="sidebar-link \${S.tab==='activity'?'active':''}" onclick="setTab('activity')">
          <span style="font-size:16px;">📊</span> Activity
        </div>
        <div class="sidebar-link \${S.tab==='sales'?'active':''}" onclick="setTab('sales')">
          <span style="font-size:16px;">💳</span> Sales
        </div>
      </nav>
      <button onclick="doLogout()" style="display:flex;align-items:center;gap:8px;padding:9px 12px;border-radius:8px;background:transparent;border:1px solid #1e293b;color:#94a3b8;font-size:13px;cursor:pointer;transition:all .15s;width:100%;"
        onmouseover="this.style.borderColor='#ef4444';this.style.color='#f87171';" onmouseout="this.style.borderColor='#1e293b';this.style.color='#94a3b8';">
        <span>↩</span> Sign Out
      </button>
    </aside>

    <!-- Main -->
    <main style="flex:1;overflow-y:auto;padding:28px 32px;">
      <!-- Stats -->
      <div style="display:grid;grid-template-columns:repeat(4,1fr);gap:16px;margin-bottom:28px;">
        <div class="stat-card">
          <div style="font-size:11px;color:#64748b;font-weight:600;text-transform:uppercase;letter-spacing:.06em;margin-bottom:6px;">Domains</div>
          <div style="font-size:28px;font-weight:800;color:#f1f5f9;">\${stats.totalDomains}</div>
        </div>
        <div class="stat-card">
          <div style="font-size:11px;color:#64748b;font-weight:600;text-transform:uppercase;letter-spacing:.06em;margin-bottom:6px;">Total Users</div>
          <div style="font-size:28px;font-weight:800;color:#f1f5f9;">\${stats.totalUsers}</div>
        </div>
        <div class="stat-card">
          <div style="font-size:11px;color:#64748b;font-weight:600;text-transform:uppercase;letter-spacing:.06em;margin-bottom:6px;">Active Licenses</div>
          <div style="font-size:28px;font-weight:800;color:#4ade80;">\${stats.activeLicenses}</div>
        </div>
        <div class="stat-card">
          <div style="font-size:11px;color:#64748b;font-weight:600;text-transform:uppercase;letter-spacing:.06em;margin-bottom:6px;">Total Revenue</div>
          <div style="font-size:28px;font-weight:800;color:#f1f5f9;">$\${stats.totalRevenue.toLocaleString('en-US', {minimumFractionDigits:2,maximumFractionDigits:2})}</div>
        </div>
      </div>

      <!-- Tab Content -->
      \${S.tab === 'licenses' ? renderLicenses() : ''}
      \${S.tab === 'activity'  ? renderActivity()  : ''}
      \${S.tab === 'sales'     ? renderSales()     : ''}
    </main>
  </div>\`;
}

// ── Tab: Licenses ─────────────────────────────────────────────────────────────
function renderLicenses() {
  const { domains, individuals } = S.data;
  const q = S.search.toLowerCase();

  const filteredDomains = domains.filter(d =>
    !q || d.domain.includes(q) || d.users.some(u => u.email.includes(q))
  );
  const filteredIndividuals = individuals.filter(u =>
    !q || u.email.includes(q)
  );

  return \`
  <div>
    <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:20px;">
      <h2 style="font-size:18px;font-weight:700;color:#f1f5f9;margin:0;">Licenses</h2>
      <div style="display:flex;gap:10px;align-items:center;">
        <input type="search" id="searchInput" placeholder="Search email or domain…" value="\${esc(S.search)}"
          style="width:220px;" oninput="S.search=this.value;renderContent();" />
        <button class="btn-primary" onclick="openDrawer('add','user')" style="padding:8px 14px;font-size:13px;white-space:nowrap;">+ Add License</button>
      </div>
    </div>

    \${filteredDomains.length ? \`
    <h3 style="font-size:12px;font-weight:600;color:#64748b;text-transform:uppercase;letter-spacing:.06em;margin:0 0 12px;">Domain Licenses (\${filteredDomains.length})</h3>
    \${filteredDomains.map(d => renderDomainCard(d)).join('')}
    \` : ''}

    \${filteredIndividuals.length ? \`
    <h3 style="font-size:12px;font-weight:600;color:#64748b;text-transform:uppercase;letter-spacing:.06em;margin:24px 0 12px;">Individual Licenses (\${filteredIndividuals.length})</h3>
    <div class="panel">
      <table class="tbl">
        <thead><tr>
          <th>Email</th><th>Plan</th><th>Status</th><th>Renewal</th><th>Price</th><th>Last Seen</th><th></th>
        </tr></thead>
        <tbody>
          \${filteredIndividuals.map(u => \`<tr>
            <td style="font-family:monospace;color:#93c5fd;">\${esc(u.email)}</td>
            <td>\${planBadge(u.planLabel || normPlan(u.plan))}</td>
            <td>\${statusBadge(u.status)}</td>
            <td style="color:#94a3b8;">\${fmtDate(u.renewsAt)}</td>
            <td>\${priceDisplay(u)}</td>
            <td style="color:#64748b;font-size:12px;">\${fmtDateTime(u.lastSeen)}</td>
            <td style="white-space:nowrap;">
              <button class="btn-ghost" onclick="openDrawer('edit','user',\${JSON.stringify(u).replace(/"/g,'&quot;')})">Edit</button>
              <button class="btn-danger" style="margin-left:6px;" onclick="deleteLicense('user',\${JSON.stringify({email:u.email}).replace(/"/g,'&quot;')})">Delete</button>
            </td>
          </tr>\`).join('')}
        </tbody>
      </table>
    </div>
    \` : ''}

    \${!filteredDomains.length && !filteredIndividuals.length ? \`
    <div style="text-align:center;padding:60px 20px;color:#475569;">
      \${q ? 'No results for "' + esc(q) + '"' : 'No license records yet.'}
    </div>\` : ''}
  </div>\`;
}

function renderDomainCard(d) {
  const expanded = S.expandedDomains.has(d.domain);
  const plan = d.planLabel || normPlan(d.plan);
  return \`
  <div class="domain-card">
    <div class="domain-header" onclick="toggleDomain('\${esc(d.domain)}')">
      <div style="display:flex;align-items:center;gap:12px;">
        <div style="width:36px;height:36px;background:#1e3a8a33;border-radius:8px;display:flex;align-items:center;justify-content:center;font-size:16px;">🏢</div>
        <div>
          <div style="font-weight:700;color:#f1f5f9;font-size:14px;">\${esc(d.domain)}</div>
          <div style="font-size:12px;color:#64748b;margin-top:2px;">\${d.users.length} user\${d.users.length!==1?'s':''} · purchased by \${esc(d.email||'unknown')}</div>
        </div>
      </div>
      <div style="display:flex;align-items:center;gap:12px;">
        \${planBadge(plan)}
        \${statusBadge(d.status)}
        <div style="text-align:right;font-size:12px;color:#64748b;">
          <div>\${priceDisplay(d)}</div>
          <div>Renews \${fmtDate(d.renewsAt)}</div>
        </div>
        <button class="btn-ghost" onclick="event.stopPropagation();openDrawer('edit','domain',\${JSON.stringify(d).replace(/"/g,'&quot;')})">Edit</button>
        <button class="btn-danger" onclick="event.stopPropagation();deleteLicense('domain',\${JSON.stringify({domain:d.domain}).replace(/"/g,'&quot;')})">Delete</button>
        <span style="color:#475569;font-size:18px;transition:transform .2s;\${expanded?'transform:rotate(90deg)':''}">›</span>
      </div>
    </div>
    \${expanded && d.users.length ? \`
    <div style="border-top:1px solid #1e293b;">
      <table class="tbl">
        <thead><tr>
          <th>User Email</th><th>Plan</th><th>Status</th><th>Renewal</th><th>Price</th><th>Last Seen</th>
        </tr></thead>
        <tbody>
          \${d.users.map(u => \`<tr>
            <td style="font-family:monospace;color:#93c5fd;">\${esc(u.email)}</td>
            <td>\${planBadge(u.planLabel || normPlan(u.plan))}</td>
            <td>\${statusBadge(u.status)}</td>
            <td style="color:#94a3b8;">\${fmtDate(u.renewsAt)}</td>
            <td>\${priceDisplay(u)}</td>
            <td style="color:#64748b;font-size:12px;">\${fmtDateTime(u.lastSeen)}</td>
          </tr>\`).join('')}
        </tbody>
      </table>
    </div>\` : expanded ? \`<div style="padding:16px 18px;color:#475569;font-size:13px;">No individual user records under this domain yet.</div>\` : ''}
  </div>\`;
}

// ── Tab: Activity ─────────────────────────────────────────────────────────────
function renderActivity() {
  const q = S.activitySearch.toLowerCase();
  const { col, dir } = S.activitySort;

  let logins = S.data.logins.map(l => ({ ...l, domain: l.email.split('@')[1] || '' }));
  if (q) logins = logins.filter(l => l.email.includes(q) || l.domain.includes(q));
  logins = [...logins].sort((a, b) => {
    const av = col === 'count' ? (a[col] || 0) : (a[col] || '');
    const bv = col === 'count' ? (b[col] || 0) : (b[col] || '');
    if (av < bv) return dir === 'asc' ? -1 : 1;
    if (av > bv) return dir === 'asc' ? 1 : -1;
    return 0;
  });

  const sh = (c, label, align) => {
    const arrow = col === c ? (dir === 'asc' ? '↑' : '↓') : '⇅';
    const arrowColor = col === c ? '#60a5fa' : '#374151';
    const ta = align ? \`text-align:\${align};\` : '';
    return \`<th style="cursor:pointer;user-select:none;\${ta}" onclick="sortActivity('\${c}')">\${label} <span style="color:\${arrowColor}">\${arrow}</span></th>\`;
  };

  return \`
  <div>
    <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:20px;">
      <h2 style="font-size:18px;font-weight:700;color:#f1f5f9;margin:0;">Login Activity</h2>
      <div style="display:flex;align-items:center;gap:12px;">
        <input type="search" id="activitySearch" placeholder="Filter by email or domain…" value="\${esc(S.activitySearch)}"
          style="width:240px;" oninput="S.activitySearch=this.value;renderContent();" />
        <span style="color:#64748b;font-size:13px;">\${logins.length} user\${logins.length!==1?'s':''}</span>
      </div>
    </div>
    \${logins.length ? \`
    <div class="panel">
      <table class="tbl">
        <thead><tr>
          \${sh('email','Email')}
          \${sh('domain','Domain')}
          \${sh('plan','Plan')}
          \${sh('firstSeen','First Seen')}
          \${sh('lastSeen','Last Seen')}
          \${sh('count','Visits','right')}
        </tr></thead>
        <tbody>
          \${logins.map(l => \`<tr>
            <td style="font-family:monospace;color:#93c5fd;">\${esc(l.email)}</td>
            <td style="color:#94a3b8;">\${esc(l.domain)}</td>
            <td>\${planBadge(normPlan(l.plan))}</td>
            <td style="color:#94a3b8;font-size:12px;">\${fmtDateTime(l.firstSeen)}</td>
            <td style="color:#94a3b8;font-size:12px;">\${fmtDateTime(l.lastSeen)}</td>
            <td style="text-align:right;font-weight:600;color:#f1f5f9;">\${l.count||1}</td>
          </tr>\`).join('')}
        </tbody>
      </table>
    </div>\` : \`
    <div style="text-align:center;padding:60px 20px;color:#475569;">\${q ? \`No results for "\${esc(q)}"\` : 'No login activity recorded yet.'}</div>\`}
  </div>\`;
}

// ── Tab: Sales ────────────────────────────────────────────────────────────────
function renderSales() {
  const { sales, stats } = S.data;
  return \`
  <div>
    <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:20px;">
      <h2 style="font-size:18px;font-weight:700;color:#f1f5f9;margin:0;">Sales History</h2>
      <div style="background:#0f172a;border:1px solid #1e293b;border-radius:10px;padding:10px 18px;text-align:right;">
        <div style="font-size:11px;color:#64748b;text-transform:uppercase;letter-spacing:.06em;">Total Revenue</div>
        <div style="font-size:20px;font-weight:800;color:#4ade80;">$\${stats.totalRevenue.toLocaleString('en-US',{minimumFractionDigits:2,maximumFractionDigits:2})}</div>
      </div>
    </div>
    \${sales.length ? \`
    <div class="panel">
      <table class="tbl">
        <thead><tr>
          <th>Date</th><th>Email</th><th>Plan</th><th>Domain</th><th>Stripe Customer</th><th style="text-align:right;">Amount</th>
        </tr></thead>
        <tbody>
          \${sales.map(s => \`<tr>
            <td style="color:#94a3b8;font-size:12px;white-space:nowrap;">\${fmtDate(s.date)}</td>
            <td style="font-family:monospace;color:#93c5fd;">\${esc(s.email)}</td>
            <td>\${planBadge(normPlan(s.plan))}</td>
            <td style="color:#94a3b8;">\${esc(s.domain||'—')}</td>
            <td style="font-family:monospace;font-size:11px;color:#475569;">\${esc(s.customer||'—')}</td>
            <td style="text-align:right;font-weight:600;color:#4ade80;">$\${parseFloat(s.amount||0).toFixed(2)}</td>
          </tr>\`).join('')}
        </tbody>
      </table>
    </div>\` : \`
    <div style="text-align:center;padding:60px 20px;color:#475569;">No sales recorded yet.</div>\`}
  </div>\`;
}

// ── Drawer ────────────────────────────────────────────────────────────────────
function renderDrawer() {
  const d = S.drawer;
  return \`
  <div class="drawer \${d ? 'open' : ''}" id="drawer">
    \${d ? \`
    <div style="display:flex;align-items:center;justify-content:space-between;padding:20px 24px;border-bottom:1px solid #1e293b;">
      <div>
        <div style="font-size:15px;font-weight:700;color:#f1f5f9;">\${d.mode === 'add' ? 'Add License' : 'Edit License'}</div>
        <div style="font-size:11px;color:#475569;margin-top:2px;">\${d.type === 'domain' ? 'Domain license' : 'User license'}</div>
      </div>
      <button onclick="closeDrawer()" style="background:none;border:none;color:#64748b;cursor:pointer;font-size:20px;line-height:1;">×</button>
    </div>
    <div style="flex:1;overflow-y:auto;padding:24px;display:flex;flex-direction:column;gap:18px;">
      \${S.drawerErr ? \`<div style="background:#450a0a22;border:1px solid #7f1d1d;color:#fca5a5;font-size:13px;padding:10px 12px;border-radius:8px;">\${esc(S.drawerErr)}</div>\` : ''}

      \${d.mode === 'add' ? \`
      <div class="form-row">
        <label class="form-label">License Type</label>
        <select class="form-input" id="df_type" onchange="S.drawer.type=this.value;renderDrawerContent();">
          <option value="user" \${d.type==='user'?'selected':''}>User (by email)</option>
          <option value="domain" \${d.type==='domain'?'selected':''}>Domain (all @domain users)</option>
        </select>
      </div>\` : ''}

      \${d.type === 'user' ? \`
      <div class="form-row">
        <label class="form-label">Email</label>
        <input type="text" class="form-input" id="df_email" placeholder="user@example.com"
          value="\${esc(d.fields.email||'')}" \${d.mode==='edit'?'readonly style="opacity:.5;"':''} />
      </div>
      <div class="form-row">
        <label class="form-label">Domain (optional)</label>
        <input type="text" class="form-input" id="df_domain" placeholder="example.com" value="\${esc(d.fields.domain||'')}" />
      </div>\` : \`
      <div class="form-row">
        <label class="form-label">Domain</label>
        <input type="text" class="form-input" id="df_domain" placeholder="company.com"
          value="\${esc(d.fields.domain||'')}" \${d.mode==='edit'?'readonly style="opacity:.5;"':''} />
      </div>
      <div class="form-row">
        <label class="form-label">Buyer Email (optional)</label>
        <input type="text" class="form-input" id="df_email" placeholder="owner@company.com" value="\${esc(d.fields.email||'')}" />
      </div>\`}

      <div class="form-row">
        <label class="form-label">Plan</label>
        <select class="form-input" id="df_plan">
          <option value="starter"    \${d.fields.plan==='starter'   ?'selected':''}>Starter</option>
          <option value="pro"        \${d.fields.plan==='pro'       ?'selected':''}>Pro</option>
          <option value="enterprise" \${d.fields.plan==='enterprise'?'selected':''}>Enterprise</option>
        </select>
      </div>

      <div class="form-row">
        <label class="form-label">Status</label>
        <select class="form-input" id="df_status">
          <option value="active"   \${d.fields.status==='active'  ?'selected':''}>Active</option>
          <option value="trialing" \${d.fields.status==='trialing'?'selected':''}>Trialing</option>
          <option value="canceled" \${d.fields.status==='canceled'?'selected':''}>Canceled</option>
        </select>
      </div>

      <div class="form-row">
        <label class="form-label">Renews At (optional)</label>
        <input type="text" class="form-input" id="df_renewsAt" placeholder="YYYY-MM-DD"
          value="\${esc((d.fields.renewsAt||'').slice(0,10))}" />
      </div>
    </div>
    <div style="padding:20px 24px;border-top:1px solid #1e293b;display:flex;gap:10px;">
      <button class="btn-primary" onclick="saveDrawer()" \${S.drawerSaving?'disabled':''} style="flex:1;">
        \${S.drawerSaving ? 'Saving…' : (d.mode === 'add' ? 'Add License' : 'Save Changes')}
      </button>
      <button class="btn-ghost" onclick="closeDrawer()" style="padding:10px 16px;">Cancel</button>
    </div>
    \` : ''}
  </div>\`;
}

function renderDrawerContent() {
  const el = $('drawer');
  if (el) el.outerHTML = renderDrawer();
}

function openDrawer(mode, type, row) {
  const fields = row ? { ...row } : {};
  if (!fields.plan) fields.plan = 'starter';
  if (!fields.status) fields.status = 'active';
  S.drawer = { mode, type, fields };
  S.drawerErr = null;
  S.drawerSaving = false;
  render();
}

function closeDrawer() {
  S.drawer = null;
  S.drawerErr = null;
  render();
}

async function saveDrawer() {
  const d = S.drawer;
  if (!d || S.drawerSaving) return;

  const email    = $('df_email')?.value.trim();
  const domain   = $('df_domain')?.value.trim();
  const plan     = $('df_plan').value;
  const status   = $('df_status').value;
  const renewsAt = $('df_renewsAt')?.value.trim() || null;

  if (d.type === 'user' && !email)   { S.drawerErr = 'Email is required.'; renderDrawerContent(); return; }
  if (d.type === 'domain' && !domain){ S.drawerErr = 'Domain is required.'; renderDrawerContent(); return; }

  S.drawerSaving = true;
  renderDrawerContent();

  const r = await apiFetch('/api/license', {
    method: 'POST',
    body: JSON.stringify({ type: d.type, email, domain, plan, status, renewsAt }),
  });

  if (r.ok) {
    S.drawer = null;
    await loadData();
    render();
  } else {
    S.drawerErr = r.data.error || 'Save failed.';
    S.drawerSaving = false;
    renderDrawerContent();
  }
}

async function deleteLicense(type, keyObj) {
  const label = type === 'domain' ? keyObj.domain : keyObj.email;
  if (!confirm(\`Delete license for \${label}?\`)) return;
  const r = await apiFetch('/api/license', {
    method: 'DELETE',
    body: JSON.stringify({ type, ...keyObj }),
  });
  if (r.ok) { await loadData(); render(); }
  else alert('Delete failed: ' + (r.data.error || 'unknown error'));
}

// ── Interactions ──────────────────────────────────────────────────────────────
function setTab(tab) { S.tab = tab; render(); }

function sortActivity(col) {
  if (S.activitySort.col === col) {
    S.activitySort.dir = S.activitySort.dir === 'asc' ? 'desc' : 'asc';
  } else {
    S.activitySort = { col, dir: col === 'count' ? 'desc' : 'asc' };
  }
  renderContent();
}

function toggleDomain(domain) {
  if (S.expandedDomains.has(domain)) S.expandedDomains.delete(domain);
  else S.expandedDomains.add(domain);
  renderContent();
}

function renderContent() {
  // Re-render only the main content area to preserve sidebar state
  const main = document.querySelector('main');
  if (!main) { render(); return; }
  const statRow = main.querySelector(':scope > div:first-child');
  const tabContent = main.querySelector(':scope > div:last-child');
  if (!tabContent || !statRow) { render(); return; }
  tabContent.outerHTML =
    (S.tab === 'licenses' ? renderLicenses() : '') +
    (S.tab === 'activity' ? renderActivity() : '') +
    (S.tab === 'sales'    ? renderSales()    : '');
  // Update sidebar active state
  document.querySelectorAll('.sidebar-link').forEach(el => {
    const t = el.getAttribute('onclick')?.match(/setTab\('(\w+)'\)/)?.[1];
    if (t) el.classList.toggle('active', t === S.tab);
  });
}

function attachHandlers() {
  const loginBtn = $('loginBtn');
  if (loginBtn) {
    loginBtn.addEventListener('click', doLogin);
    $('loginPass')?.addEventListener('keydown', e => { if (e.key === 'Enter') doLogin(); });
    $('loginUser')?.addEventListener('keydown', e => { if (e.key === 'Enter') $('loginPass')?.focus(); });
  }
}

init();
</script>
</body>
</html>`;

  return new Response(html, {
    headers: { 'Content-Type': 'text/html; charset=utf-8' },
  });
}

// ─── License CRUD ─────────────────────────────────────────────────────────────

async function handleUpsertLicense(request, env) {
  const sess = await getSession(request, env);
  if (!sess) return json({ error: 'Unauthorized' }, 401);

  let body;
  try { body = await request.json(); } catch { return json({ error: 'Bad request' }, 400); }

  const { type, email, domain, plan, status, renewsAt } = body;
  const allowed = (status === 'active' || status === 'trialing') ? 1 : 0;
  const now = new Date().toISOString();

  if (type === 'domain') {
    if (!domain) return json({ error: 'domain is required' }, 400);
    await env.DB.prepare(`
      INSERT INTO domain_licenses (domain, email, plan, status, allowed, renews_at, last_updated)
      VALUES (?, ?, ?, ?, ?, ?, ?)
      ON CONFLICT(domain) DO UPDATE SET
        email = excluded.email, plan = excluded.plan, status = excluded.status,
        allowed = excluded.allowed, renews_at = excluded.renews_at, last_updated = excluded.last_updated
    `).bind(domain.toLowerCase(), email || null, plan, status, allowed, renewsAt || null, now).run();
    return json({ ok: true });
  }

  // default: user license
  if (!email) return json({ error: 'email is required' }, 400);
  await env.DB.prepare(`
    INSERT INTO user_licenses (email, domain, plan, status, allowed, renews_at, last_updated)
    VALUES (?, ?, ?, ?, ?, ?, ?)
    ON CONFLICT(email) DO UPDATE SET
      domain = excluded.domain, plan = excluded.plan, status = excluded.status,
      allowed = excluded.allowed, renews_at = excluded.renews_at, last_updated = excluded.last_updated
  `).bind(email.toLowerCase(), domain || null, plan, status, allowed, renewsAt || null, now).run();
  return json({ ok: true });
}

async function handleDeleteLicense(request, env) {
  const sess = await getSession(request, env);
  if (!sess) return json({ error: 'Unauthorized' }, 401);

  let body;
  try { body = await request.json(); } catch { return json({ error: 'Bad request' }, 400); }

  const { type, email, domain } = body;

  if (type === 'domain') {
    if (!domain) return json({ error: 'domain is required' }, 400);
    await env.DB.prepare('DELETE FROM domain_licenses WHERE domain = ?').bind(domain.toLowerCase()).run();
    return json({ ok: true });
  }

  if (!email) return json({ error: 'email is required' }, 400);
  await env.DB.prepare('DELETE FROM user_licenses WHERE email = ?').bind(email.toLowerCase()).run();
  return json({ ok: true });
}

// ─── Router ───────────────────────────────────────────────────────────────────

export default {
  async fetch(request, env) {
    const url = new URL(request.url);
    const path = url.pathname;

    if (request.method === 'OPTIONS') {
      return new Response(null, { headers: { Allow: 'GET, POST, DELETE, OPTIONS' } });
    }

    if (path === '/api/auth/login'  && request.method === 'POST') return handleLogin(request, env);
    if (path === '/api/auth/logout' && request.method === 'POST') return handleLogout(request, env);
    if (path === '/api/me')   return handleMe(request, env);
    if (path === '/api/data') return handleData(request, env);
    if (path === '/api/license' && request.method === 'POST')   return handleUpsertLicense(request, env);
    if (path === '/api/license' && request.method === 'DELETE') return handleDeleteLicense(request, env);

    return serveSPA();
  },
};
