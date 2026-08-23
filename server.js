const express = require('express');
const path = require('path');

const VIBE_BASE_URL = process.env.VIBE_BASE_URL || 'https://vibecode.bitrix24.com/v1';
const VIBE_APP_KEY = process.env.VIBE_APP_KEY || '';
const PORT = process.env.PORT || 3000;
const ADDRESS_TYPE_ID = 1; // "actual" address — available in every country zone
const CONTACT_ENTITY_TYPE_ID = 3;

if (!VIBE_APP_KEY) {
  console.error('FATAL: VIBE_APP_KEY is not set. Set it via the deploy env (the app\'s own vibe_app_* key) so this server can identify itself to the Vibecode API alongside each user\'s forwarded session.');
}

// ---------------------------------------------------------------------------
// Rate limiter: Vibecode enforces a request rate per key. The X-Api-Key here
// identifies the app (shared across every user's session), so gate it once
// for the whole process.
// ---------------------------------------------------------------------------
class RateLimiter {
  constructor(maxPerWindow, windowMs) {
    this.maxPerWindow = maxPerWindow;
    this.windowMs = windowMs;
    this.timestamps = [];
    this.queue = [];
  }
  acquire() {
    return new Promise((resolve) => {
      this.queue.push(resolve);
      this._pump();
    });
  }
  _pump() {
    if (this.queue.length === 0) return;
    const now = Date.now();
    this.timestamps = this.timestamps.filter((t) => now - t < this.windowMs);
    if (this.timestamps.length < this.maxPerWindow) {
      this.timestamps.push(now);
      const resolve = this.queue.shift();
      resolve();
      if (this.queue.length > 0) setImmediate(() => this._pump());
    } else {
      const waitFor = this.windowMs - (now - this.timestamps[0]) + 5;
      setTimeout(() => this._pump(), waitFor);
    }
  }
}
const limiter = new RateLimiter(8, 1000);

// Per-user Vibecode call: the app's own key identifies the app, the forwarded
// bearer identifies the acting Bitrix24 employee, so every write (contact,
// company, address) lands in CRM as created by that employee, not the app.
async function vibeFetch(bearer, pathAndQuery, options = {}) {
  await limiter.acquire();
  const url = `${VIBE_BASE_URL}${pathAndQuery}`;
  const headers = Object.assign(
    { 'X-Api-Key': VIBE_APP_KEY, Authorization: `Bearer ${bearer}` },
    options.body ? { 'Content-Type': 'application/json' } : {},
    options.headers || {}
  );
  const res = await fetch(url, {
    method: options.method || 'GET',
    headers,
    body: options.body ? JSON.stringify(options.body) : undefined,
  });

  if (res.status === 429) {
    const retryAfter = Number(res.headers.get('Retry-After') || '2');
    await new Promise((r) => setTimeout(r, Math.min(retryAfter, 10) * 1000));
    return vibeFetch(bearer, pathAndQuery, options);
  }

  const json = await res.json().catch(() => ({ success: false, error: { code: 'BAD_JSON', message: 'Non-JSON response' } }));
  return { ok: res.ok && json.success !== false, status: res.status, json };
}

function qs(params) {
  const parts = [];
  for (const [k, v] of Object.entries(params)) {
    if (v === undefined || v === null || v === '') continue;
    if (k === 'filter') {
      parts.push(`filter=${encodeURIComponent(JSON.stringify(v))}`);
    } else if (k === 'select' && Array.isArray(v)) {
      parts.push(`select=${encodeURIComponent(v.join(','))}`);
    } else {
      parts.push(`${encodeURIComponent(k)}=${encodeURIComponent(String(v))}`);
    }
  }
  return parts.length ? `?${parts.join('&')}` : '';
}

// ---------------------------------------------------------------------------
// Identity: resolve the Gateway-forwarded session once per token, per the
// Black Hole BFF pattern (see /docs/infra/app-runtime).
// ---------------------------------------------------------------------------
const identityCache = new Map(); // bearer -> { at, me }
const IDENTITY_TTL_MS = 10 * 60 * 1000;

async function resolveIdentity(bearer) {
  const cached = identityCache.get(bearer);
  if (cached && Date.now() - cached.at < IDENTITY_TTL_MS) return cached.me;
  const res = await fetch(`${VIBE_BASE_URL}/me`, {
    headers: { 'X-Api-Key': VIBE_APP_KEY, Authorization: `Bearer ${bearer}` },
  });
  if (!res.ok) return null;
  const json = await res.json().catch(() => null);
  if (!json || !json.success) return null;
  identityCache.set(bearer, { at: Date.now(), me: json.data });
  return json.data;
}

async function requireUser(req, res, next) {
  const raw = req.headers['x-vibe-authorization'];
  const bearer = raw ? String(raw).replace(/^Bearer /, '') : '';
  if (!bearer) {
    return res.status(401).json({ error: { code: 'NO_SESSION', message: 'Sign in again or grant the app access, then reopen it from the Bitrix24 left menu.' } });
  }
  const me = await resolveIdentity(bearer);
  const userId = me && me.currentUser ? me.currentUser.bitrixUserId : null;
  if (!me || !userId) {
    return res.status(401).json({ error: { code: 'NO_SESSION', message: 'Sign in again or grant the app access, then reopen it from the Bitrix24 left menu.' } });
  }
  req.bearer = bearer;
  req.me = me;
  next();
}

// ---------------------------------------------------------------------------
// HTTP API
// ---------------------------------------------------------------------------
const app = express();
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

app.get('/health', (req, res) => res.json({ ok: true }));

app.get('/api/session', async (req, res) => {
  const raw = req.headers['x-vibe-authorization'];
  const bearer = raw ? String(raw).replace(/^Bearer /, '') : '';
  if (!bearer) return res.json({ authenticated: false });
  const me = await resolveIdentity(bearer);
  const userId = me && me.currentUser ? me.currentUser.bitrixUserId : null;
  if (!me || !userId) return res.json({ authenticated: false });
  const nameEncoded = req.headers['x-vibe-user-name-encoded'];
  const userName = nameEncoded ? decodeURIComponent(String(nameEncoded)) : `User ${userId}`;
  res.json({ authenticated: true, portal: me.portal, user: { id: userId, name: userName } });
});

// Lightweight request whose only purpose is to travel through the Gateway
// and keep the placement session's cookie fresh during a long form fill.
app.get('/api/ping', (req, res) => res.json({ ok: true }));

app.get('/api/companies', requireUser, async (req, res) => {
  const q = (req.query.q || '').toString().trim();
  if (!q) return res.json({ data: [] });
  try {
    const { ok, status, json } = await vibeFetch(
      req.bearer,
      `/companies${qs({ filter: { '%title': q }, select: ['id', 'title'], limit: 15 })}`
    );
    if (!ok) return res.status(status || 502).json({ error: json.error || { message: 'Failed to search companies' } });
    res.json({ data: json.data || [] });
  } catch (e) {
    res.status(502).json({ error: { message: e.message } });
  }
});

function multifield(value) {
  const v = (value || '').toString().trim();
  return v || undefined;
}

app.post('/api/contacts', requireUser, async (req, res) => {
  const body = req.body || {};
  const firstName = (body.firstName || '').toString().trim();
  const lastName = (body.lastName || '').toString().trim();
  if (!firstName && !lastName) {
    return res.status(400).json({ error: { code: 'NO_NAME', message: 'Enter at least a first or last name.' } });
  }

  try {
    let companyId = body.companyId ? Number(body.companyId) : undefined;

    if (!companyId && body.newCompanyTitle) {
      const title = body.newCompanyTitle.toString().trim();
      if (title) {
        const { ok, status, json } = await vibeFetch(req.bearer, '/companies', {
          method: 'POST',
          body: { title },
        });
        if (!ok) return res.status(status || 502).json({ error: json.error || { message: 'Failed to create the company' }, step: 'company' });
        companyId = json.data.id;
      }
    }

    const contactBody = {
      name: firstName || undefined,
      lastName: lastName || undefined,
      phone: multifield(body.phone),
      email: multifield(body.email),
      companyId,
      comments: multifield(body.comments),
    };
    Object.keys(contactBody).forEach((k) => contactBody[k] === undefined && delete contactBody[k]);

    const created = await vibeFetch(req.bearer, '/contacts', { method: 'POST', body: contactBody });
    if (!created.ok) {
      return res.status(created.status || 502).json({ error: created.json.error || { message: 'Failed to create the contact' }, step: 'contact' });
    }
    const contactId = created.json.data.id;

    const address = body.address || {};
    const hasAddress = ['line1', 'city', 'region', 'postalCode', 'country'].some((k) => (address[k] || '').toString().trim());
    if (hasAddress) {
      const addrBody = {
        typeId: ADDRESS_TYPE_ID,
        entityTypeId: CONTACT_ENTITY_TYPE_ID,
        entityId: contactId,
        address1: multifield(address.line1),
        city: multifield(address.city),
        province: multifield(address.region),
        postalCode: multifield(address.postalCode),
        country: multifield(address.country),
      };
      Object.keys(addrBody).forEach((k) => addrBody[k] === undefined && delete addrBody[k]);
      const addrRes = await vibeFetch(req.bearer, '/addresses', { method: 'POST', body: addrBody });
      if (!addrRes.ok) {
        // Contact already exists at this point — surface the address failure
        // without pretending the whole operation failed.
        return res.json({
          id: contactId,
          portal: req.me.portal,
          addressWarning: addrRes.json.error ? addrRes.json.error.message : 'The contact was created, but saving the address failed.',
        });
      }
    }

    res.json({ id: contactId, portal: req.me.portal });
  } catch (e) {
    res.status(502).json({ error: { message: e.message } });
  }
});

app.listen(PORT, () => {
  console.log(`Contact Wizard app listening on :${PORT}`);
});
