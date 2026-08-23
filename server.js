const express = require('express');
const multer = require('multer');
const crypto = require('crypto');
const path = require('path');

const app = express();
app.disable('x-powered-by');
app.use(express.json({ limit: '1mb' }));
app.use(express.static(path.join(__dirname, 'public')));

const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 25 * 1024 * 1024 } });

const VIBE_API_KEY = process.env.VIBE_API_KEY || '';
const APP_PIN = process.env.APP_PIN || '';
const VIBE_BASE = 'https://vibecode.bitrix24.com/v1';
const ENTITY_TYPE_ID = 1126;
const FIELD_BERICHT = 'ufCrm79_1787525651154';
const FIELD_KUNDE = 'ufCrm79_1787525659286';
const FIELD_BESUCHSDATUM = 'ufCrm79_1787525674336';
const FIELD_TODOS = 'ufCrm79_1787525690078';
const AI_MODEL = 'bitrix/bitrixgpt-5.5';

if (!VIBE_API_KEY) {
  console.error('WARNING: VIBE_API_KEY is not set. Calls to the Vibecode API will fail.');
}
if (!APP_PIN) {
  console.error('WARNING: APP_PIN is not set. Nobody will be able to log in.');
}

// ---- Sessions (in-memory; simple deterrent, not bank-grade security) ----
const SESSION_TTL_MS = 12 * 60 * 60 * 1000; // 12h
const sessions = new Map(); // token -> expiry timestamp

function issueSession() {
  const token = crypto.randomBytes(24).toString('hex');
  sessions.set(token, Date.now() + SESSION_TTL_MS);
  return token;
}

function isValidSession(token) {
  if (!token) return false;
  const exp = sessions.get(token);
  if (!exp) return false;
  if (Date.now() > exp) {
    sessions.delete(token);
    return false;
  }
  return true;
}

setInterval(() => {
  const now = Date.now();
  for (const [token, exp] of sessions) {
    if (now > exp) sessions.delete(token);
  }
}, 10 * 60 * 1000).unref();

function requireSession(req, res, next) {
  const auth = req.headers.authorization || '';
  const token = auth.startsWith('Bearer ') ? auth.slice(7) : null;
  if (!isValidSession(token)) {
    return res.status(401).json({ error: 'UNAUTHORIZED' });
  }
  next();
}

function pinMatches(input) {
  if (!APP_PIN || typeof input !== 'string' || input.length === 0) return false;
  const a = Buffer.from(input);
  const b = Buffer.from(APP_PIN);
  if (a.length !== b.length) return false;
  return crypto.timingSafeEqual(a, b);
}

// Basic brute-force deterrent per IP.
const failedAttempts = new Map(); // ip -> { count, lockUntil }

function getIp(req) {
  const fwd = req.headers['x-forwarded-for'];
  if (typeof fwd === 'string' && fwd.length > 0) return fwd.split(',')[0].trim();
  return req.socket.remoteAddress || 'unknown';
}

app.post('/api/login', (req, res) => {
  const ip = getIp(req);
  const now = Date.now();
  const rec = failedAttempts.get(ip);
  if (rec && rec.lockUntil > now) {
    return res.status(429).json({ error: 'LOCKED', retryAfterMs: rec.lockUntil - now });
  }

  const pin = req.body && typeof req.body.pin !== 'undefined' ? String(req.body.pin) : '';
  if (!pinMatches(pin)) {
    const count = (rec ? rec.count : 0) + 1;
    const lockUntil = count >= 5 ? now + 60_000 : 0;
    failedAttempts.set(ip, { count, lockUntil });
    return res.status(401).json({ error: 'INVALID_PIN' });
  }

  failedAttempts.delete(ip);
  const token = issueSession();
  res.json({ token });
});

app.get('/api/session', requireSession, (req, res) => {
  res.json({ ok: true });
});

function todayBerlin() {
  return new Intl.DateTimeFormat('sv-SE', { timeZone: 'Europe/Berlin' }).format(new Date());
}

function weekdayBerlin() {
  return new Intl.DateTimeFormat('de-DE', { timeZone: 'Europe/Berlin', weekday: 'long' }).format(new Date());
}

function addDaysToDateStr(dateStr, days) {
  const d = new Date(dateStr + 'T00:00:00Z');
  d.setUTCDate(d.getUTCDate() + days);
  return d.toISOString().slice(0, 10);
}

const WEEKDAYS_DE = ['Sonntag', 'Montag', 'Dienstag', 'Mittwoch', 'Donnerstag', 'Freitag', 'Samstag'];

// Precompute a lookup table of relative German date expressions -> concrete
// dates, so the model only has to pick the right entry instead of doing
// date arithmetic itself (small/free models are unreliable at that).
function relativeDateTable(todayStr) {
  const table = { heute: todayStr, gestern: addDaysToDateStr(todayStr, -1), vorgestern: addDaysToDateStr(todayStr, -2) };
  const todayDow = new Date(todayStr + 'T00:00:00Z').getUTCDay();
  for (let back = 1; back <= 7; back++) {
    const dow = (todayDow - back + 7) % 7;
    const name = WEEKDAYS_DE[dow];
    if (!(name in table)) table[name] = addDaysToDateStr(todayStr, -back);
  }
  return table;
}

app.post('/api/transcribe', requireSession, upload.single('audio'), async (req, res) => {
  try {
    if (!req.file || !req.file.buffer || req.file.buffer.length === 0) {
      return res.status(400).json({ error: 'NO_AUDIO' });
    }

    const form = new FormData();
    const blob = new Blob([req.file.buffer], { type: req.file.mimetype || 'audio/webm' });
    form.append('file', blob, 'aufnahme.webm');
    form.append('language', 'de');
    form.append('response_format', 'json');

    const r = await fetch(`${VIBE_BASE}/audio/transcriptions`, {
      method: 'POST',
      headers: { 'X-Api-Key': VIBE_API_KEY },
      body: form,
    });
    const data = await r.json();
    if (!r.ok) {
      console.error('transcribe upstream error', data);
      return res.status(502).json({ error: 'TRANSCRIBE_FAILED', detail: (data && data.error && data.error.message) || 'unknown' });
    }
    const text = ((data && data.text) || '').trim();
    res.json({ text });
  } catch (err) {
    console.error('transcribe error', err);
    res.status(500).json({ error: 'SERVER_ERROR' });
  }
});

app.post('/api/extract', requireSession, async (req, res) => {
  try {
    const transcript = req.body && typeof req.body.transcript === 'string' ? req.body.transcript : '';
    if (transcript.trim().length < 3) {
      return res.status(400).json({ error: 'EMPTY_TRANSCRIPT' });
    }

    const today = todayBerlin();
    const weekday = weekdayBerlin();
    const dateTable = relativeDateTable(today);

    const system = `Du bist ein Assistent, der aus dem Transkript eines mündlich diktierten Kundenbesuchsberichts (Deutsch, oft Umgangssprache) strukturierte Daten extrahiert.
Heutiges Datum: ${today} (${weekday}), Zeitzone Europe/Berlin.
Tabelle bekannter relativer Datumsangaben, bereits für dich berechnet (Schlüssel -> YYYY-MM-DD):
${JSON.stringify(dateTable, null, 2)}
Wird im Transkript einer dieser Begriffe (oder ein Wochentag wie "letzten Montag", "am Dienstag") genannt, verwende GENAU den zugehörigen Wert aus der Tabelle für "besuchsdatum" - rechne selbst nichts um. Wird stattdessen ein explizites Datum genannt (z. B. "am 3. März" oder "12.05."), wandle es selbst in YYYY-MM-DD um (fehlt das Jahr, nimm das Jahr von ${today}). Wird gar kein Datum erwähnt, verwende das heutige Datum: ${today}.
Gib ausschließlich ein einziges JSON-Objekt mit genau diesen drei Feldern zurück, ohne jeglichen weiteren Text:
{
  "kunde": string,        // Name des besuchten Kunden/Unternehmens als Freitext. Leerer String, wenn nicht erkennbar.
  "besuchsdatum": string, // Datum des Besuchs im Format YYYY-MM-DD, nach obiger Regel bestimmt.
  "naechsteTodos": string // Kurze, prägnante Zusammenfassung der als Nächstes zu erledigenden Schritte/To-Dos aus dem Text. Leerer String, wenn keine genannt werden.
}
Erfinde keine Informationen, die nicht im Text stehen oder sich nicht daraus ableiten lassen (Ausnahme: der Datumsfallback auf das heutige Datum).`;

    const r = await fetch(`${VIBE_BASE}/chat/completions`, {
      method: 'POST',
      headers: { 'X-Api-Key': VIBE_API_KEY, 'Content-Type': 'application/json' },
      body: JSON.stringify({
        model: AI_MODEL,
        messages: [
          { role: 'system', content: system },
          { role: 'user', content: transcript },
        ],
        response_format: { type: 'json_object' },
        temperature: 0.2,
      }),
    });
    const data = await r.json();
    if (!r.ok) {
      console.error('extract upstream error', data);
      return res.status(502).json({ error: 'EXTRACT_FAILED', detail: (data && data.error && data.error.message) || 'unknown' });
    }

    let parsed;
    try {
      parsed = JSON.parse(data.choices[0].message.content);
    } catch (e) {
      console.error('extract parse error', e, data);
      return res.status(502).json({ error: 'EXTRACT_PARSE_FAILED' });
    }

    const besuchsdatum = typeof parsed.besuchsdatum === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(parsed.besuchsdatum)
      ? parsed.besuchsdatum
      : today;

    res.json({
      kunde: typeof parsed.kunde === 'string' ? parsed.kunde : '',
      besuchsdatum,
      naechsteTodos: typeof parsed.naechsteTodos === 'string' ? parsed.naechsteTodos : '',
    });
  } catch (err) {
    console.error('extract error', err);
    res.status(500).json({ error: 'SERVER_ERROR' });
  }
});

app.post('/api/submit', requireSession, async (req, res) => {
  try {
    const body = req.body || {};
    const bericht = typeof body.bericht === 'string' ? body.bericht.trim() : '';
    const kunde = typeof body.kunde === 'string' ? body.kunde.trim() : '';
    const besuchsdatum = typeof body.besuchsdatum === 'string' ? body.besuchsdatum.trim() : '';
    const naechsteTodos = typeof body.naechsteTodos === 'string' ? body.naechsteTodos.trim() : '';

    if (bericht.length < 3) {
      return res.status(400).json({ error: 'EMPTY_REPORT' });
    }
    if (!/^\d{4}-\d{2}-\d{2}$/.test(besuchsdatum)) {
      return res.status(400).json({ error: 'INVALID_DATE' });
    }

    const title = `Besuchsbericht${kunde ? ' - ' + kunde : ''} (${besuchsdatum})`;

    const itemBody = {
      title,
      [FIELD_BERICHT]: bericht,
      [FIELD_KUNDE]: kunde,
      [FIELD_BESUCHSDATUM]: besuchsdatum,
      [FIELD_TODOS]: naechsteTodos,
    };

    const r = await fetch(`${VIBE_BASE}/items/${ENTITY_TYPE_ID}`, {
      method: 'POST',
      headers: { 'X-Api-Key': VIBE_API_KEY, 'Content-Type': 'application/json' },
      body: JSON.stringify(itemBody),
    });
    const data = await r.json();
    if (!r.ok || data.success === false) {
      console.error('submit upstream error', data);
      return res.status(502).json({ error: 'CRM_CREATE_FAILED', detail: (data && data.error && data.error.message) || 'unknown' });
    }

    res.json({ id: data.data.id });
  } catch (err) {
    console.error('submit error', err);
    res.status(500).json({ error: 'SERVER_ERROR' });
  }
});

app.get('/health', (req, res) => {
  res.json({ ok: true });
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Besuchsbericht-App listening on port ${PORT}`);
});
