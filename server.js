const express = require('express');
const crypto = require('crypto');
const path = require('path');

const VIBE_BASE_URL = process.env.VIBE_BASE_URL || 'https://vibecode.bitrix24.com/v1';
const VIBE_API_KEY = process.env.VIBE_API_KEY || '';
const TASK_SCAN_CAP = Number(process.env.TASK_SCAN_CAP || 500);
const PORT = process.env.PORT || 3000;

if (!VIBE_API_KEY) {
  console.error('FATAL: VIBE_API_KEY is not set. Set it via the deploy env so this server can call the Vibecode API as its owner.');
}

// ---------------------------------------------------------------------------
// Rate limiter: Vibecode enforces 10 req/s per key. Stay under it with a
// sliding-window gate shared by every outbound call this process makes.
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

async function vibeFetch(pathAndQuery, options = {}) {
  await limiter.acquire();
  const url = `${VIBE_BASE_URL}${pathAndQuery}`;
  const headers = Object.assign(
    { 'X-Api-Key': VIBE_API_KEY },
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
    return vibeFetch(pathAndQuery, options);
  }

  const json = await res.json().catch(() => ({ success: false, error: { code: 'BAD_JSON', message: 'Non-JSON response' } }));
  return { ok: res.ok && json.success !== false, status: res.status, json };
}

// Query-string builder matching the Vibecode API's per-param conventions:
// filter is a raw JSON blob, select is comma-joined, sort/order use bracket
// notation, everything else is a plain scalar.
function qs(params) {
  const parts = [];
  for (const [k, v] of Object.entries(params)) {
    if (v === undefined || v === null) continue;
    if (k === 'filter') {
      parts.push(`filter=${encodeURIComponent(JSON.stringify(v))}`);
    } else if (k === 'select' && Array.isArray(v)) {
      parts.push(`select=${encodeURIComponent(v.join(','))}`);
    } else if ((k === 'sort' || k === 'order') && v && typeof v === 'object' && !Array.isArray(v)) {
      for (const [field, dir] of Object.entries(v)) {
        parts.push(`${k}[${encodeURIComponent(field)}]=${encodeURIComponent(dir)}`);
      }
    } else if (typeof v === 'object') {
      parts.push(`${encodeURIComponent(k)}=${encodeURIComponent(JSON.stringify(v))}`);
    } else {
      parts.push(`${encodeURIComponent(k)}=${encodeURIComponent(String(v))}`);
    }
  }
  return parts.length ? `?${parts.join('&')}` : '';
}

// ---------------------------------------------------------------------------
// In-memory job store (per-run report generation) + short result cache.
// Not persisted across a redeploy or a sleep/wake cycle - see README.
// ---------------------------------------------------------------------------
const jobs = new Map();
const resultCache = new Map();
const CACHE_TTL_MS = 5 * 60 * 1000;
const JOB_TTL_MS = 30 * 60 * 1000;

setInterval(() => {
  const now = Date.now();
  for (const [id, job] of jobs) {
    if (now - job.createdAt > JOB_TTL_MS) jobs.delete(id);
  }
  for (const [key, entry] of resultCache) {
    if (now - entry.at > CACHE_TTL_MS) resultCache.delete(key);
  }
}, 60 * 1000).unref();

function cacheKey(groupIds, from, to) {
  return JSON.stringify({ g: [...groupIds].sort((a, b) => a - b), from, to });
}

// ---------------------------------------------------------------------------
// Core report logic
// ---------------------------------------------------------------------------
async function fetchAllTasks(groupIds) {
  const filter = { groupId: { $in: groupIds } };
  const { ok, status, json } = await vibeFetch(
    `/tasks${qs({ filter, select: ['id', 'groupId', 'title'], limit: 5000, withTotal: true })}`
  );
  if (!ok) {
    const err = new Error(json.error ? json.error.message : `HTTP ${status}`);
    err.code = json.error ? json.error.code : 'UNKNOWN_ERROR';
    throw err;
  }
  return json.data || [];
}

async function fetchTaskCommentsInRange(taskId, from, to) {
  const filter = { '>=POST_DATE': from, '<=POST_DATE': to };
  const seen = new Set();
  const comments = [];
  let offset = 0;
  let truncated = false;
  for (let page = 0; page < 5; page++) {
    const { ok, status, json } = await vibeFetch(
      `/tasks/${taskId}/comments${qs({ filter, limit: 200, offset, sort: 'id:asc' })}`
    );
    if (!ok) {
      const err = new Error(json.error ? json.error.message : `HTTP ${status}`);
      err.code = json.error ? json.error.code : 'UNKNOWN_ERROR';
      throw err;
    }
    const data = json.data || [];
    let newCount = 0;
    for (const c of data) {
      if (!seen.has(c.id)) {
        seen.add(c.id);
        comments.push(c);
        newCount++;
      }
    }
    if (json.meta && json.meta.truncated) truncated = true;
    if (!json.meta || !json.meta.hasMore) break;
    if (newCount === 0) {
      // offset is not honored on this read path (old-card filtered list) - stop rather than loop forever.
      truncated = true;
      break;
    }
    offset += data.length;
    if (page === 4 && json.meta.hasMore) truncated = true;
  }
  return { comments, truncated };
}

async function resolveUsers(authorIds) {
  const ids = [...authorIds];
  const byId = new Map();
  const chunkSize = 50;
  for (let i = 0; i < ids.length; i += chunkSize) {
    const chunk = ids.slice(i, i + chunkSize);
    const { ok, json } = await vibeFetch(
      `/users${qs({ filter: { id: { $in: chunk } }, select: ['id', 'name', 'lastName', 'personalPhoto'], limit: 50 })}`
    );
    if (ok) {
      for (const u of json.data || []) byId.set(Number(u.id), u);
    }
  }
  return byId;
}

// ---------------------------------------------------------------------------
// HTTP API
// ---------------------------------------------------------------------------
const app = express();
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

app.get('/health', (req, res) => res.json({ ok: true }));

app.get('/api/workgroups', async (req, res) => {
  const search = (req.query.search || '').toString().trim();
  const filter = { archived: 'N' };
  if (search) filter['%name'] = search;
  try {
    const { ok, status, json } = await vibeFetch(
      `/workgroups${qs({ filter, select: ['id', 'name', 'isProject'], sort: { name: 'ASC' }, limit: 500 })}`
    );
    if (!ok) return res.status(status || 502).json({ error: json.error || { message: 'Failed to list workgroups' } });
    res.json({ data: json.data || [], total: json.meta ? json.meta.total : undefined });
  } catch (e) {
    res.status(502).json({ error: { message: e.message } });
  }
});

app.post('/api/report', (req, res) => {
  const { groupIds, from, to, force } = req.body || {};
  if (!Array.isArray(groupIds) || groupIds.length === 0) {
    return res.status(400).json({ error: { code: 'NO_GROUPS', message: 'Select at least one project group.' } });
  }
  const fromIso = from ? new Date(from).toISOString() : null;
  const toIso = to ? new Date(to).toISOString() : null;
  if (!fromIso || !toIso || fromIso > toIso) {
    return res.status(400).json({ error: { code: 'BAD_RANGE', message: 'Invalid date range.' } });
  }

  const cached = resultCache.get(cacheKey(groupIds, fromIso, toIso));
  if (cached) {
    return res.json({ jobId: null, cached: true, result: cached.result });
  }

  const jobId = crypto.randomUUID();
  const job = { id: jobId, status: 'starting', createdAt: Date.now(), totalTasks: 0, scannedTasks: 0 };
  jobs.set(jobId, job);

  (async () => {
    try {
      job.status = 'fetching-tasks';
      const tasks = await fetchAllTasks(groupIds);
      if (tasks.length > TASK_SCAN_CAP && !force) {
        job.status = 'needs-confirmation';
        job.confirmation = { taskCount: tasks.length, cap: TASK_SCAN_CAP };
        return;
      }
      job.totalTasks = tasks.length;
      job.status = 'scanning-comments';
      // re-run through the shared runner, reusing the already-fetched task list
      await runReportWithTasks(job, tasks, groupIds, fromIso, toIso);
    } catch (e) {
      job.status = 'error';
      job.error = { code: e.code || 'UNKNOWN_ERROR', message: e.message || 'Unknown error' };
    }
  })();

  res.json({ jobId, cached: false });
});

async function runReportWithTasks(job, tasks, groupIds, from, to) {
  const tally = new Map();
  let failedTasks = 0;
  let truncatedTasks = 0;
  const CONCURRENCY = 6;
  let cursor = 0;
  async function worker() {
    while (cursor < tasks.length) {
      const idx = cursor++;
      const task = tasks[idx];
      try {
        const { comments, truncated } = await fetchTaskCommentsInRange(task.id, from, to);
        if (truncated) truncatedTasks++;
        for (const c of comments) {
          const authorId = Number(c.authorId);
          if (!tally.has(authorId)) tally.set(authorId, { commentCount: 0, taskIds: new Set() });
          const entry = tally.get(authorId);
          entry.commentCount++;
          entry.taskIds.add(task.id);
        }
      } catch (e) {
        failedTasks++;
      }
      job.scannedTasks++;
    }
  }
  await Promise.all(Array.from({ length: Math.min(CONCURRENCY, tasks.length || 1) }, worker));

  job.status = 'resolving-users';
  const authorIds = [...tally.keys()];
  const users = await resolveUsers(authorIds);

  const rows = authorIds.map((authorId) => {
    const entry = tally.get(authorId);
    const u = users.get(authorId);
    return {
      userId: authorId,
      userName: u ? `${u.name || ''} ${u.lastName || ''}`.trim() || `User ${authorId}` : `User ${authorId} (not found)`,
      avatar: u ? u.personalPhoto || null : null,
      commentCount: entry.commentCount,
      taskCount: entry.taskIds.size,
    };
  });
  rows.sort((a, b) => b.commentCount - a.commentCount);

  job.result = {
    rows,
    summary: {
      totalTasks: tasks.length,
      scannedTasks: job.scannedTasks,
      failedTasks,
      truncatedTasks,
      totalComments: rows.reduce((s, r) => s + r.commentCount, 0),
    },
  };
  job.status = 'done';
  resultCache.set(cacheKey(groupIds, from, to), { at: Date.now(), result: job.result });
}

app.post('/api/report/:jobId/confirm', (req, res) => {
  const job = jobs.get(req.params.jobId);
  if (!job || job.status !== 'needs-confirmation') {
    return res.status(404).json({ error: { message: 'No pending confirmation for this job.' } });
  }
  const { groupIds, from, to } = req.body || {};
  job.status = 'fetching-tasks';
  (async () => {
    try {
      const tasks = await fetchAllTasks(groupIds);
      job.totalTasks = tasks.length;
      job.status = 'scanning-comments';
      await runReportWithTasks(job, tasks, groupIds, from, to);
    } catch (e) {
      job.status = 'error';
      job.error = { code: e.code || 'UNKNOWN_ERROR', message: e.message || 'Unknown error' };
    }
  })();
  res.json({ ok: true });
});

app.get('/api/report/:jobId', (req, res) => {
  const job = jobs.get(req.params.jobId);
  if (!job) return res.status(404).json({ error: { message: 'Job not found (it may have expired).' } });
  res.json({
    status: job.status,
    totalTasks: job.totalTasks,
    scannedTasks: job.scannedTasks,
    confirmation: job.confirmation,
    error: job.error,
    result: job.status === 'done' ? job.result : undefined,
  });
});

app.listen(PORT, () => {
  console.log(`Task Comment Activity app listening on :${PORT}`);
});
