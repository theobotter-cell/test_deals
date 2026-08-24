const express = require('express');
const multer = require('multer');
const path = require('path');
const MsgReader = require('@kenjiuno/msgreader').default;
const sanitizeHtml = require('sanitize-html');
const iconv = require('iconv-lite');

const VIBE_BASE_URL = process.env.VIBE_BASE_URL || 'https://vibecode.bitrix24.com/v1';
const VIBE_API_KEY = process.env.VIBE_API_KEY || '';
const PORT = process.env.PORT || 3000;

if (!VIBE_API_KEY) {
  console.error('FATAL: VIBE_API_KEY is not set. Set it via the deploy env (the app\'s own vibe_app_ key) so this server can identify itself to the Vibecode API.');
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

// ---------------------------------------------------------------------------
// Vibecode API client.
//
// Every call carries X-Api-Key (this app's own key). Calls made on behalf of
// the signed-in Bitrix24 employee additionally forward that employee's own
// session token (userAuth), taken from the X-Vibe-Authorization header the
// platform Gateway injects on every request coming through a placement
// iframe. That is what makes the resulting CRM activity's author the actual
// uploading employee instead of a shared service identity.
// ---------------------------------------------------------------------------
async function vibeFetch(pathAndQuery, { method = 'GET', body, userAuth } = {}) {
  await limiter.acquire();
  const url = `${VIBE_BASE_URL}${pathAndQuery}`;
  const headers = { 'X-Api-Key': VIBE_API_KEY };
  if (body !== undefined) headers['Content-Type'] = 'application/json';
  // The Gateway's X-Vibe-Authorization header already carries the "Bearer "
  // prefix — forward it as-is. Guard against double-prefixing if a caller
  // ever passes a bare token instead.
  if (userAuth) headers['Authorization'] = /^bearer\s/i.test(userAuth) ? userAuth : `Bearer ${userAuth}`;

  const res = await fetch(url, {
    method,
    headers,
    body: body !== undefined ? JSON.stringify(body) : undefined,
  });

  if (res.status === 429) {
    const retryAfter = Number(res.headers.get('Retry-After') || '2');
    await new Promise((r) => setTimeout(r, Math.min(retryAfter, 10) * 1000));
    return vibeFetch(pathAndQuery, { method, body, userAuth });
  }

  const json = await res.json().catch(() => ({ success: false, error: { code: 'BAD_JSON', message: 'Non-JSON response from Vibecode API' } }));
  return { ok: res.ok && json.success !== false, status: res.status, json };
}

// ---------------------------------------------------------------------------
// .msg parsing
// ---------------------------------------------------------------------------
const OLE_SIGNATURE = Buffer.from([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]);

class MsgParseError extends Error {}

function parseMsgBuffer(buffer) {
  if (buffer.length < 8 || !buffer.subarray(0, 8).equals(OLE_SIGNATURE)) {
    throw new MsgParseError('This file is not a valid Outlook .msg file (missing the expected file signature). It may be corrupted, renamed, or a different file type.');
  }

  let reader;
  let fileData;
  try {
    reader = new MsgReader(bufferToArrayBuffer(buffer));
    fileData = reader.getFileData();
  } catch (e) {
    throw new MsgParseError(`Could not parse the .msg file — it appears to be corrupted. (${e.message || e})`);
  }

  if (!fileData || fileData.error || fileData.dataType !== 'msg') {
    throw new MsgParseError(`Could not parse the .msg file — it appears to be corrupted or is not a supported Outlook message.${fileData && fileData.error ? ` (${fileData.error})` : ''}`);
  }

  const recipients = fileData.recipients || [];
  const to = recipients.filter((r) => (r.recipType || 'to') === 'to');
  const cc = recipients.filter((r) => r.recipType === 'cc');
  const toDisplay = (to.length ? to : recipients).map(formatRecipient).join('; ') || '(none found)';
  const ccDisplay = cc.map(formatRecipient).join('; ');

  // Internal/Exchange-originated mail often stores the sender as an X.500
  // directory name (e.g. "/O=EXCHANGELABS/OU=.../CN=RECIPIENTS/CN=...")
  // rather than an SMTP address. Prefer the resolved SMTP properties and
  // never surface a raw directory name to the user or to Bitrix24.
  const fromEmail = pickEmail(fileData.senderSmtpAddress, fileData.senderEmail, fileData.sentRepresentingSmtpAddress);
  const fromName = fileData.senderName || null;
  const fromDisplay = fromName && fromEmail ? `${fromName} <${fromEmail}>` : (fromEmail || fromName || 'Unknown sender');

  // "Sent" is when the sender's client submitted the message; delivery time
  // (when it reached the recipient's mailbox) is the best fallback for
  // messages where the submit time wasn't preserved.
  let sentDate = null;
  const dateSource = fileData.clientSubmitTime || fileData.messageDeliveryTime;
  if (dateSource) {
    const d = new Date(dateSource);
    if (!isNaN(d.getTime())) sentDate = d;
  }

  // An attachment referenced by the HTML body via cid: (e.g. an embedded
  // signature logo) is part of the message's layout, not a file the sender
  // meant to hand over — splice it into the body as a data URI instead of
  // listing it as a downloadable attachment.
  // PidTagBodyHtml (string) is the "nice" property, but many writers only
  // populate PidTagHtml (raw bytes in whatever codepage the message declares).
  const bodyHtmlRaw = fileData.bodyHtml || decodeHtmlBytes(fileData.html, fileData.internetCodepage || fileData.messageCodepage) || null;
  const attachmentsMeta = (fileData.attachments || []).filter((a) => !a.innerMsgContent);
  const attachments = [];
  const inlineImages = [];
  for (const meta of attachmentsMeta) {
    const att = reader.getAttachment(meta);
    const fileName = att.fileName || meta.fileNameShort || meta.fileName || 'attachment';
    const content = Buffer.from(att.content);
    const cid = meta.pidContentId;
    const isInline = !!(cid && bodyHtmlRaw && bodyHtmlRaw.includes(cid));
    if (isInline) {
      inlineImages.push({ cid, mimeType: meta.attachMimeTag || guessMimeType(fileName), content });
    } else {
      attachments.push({ fileName, content });
    }
  }

  return {
    subject: fileData.subject || '(no subject)',
    fromDisplay,
    fromEmail,
    toDisplay,
    ccDisplay,
    sentDate,
    bodyText: fileData.body || '',
    bodyHtml: bodyHtmlRaw,
    inlineImages,
    attachments,
  };
}

// Windows codepage ID -> iconv-lite encoding name, for decoding PidTagHtml
// (raw bytes) when the message has no PidTagBodyHtml string property.
const CODEPAGE_MAP = {
  1200: 'utf-16le', 1201: 'utf-16be', 65001: 'utf-8', 20127: 'ascii',
  1250: 'windows-1250', 1251: 'windows-1251', 1252: 'windows-1252', 1253: 'windows-1253',
  1254: 'windows-1254', 1255: 'windows-1255', 1256: 'windows-1256', 1257: 'windows-1257', 1258: 'windows-1258',
  28591: 'iso-8859-1', 932: 'shiftjis', 936: 'gbk', 949: 'euc-kr', 950: 'big5',
};

function decodeHtmlBytes(bytes, codepage) {
  if (!bytes || !bytes.length) return null;
  const encoding = CODEPAGE_MAP[codepage] || 'utf-8';
  try {
    return iconv.decode(Buffer.from(bytes), encoding);
  } catch (e) {
    return Buffer.from(bytes).toString('utf-8');
  }
}

function guessMimeType(fileName) {
  const ext = (fileName.match(/\.([a-z0-9]+)$/i) || [])[1] || '';
  const map = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp', svg: 'image/svg+xml', webp: 'image/webp' };
  return map[ext.toLowerCase()] || 'application/octet-stream';
}

const EMAIL_RE = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

// Returns the first candidate that actually looks like an email address —
// filters out X.500 directory names ("/O=EXCHANGELABS/...") and other junk.
function pickEmail(...candidates) {
  for (const c of candidates) {
    if (c && EMAIL_RE.test(c)) return c;
  }
  return null;
}

function formatRecipient(r) {
  const email = pickEmail(r.smtpAddress, r.email);
  if (r.name && email && r.name !== email) return `${r.name} <${email}>`;
  return email || r.name || 'unknown';
}

function bufferToArrayBuffer(buffer) {
  return buffer.buffer.slice(buffer.byteOffset, buffer.byteOffset + buffer.byteLength);
}

// ---------------------------------------------------------------------------
// Disk: get-or-create a folder to hold email attachments for a deal.
// ---------------------------------------------------------------------------
async function findChildFolder(parentId, name, userAuth) {
  const { ok, json } = await vibeFetch(`/folders?parentId=${encodeURIComponent(parentId)}&limit=500`, { userAuth });
  if (!ok) return null;
  const match = (json.data || []).find((item) => item.type === 'folder' && item.name === name);
  return match ? match.id : null;
}

async function getOrCreateFolder(parentId, name, userAuth) {
  const existing = await findChildFolder(parentId, name, userAuth);
  if (existing) return existing;
  const { ok, json } = await vibeFetch('/folders', { method: 'POST', body: { parentId, name }, userAuth });
  if (!ok) {
    if (json.error && json.error.code === 'ALREADY_EXISTS') {
      const retry = await findChildFolder(parentId, name, userAuth);
      if (retry) return retry;
    }
    throw new Error(`Could not create Disk folder "${name}": ${json.error ? json.error.message : 'unknown error'}`);
  }
  return json.data.id;
}

async function uploadAttachmentsToDisk(dealId, attachments, userAuth) {
  // Prefer the shared Company Drive ("common" storage) so every team member
  // with deal access can open the attachment, not just the uploader's own
  // personal Drive.
  const { ok: commonOk, json: commonJson } = await vibeFetch(
    `/storages?limit=1&filter=${encodeURIComponent(JSON.stringify({ entityType: 'common' }))}`,
    { userAuth }
  );
  let rootFolderId = commonOk && commonJson.data && commonJson.data.length ? commonJson.data[0].rootFolderId : null;

  if (!rootFolderId) {
    const { ok, json } = await vibeFetch('/storages?limit=1', { userAuth });
    if (!ok || !json.data || !json.data.length) {
      console.warn('No Disk storage available for this account; skipping attachment upload to Disk.');
      return [];
    }
    rootFolderId = json.data[0].rootFolderId;
  }

  const emailFolderId = await getOrCreateFolder(rootFolderId, 'Email Attachments', userAuth);
  const dealFolderId = await getOrCreateFolder(emailFolderId, `Deal ${dealId}`, userAuth);

  const uploaded = [];
  for (const att of attachments) {
    const { ok: upOk, json: upJson } = await vibeFetch('/files/upload', {
      method: 'POST',
      body: { folderId: dealFolderId, filename: att.fileName, content: att.content.toString('base64') },
      userAuth,
    });
    if (upOk) {
      // detailUrl opens the file in Bitrix24 Drive under the viewer's own
      // session — safe to store permanently. downloadUrl carries a signed,
      // time-limited auth token and must never be persisted.
      uploaded.push({ id: upJson.data.id, fileName: att.fileName, url: upJson.data.detailUrl });
    } else {
      console.warn(`Failed to upload attachment "${att.fileName}" to Disk:`, upJson.error);
    }
  }
  return uploaded;
}

// ---------------------------------------------------------------------------
// Build the timeline activity description. descriptionType: 3 on a CRM
// activity is real HTML (not BBCode) — the timeline renders it verbatim, so
// line breaks need <br> and every piece of untrusted text (headers, body,
// file names) must be escaped.
// ---------------------------------------------------------------------------
function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}

// The .msg body HTML is untrusted content from an uploaded file, about to be
// rendered as real HTML inside an authenticated Bitrix24 page — sanitize it
// like any other untrusted HTML sink. No scripts, no event handlers, no
// external stylesheets; only tags/attributes/CSS properties an email
// signature or formatted message plausibly needs.
const SANITIZE_OPTIONS = {
  allowedTags: [
    'p', 'br', 'div', 'span', 'b', 'strong', 'i', 'em', 'u', 's', 'a', 'img',
    'table', 'thead', 'tbody', 'tfoot', 'tr', 'td', 'th',
    'ul', 'ol', 'li', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'blockquote', 'hr', 'pre', 'font',
  ],
  allowedAttributes: {
    a: ['href', 'title'],
    img: ['src', 'alt', 'width', 'height'],
    td: ['colspan', 'rowspan', 'align', 'valign'],
    th: ['colspan', 'rowspan', 'align', 'valign'],
    table: ['border', 'cellpadding', 'cellspacing', 'width'],
    font: ['color', 'size', 'face'],
    '*': ['style'],
  },
  allowedSchemes: ['http', 'https', 'mailto'],
  allowedSchemesByTag: { img: ['data', 'http', 'https', 'cid'] },
  allowedStyles: {
    '*': {
      color: [/^#[0-9a-f]{3,8}$/i, /^rgb\(.*\)$/i, /^[a-z]+$/i],
      'background-color': [/^#[0-9a-f]{3,8}$/i, /^rgb\(.*\)$/i, /^[a-z]+$/i],
      'font-weight': [/^(normal|bold|[1-9]00)$/],
      'font-style': [/^(normal|italic)$/],
      'font-size': [/^\d+(\.\d+)?(px|pt|em|%)$/],
      'font-family': [/^[a-z0-9\s,'"-]+$/i],
      'text-align': [/^(left|right|center|justify)$/],
      'text-decoration': [/^(none|underline|line-through)$/],
    },
  },
  transformTags: { a: sanitizeHtml.simpleTransform('a', { target: '_blank', rel: 'noopener noreferrer' }) },
  nonTextTags: ['script', 'style', 'textarea', 'title', 'noscript', 'iframe', 'object', 'embed', 'form'],
};

function escapeRegExp(s) {
  return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

// Splices embedded images (signature logos, etc.) back into the sanitized
// HTML as data URIs, in place of the cid: reference the original message used.
function inlineDataUris(html, inlineImages) {
  let result = html;
  for (const img of inlineImages) {
    if (!img.cid) continue;
    const dataUri = `data:${img.mimeType};base64,${img.content.toString('base64')}`;
    result = result.replace(new RegExp(`cid:${escapeRegExp(img.cid)}`, 'gi'), dataUri);
  }
  return result;
}

function buildDescription(email, uploadedFiles, attachmentsTotal) {
  const parts = [];
  parts.push(`<b>From:</b> ${escapeHtml(email.fromDisplay)}<br>`);
  parts.push(`<b>To:</b> ${escapeHtml(email.toDisplay)}<br>`);
  if (email.ccDisplay) parts.push(`<b>Cc:</b> ${escapeHtml(email.ccDisplay)}<br>`);
  parts.push(`<b>Date:</b> ${escapeHtml(email.sentDate ? email.sentDate.toUTCString() : 'Unknown')}<br>`);
  parts.push('<hr style="margin:10px 0;border:none;border-top:1px solid #ddd;">');

  if (email.bodyHtml) {
    const sanitized = sanitizeHtml(email.bodyHtml, SANITIZE_OPTIONS);
    parts.push(inlineDataUris(sanitized, email.inlineImages));
  } else {
    const bodyText = email.bodyText && email.bodyText.trim() ? email.bodyText.trim() : '(no body text found in this email)';
    // Preserve the email's own line breaks; everything else is plain escaped text.
    parts.push(escapeHtml(bodyText).replace(/\r\n|\r|\n/g, '<br>'));
  }

  if (attachmentsTotal > 0) {
    parts.push('<br><br>');
    parts.push(`<b>Attachments (${attachmentsTotal}):</b><br>`);
    for (const f of uploadedFiles) {
      // detailUrl can contain literal spaces (folder/file names in the path).
      parts.push(`<a href="${escapeHtml(encodeURI(f.url))}">${escapeHtml(f.fileName)}</a><br>`);
    }
    const failed = attachmentsTotal - uploadedFiles.length;
    if (failed > 0) {
      parts.push(`(${failed} attachment${failed === 1 ? '' : 's'} could not be uploaded to Disk)<br>`);
    }
  }

  return parts.join('');
}

// ---------------------------------------------------------------------------
// HTTP API
// ---------------------------------------------------------------------------
const app = express();
app.use(express.static(path.join(__dirname, 'public')));

const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 40 * 1024 * 1024, files: 1 },
  fileFilter(req, file, cb) {
    const name = (file.originalname || '').toLowerCase();
    if (!name.endsWith('.msg')) {
      cb(new MsgParseError('Only .msg files (Outlook email messages) are supported.'));
      return;
    }
    cb(null, true);
  },
});

app.get('/health', (req, res) => res.json({ ok: true }));

app.post('/api/upload-msg', (req, res) => {
  upload.single('file')(req, res, async (err) => {
    if (err) {
      const message = err instanceof MsgParseError ? err.message
        : err.code === 'LIMIT_FILE_SIZE' ? 'This .msg file is too large (max 40 MB).'
        : 'Could not read the uploaded file.';
      return res.status(400).json({ error: { message } });
    }

    const userAuth = req.headers['x-vibe-authorization'];
    if (!userAuth) {
      return res.status(401).json({ error: { message: 'No Bitrix24 session found for this request. Please reopen the app from the deal.' } });
    }

    if (!req.file) {
      return res.status(400).json({ error: { message: 'No file was uploaded.' } });
    }

    const dealId = Number(req.body.dealId);
    if (!dealId || !Number.isInteger(dealId) || dealId <= 0) {
      return res.status(400).json({ error: { message: 'Missing or invalid deal ID. Please reopen the app from the deal\'s timeline.' } });
    }

    // 1. Parse the .msg file. Nothing is written to Bitrix24 until this succeeds.
    let email;
    try {
      email = parseMsgBuffer(req.file.buffer);
    } catch (e) {
      if (e instanceof MsgParseError) {
        return res.status(400).json({ error: { message: e.message } });
      }
      console.error('Unexpected .msg parse failure:', e);
      return res.status(400).json({ error: { message: 'Could not parse the .msg file — it appears to be corrupted.' } });
    }

    try {
      // 2. Best-effort: upload attachments to a per-deal Disk folder.
      const uploadedFiles = email.attachments.length
        ? await uploadAttachmentsToDisk(dealId, email.attachments, userAuth)
        : [];

      // 3. Create the native CRM Email activity on the deal's timeline.
      const description = buildDescription(email, uploadedFiles, email.attachments.length);
      const activityBody = {
        typeId: 4, // email
        ownerTypeId: 2, // deal
        ownerId: dealId,
        subject: `Email: ${email.subject}`,
        description,
        descriptionType: 3, // BBCode/HTML
        direction: 1, // incoming
        completed: true,
        communications: [
          { value: email.fromEmail || email.fromDisplay || 'unknown@unknown.invalid', entityTypeId: 2, entityId: dealId },
        ],
      };
      if (email.sentDate) {
        activityBody.startTime = email.sentDate.toISOString();
        activityBody.endTime = email.sentDate.toISOString();
      }

      const created = await vibeFetch('/activities', { method: 'POST', body: activityBody, userAuth });

      if (!created.ok) {
        console.error('Activity creation failed:', created.json);
        return res.status(502).json({ error: { message: `Could not write the email to the deal's timeline: ${created.json.error ? created.json.error.message : 'unknown error'}` } });
      }

      return res.json({
        success: true,
        activityId: created.json.data.id,
        subject: email.subject,
        attachmentCount: email.attachments.length,
        attachmentsUploaded: uploadedFiles.length,
      });
    } catch (e) {
      console.error('Upload handling failed:', e);
      return res.status(502).json({ error: { message: e.message || 'Unexpected error while saving the email to Bitrix24.' } });
    }
  });
});

app.use((err, req, res, next) => {
  console.error('Unhandled error:', err);
  res.status(500).json({ error: { message: 'Unexpected server error.' } });
});

app.listen(PORT, () => {
  console.log(`Outlook .msg importer listening on :${PORT}`);
});
