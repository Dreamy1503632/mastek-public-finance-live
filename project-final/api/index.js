const express = require('express');
const cors = require('cors');
const RateLimit = require('express-rate-limit');
const fetch = (...args) => import('node-fetch').then(({ default: f }) => f(...args));
const path = require('path');
const XLSX = require('xlsx');
const nodemailer = require('nodemailer');
const { put, head } = require('@vercel/blob');

const app = express();

var limiter = RateLimit({
  windowMs: 15 * 60 * 1000,
  max: 100,
});
app.use(limiter);
app.use(cors());
app.use(express.json());

// --- Azure OpenAI config: environment variables (set in Vercel dashboard),
// never committed to the repo. ---
// Accept a plain resource root ("https://xyz.openai.azure.com") for
// AZURE_OPENAI_ENDPOINT, but also tolerate someone pasting the FULL
// request URL Azure shows in its portal (which already includes
// /openai/deployments/<name>/chat/completions?api-version=...) by
// stripping everything after the host so we don't double up the path.
function normalizeAzureEndpoint(raw) {
  if (!raw) return raw;
  try {
    const u = new URL(raw);
    return `${u.protocol}//${u.host}`;
  } catch {
    return raw;
  }
}

// If a full URL was pasted into AZURE_OPENAI_ENDPOINT, pull the
// deployment name and api-version out of it too, so config still works
// even if only the endpoint var was set.
function extractFromFullUrl(raw) {
  const result = { deployment: null, apiVersion: null };
  if (!raw) return result;
  try {
    const u = new URL(raw);
    const match = u.pathname.match(/\/deployments\/([^/]+)/);
    if (match) result.deployment = decodeURIComponent(match[1]);
    const v = u.searchParams.get('api-version');
    if (v) result.apiVersion = v;
  } catch {
    // not a full URL, nothing to extract
  }
  return result;
}

const rawEndpoint = process.env.AZURE_OPENAI_ENDPOINT;
const extracted = extractFromFullUrl(rawEndpoint);

const azureConfig = {
  endpoint: normalizeAzureEndpoint(rawEndpoint),
  // Support either env var name - AZURE_OPENAI_KEY is the canonical one
  // used elsewhere in this file, but AZURE_OPENAI_API_KEY (the name Azure
  // itself shows in the portal) is accepted too so a copy-paste from
  // Azure's own docs still works.
  apiKey: process.env.AZURE_OPENAI_KEY || process.env.AZURE_OPENAI_API_KEY,
  deployment: process.env.AZURE_OPENAI_DEPLOYMENT || extracted.deployment || 'gpt-4o',
  apiVersion: process.env.AZURE_OPENAI_API_VERSION || extracted.apiVersion || '2024-08-01-preview',
};

if (!azureConfig.apiKey) {
  console.warn('AZURE_OPENAI_KEY (or AZURE_OPENAI_API_KEY) is not set - report generation will fail until it is configured.');
}

// --- SMTP/email config: also environment variables, same pattern as Azure.
// Set these in the Vercel dashboard: EMAIL_HOST, EMAIL_PORT, EMAIL_SECURE,
// EMAIL_USER, EMAIL_PASS, EMAIL_FROM, EMAIL_FROM_NAME. ---
const emailConfig = {
  host: process.env.EMAIL_HOST,
  port: Number(process.env.EMAIL_PORT || 587),
  secure: process.env.EMAIL_SECURE === 'true',
  user: process.env.EMAIL_USER,
  pass: process.env.EMAIL_PASS,
  from: process.env.EMAIL_FROM,
  fromName: process.env.EMAIL_FROM_NAME || 'Mastek AI',
};

let emailTransporter = null;
if (emailConfig.host && emailConfig.user && emailConfig.pass) {
  emailTransporter = nodemailer.createTransport({
    host: emailConfig.host,
    port: emailConfig.port,
    secure: emailConfig.secure,
    auth: { user: emailConfig.user, pass: emailConfig.pass },
  });
} else {
  console.warn('Email is not fully configured - /api/email-report will fail until EMAIL_HOST/EMAIL_USER/EMAIL_PASS are set.');
}

const reportLimiter = RateLimit({
  windowMs: 10 * 60 * 1000,
  max: 20,
});

const emailLimiter = RateLimit({
  windowMs: 10 * 60 * 1000,
  max: 20,
});

// ============================================
// SHARED VALIDATION
// Kept identical to the blocklist used client-side in assessment.html and
// book-session.html - this is the server-side backstop since client-side
// checks alone can be bypassed.
// ============================================
function isValidEmail(email) {
  return typeof email === 'string' && /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}

const BLOCKED_EMAIL_DOMAINS = [
  'gmail.com', 'googlemail.com', 'yahoo.com', 'yahoo.co.uk', 'ymail.com',
  'outlook.com', 'hotmail.com', 'hotmail.co.uk', 'live.com', 'msn.com',
  'icloud.com', 'me.com', 'mac.com', 'aol.com', 'aim.com',
  'protonmail.com', 'proton.me', 'gmx.com', 'gmx.co.uk', 'mail.com',
  'zoho.com', 'yandex.com', 'inbox.com', 'fastmail.com',
];

function isCompanyEmail(email) {
  const domain = email.split('@')[1]?.toLowerCase().trim();
  return !!domain && !BLOCKED_EMAIL_DOMAINS.includes(domain);
}

function escapeHtml(str) {
  return String(str || '').replace(/[&<>"']/g, (c) => ({
    '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;',
  }[c]));
}

// ============================================
// LEAD / BOOKING STORAGE (Vercel Blob, JSONL format)
// One JSON object per line, appended on every submission - cheap, safe
// for near-simultaneous writes, and trivially small at ~100 entries.
// Convert to .xlsx on demand via /api/export-leads and
// /api/export-bookings whenever you actually want the spreadsheet.
// ============================================
async function readJsonlBlob(filename) {
  try {
    const meta = await head(filename).catch(() => null);
    if (!meta) return [];
    const res = await fetch(meta.url);
    if (!res.ok) return [];
    const text = await res.text();
    return text
      .split('\n')
      .map((line) => line.trim())
      .filter(Boolean)
      .map((line) => {
        try {
          return JSON.parse(line);
        } catch {
          return null;
        }
      })
      .filter(Boolean);
  } catch (error) {
    console.error(`Failed to read ${filename} from Blob:`, error);
    return [];
  }
}

async function appendJsonlBlob(filename, entry) {
  const existingRows = await readJsonlBlob(filename);
  existingRows.push(entry);
  const body = existingRows.map((row) => JSON.stringify(row)).join('\n') + '\n';

  await put(filename, body, {
    access: 'public',
    addRandomSuffix: false,
    allowOverwrite: true,
    contentType: 'application/x-ndjson',
  });

  return existingRows;
}

const LEADS_FILE = 'leads.jsonl';
const BOOKINGS_FILE = 'bookings.jsonl';

app.post('/api/save-lead', async (req, res) => {
  const { firstName, lastName, email, organisation, answers } = req.body || {};

  if (!firstName || !firstName.trim() || !lastName || !lastName.trim()) {
    return res.status(400).json({ error: 'First and last name are required.' });
  }
  if (!isValidEmail(email)) {
    return res.status(400).json({ error: 'A valid work email address is required.' });
  }
  if (!isCompanyEmail(email)) {
    return res.status(400).json({ error: 'Please use your company/work email address, not a personal one (e.g. Gmail, Yahoo, Outlook).' });
  }

  const lead = {
    Timestamp: new Date().toISOString(),
    'First Name': firstName.trim(),
    'Last Name': lastName.trim(),
    Email: email.trim(),
    Organisation: (organisation || '').trim(),
    'Q1 - AI Priority': answers?.[0] || '',
    'Q2 - Budget Visibility': answers?.[1] || '',
    'Q3 - S151 Risk': answers?.[2] || '',
    'Q4 - Team Capacity': answers?.[3] || '',
    'Q5 - ERP Resilience': answers?.[4] || '',
    'Q6 - LGR Readiness': answers?.[5] || '',
  };

  try {
    await appendJsonlBlob(LEADS_FILE, lead);
    res.json({ message: 'Lead saved.' });
  } catch (error) {
    console.error('Failed to save lead:', error);
    res.status(500).json({ error: 'Failed to save lead.' });
  }
});

app.post('/api/save-booking', async (req, res) => {
  const { firstName, lastName, email, phone, organisation, message } = req.body || {};

  if (!firstName || !firstName.trim() || !lastName || !lastName.trim()) {
    return res.status(400).json({ error: 'First and last name are required.' });
  }
  if (!isValidEmail(email)) {
    return res.status(400).json({ error: 'A valid work email address is required.' });
  }
  if (!isCompanyEmail(email)) {
    return res.status(400).json({ error: 'Please use your company/work email address, not a personal one (e.g. Gmail, Yahoo, Outlook).' });
  }
  if (!organisation || !organisation.trim()) {
    return res.status(400).json({ error: 'Organisation is required.' });
  }

  const booking = {
    Timestamp: new Date().toISOString(),
    'First Name': firstName.trim(),
    'Last Name': lastName.trim(),
    Email: email.trim(),
    Phone: (phone || '').trim(),
    Organisation: organisation.trim(),
    Message: (message || '').trim(),
  };

  try {
    await appendJsonlBlob(BOOKINGS_FILE, booking);
    res.json({ message: 'Booking saved.' });
  } catch (error) {
    console.error('Failed to save booking:', error);
    res.status(500).json({ error: 'Failed to save booking.' });
  }
});

// ============================================
// EXCEL EXPORT (on demand, built fresh from the JSONL data)
// Visit /api/export-leads or /api/export-bookings any time - e.g. right
// after the event - to download the familiar .xlsx file.
// ============================================
function rowsToXlsxResponse(res, rows, sheetName, downloadName) {
  const workbook = XLSX.utils.book_new();
  const sheet = XLSX.utils.json_to_sheet(rows);
  XLSX.utils.book_append_sheet(workbook, sheet, sheetName);
  const buffer = XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' });

  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', `attachment; filename="${downloadName}"`);
  res.send(buffer);
}

app.get('/api/export-leads', async (req, res) => {
  try {
    const rows = await readJsonlBlob(LEADS_FILE);
    rowsToXlsxResponse(res, rows, 'Leads', 'leads.xlsx');
  } catch (error) {
    console.error('Failed to export leads:', error);
    res.status(500).json({ error: 'Failed to export leads.' });
  }
});

app.get('/api/export-bookings', async (req, res) => {
  try {
    const rows = await readJsonlBlob(BOOKINGS_FILE);
    rowsToXlsxResponse(res, rows, 'Bookings', 'bookings.xlsx');
  } catch (error) {
    console.error('Failed to export bookings:', error);
    res.status(500).json({ error: 'Failed to export bookings.' });
  }
});

// ============================================
// AI REPORT GENERATION (Azure OpenAI / GPT-4o)
// assessment.html builds the full prompt client-side and posts it here in
// the same {messages:[...]} shape it originally sent straight to
// Anthropic. We forward that prompt to Azure and hand back a response in
// the same {content:[{type:'text', text}]} shape, so the front-end's
// existing parsing code needs no changes.
// ============================================
app.post('/api/generate-report', reportLimiter, async (req, res) => {
  if (!azureConfig.apiKey || !azureConfig.endpoint) {
    return res.status(500).json({ error: 'Azure OpenAI is not configured. Set AZURE_OPENAI_ENDPOINT and AZURE_OPENAI_KEY.' });
  }

  const { messages, max_tokens } = req.body || {};
  const userPrompt = messages?.[0]?.content;

  if (!userPrompt) {
    return res.status(400).json({ error: 'Missing prompt in request body.' });
  }

  const url =
    `${azureConfig.endpoint}/openai/deployments/${azureConfig.deployment}/chat/completions` +
    `?api-version=${azureConfig.apiVersion}`;

  const systemMessage = {
    role: 'system',
    content: 'You output ONLY valid JSON matching the schema the user asks for. No markdown fences, no preamble, no commentary.',
  };
  const userMessage = { role: 'user', content: userPrompt };
  const tokenLimit = Math.min(max_tokens || 700, 1000);

  // Newer reasoning-tier models (e.g. gpt-5.x deployments) use
  // max_completion_tokens instead of max_tokens, and some reject a custom
  // temperature or response_format. Try the "standard" chat-completions
  // shape first; if Azure rejects it for an unsupported-parameter reason,
  // retry once with the adjusted shape rather than failing outright.
  async function callAzure(body) {
    return fetch(url, {
      method: 'POST',
      headers: {
        'api-key': azureConfig.apiKey,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify(body),
    });
  }

  const primaryBody = {
    messages: [systemMessage, userMessage],
    max_tokens: tokenLimit,
    temperature: 0.7,
    response_format: { type: 'json_object' },
  };

  try {
    let azureResponse = await callAzure(primaryBody);
    let errText = '';

    if (!azureResponse.ok) {
      errText = await azureResponse.text();
      const looksLikeUnsupportedParam =
        azureResponse.status === 400 &&
        /max_tokens|temperature|response_format|unsupported/i.test(errText);

      if (looksLikeUnsupportedParam) {
        console.warn('Azure rejected standard params, retrying with newer-model-compatible body:', errText);
        const fallbackBody = {
          messages: [systemMessage, userMessage],
          max_completion_tokens: tokenLimit,
        };
        azureResponse = await callAzure(fallbackBody);
        if (!azureResponse.ok) {
          errText = await azureResponse.text();
        }
      }
    }

    if (!azureResponse.ok) {
      console.error('Azure OpenAI error:', azureResponse.status, errText);
      return res.status(502).json({ error: 'Azure OpenAI request failed.', detail: errText });
    }

    const data = await azureResponse.json();
    const text = data.choices?.[0]?.message?.content || '';

    if (!text) {
      console.error('Azure OpenAI returned no content:', JSON.stringify(data));
      return res.status(502).json({ error: 'Azure OpenAI returned an empty response.' });
    }

    res.json({ content: [{ type: 'text', text }] });
  } catch (error) {
    console.error('Report generation failed:', error);
    res.status(500).json({ error: 'Report generation failed.' });
  }
});

// ============================================
// EMAIL THE REPORT (nodemailer over SMTP, creds from env vars)
// ============================================
function buildReportEmailHtml({ firstName, organisation, report }) {
  const name = escapeHtml(firstName || 'there');
  const org = escapeHtml(organisation || 'your organisation');
  const r = report || {};
  const score = Number.isFinite(r.score) ? r.score : '-';
  const band = escapeHtml(r.band || '');
  const headline = escapeHtml(r.headline || '');
  const desc = escapeHtml(r.desc || '');

  const findings = Array.isArray(r.executiveSummary) ? r.executiveSummary : [];
  const findingsHtml = findings.map(f => `
    <tr><td style="padding:6px 0;border-left:3px solid #C9A84C;padding-left:12px;font-size:13px;color:#333;">
      ${escapeHtml(f.icon || '')} ${escapeHtml(f.text || '')}
    </td></tr>`).join('');

  const opps = Array.isArray(r.opportunities) ? r.opportunities : [];
  const oppsHtml = opps.map((o, i) => `
    <tr><td style="padding:14px 16px;background:#F7F9FC;border-radius:10px;display:block;margin-bottom:10px;">
      <div style="font-size:11px;font-weight:700;color:#C9A84C;">0${i + 1}</div>
      <div style="font-size:15px;font-weight:700;color:#0B1F3A;">${escapeHtml(o.title || '')}</div>
      <div style="font-size:13px;color:#555;margin-top:4px;">${escapeHtml(o.desc || '')}</div>
    </td></tr>`).join('');

  const quickwin = r.quickwin || {};

  return `
  <div style="font-family:Arial,Helvetica,sans-serif;max-width:600px;margin:0 auto;color:#0B1F3A;">
    <div style="background:#0B1F3A;padding:24px 28px;border-radius:12px 12px 0 0;">
      <div style="color:#C9A84C;font-size:12px;font-weight:700;letter-spacing:1px;text-transform:uppercase;">Your Diagnostic Result</div>
      <div style="color:#fff;font-size:22px;font-weight:700;margin-top:6px;">Your Customised AI Assessment Report</div>
      <div style="color:#8A9BB0;font-size:13px;margin-top:4px;">Hi ${name}, based on your responses for ${org}</div>
    </div>
    <div style="background:#fff;padding:28px;border:1px solid #eee;border-top:none;">
      <table width="100%" cellpadding="0" cellspacing="0"><tr>
        <td style="width:90px;font-size:32px;font-weight:900;color:#0B1F3A;vertical-align:top;">${score}<div style="font-size:10px;color:#888;font-weight:400;">out of 100</div></td>
        <td>
          <div style="display:inline-block;font-size:12px;font-weight:700;color:#1A7A45;background:#E8F5EE;padding:4px 12px;border-radius:20px;">${band}</div>
          <div style="font-size:16px;font-weight:700;margin-top:8px;">${headline}</div>
          <div style="font-size:13px;color:#555;margin-top:6px;">${desc}</div>
        </td>
      </tr></table>

      <h3 style="font-size:12px;letter-spacing:1px;text-transform:uppercase;color:#888;margin:24px 0 10px;">Key Findings</h3>
      <table width="100%" cellpadding="0" cellspacing="0">${findingsHtml}</table>

      <h3 style="font-size:12px;letter-spacing:1px;text-transform:uppercase;color:#888;margin:24px 0 10px;">Your Top AI Opportunities</h3>
      <table width="100%" cellpadding="0" cellspacing="8">${oppsHtml}</table>

      <h3 style="font-size:12px;letter-spacing:1px;text-transform:uppercase;color:#888;margin:24px 0 10px;">Your 30-Day Quick Win</h3>
      <div style="background:#F7F9FC;border-radius:10px;padding:16px;">
        <div style="font-size:15px;font-weight:700;">${escapeHtml(quickwin.title || '')}</div>
        <div style="font-size:13px;color:#555;margin-top:6px;">${escapeHtml(quickwin.desc || '')}</div>
      </div>

      <p style="font-size:12px;color:#999;margin-top:28px;">Council AI Finance Opportunity Diagnostic - Public Finance Live 2026</p>
    </div>
  </div>`;
}

app.post('/api/email-report', emailLimiter, async (req, res) => {
  const { firstName, lastName, email, organisation, report } = req.body || {};

  if (!isValidEmail(email)) {
    return res.status(400).json({ error: 'A valid email address is required.' });
  }
  if (!emailTransporter) {
    return res.status(500).json({ error: 'Email is not configured. Set EMAIL_HOST, EMAIL_USER, and EMAIL_PASS.' });
  }

  try {
    const html = buildReportEmailHtml({ firstName, lastName, organisation, report });
    await emailTransporter.sendMail({
      from: `"${emailConfig.fromName}" <${emailConfig.from}>`,
      to: email.trim(),
      subject: 'Your Council AI Assessment Report',
      html,
    });
    res.json({ message: 'Report emailed.' });
  } catch (error) {
    console.error('Failed to send report email:', error);
    res.status(500).json({ error: 'Failed to send email.' });
  }
});

// ============================================
// PAGES
// Vercel serves files in the project root as static assets automatically;
// these routes are kept so the same URLs you already use keep working.
// ============================================
const ROOT = path.join(__dirname, '..');

app.get('/', (req, res) => {
  res.sendFile(path.join(ROOT, 'avatar.html'));
});

app.get('/avatar.html', (req, res) => {
  res.sendFile(path.join(ROOT, 'avatar.html'));
});

app.get('/assessment.html', (req, res) => {
  res.sendFile(path.join(ROOT, 'assessment.html'));
});

app.get('/book-session.html', (req, res) => {
  res.sendFile(path.join(ROOT, 'book-session.html'));
});

module.exports = app;
