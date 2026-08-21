/**
 * Local HTTPS dev server for the Office.js POC add-in.
 * Uses office-addin-dev-certs to generate trusted localhost certs.
 *
 * Run:  npm install && npm run install-certs && npm start
 */
'use strict';

const https = require('https');
const http = require('http');
const fs = require('fs');
const path = require('path');
const os = require('os');
const zlib = require('zlib');
const { execSync } = require('child_process');

const PORT = 3000;
const OLLAMA_CHAT_URL = process.env.OLLAMA_CHAT_URL || 'http://localhost:11434/api/chat';
const OLLAMA_MODEL = process.env.OLLAMA_MODEL || 'qwen2.5:7b';
const CERT_DIR = path.join(os.homedir(), '.office-addin-dev-certs');
const CERT_FILE = path.join(CERT_DIR, 'localhost.crt');
const KEY_FILE  = path.join(CERT_DIR, 'localhost.key');

const MIME = {
  '.html': 'text/html',
  '.js':   'application/javascript',
  '.css':  'text/css',
  '.png':  'image/png',
  '.svg':  'image/svg+xml',
  '.xml':  'application/xml',
  '.json': 'application/json',
};

const CHAT_SYSTEM_PROMPT = [
  'You are an Office.js command planner for a PowerPoint add-in.',
  'Convert user requests into JSON actions only.',
  'Return strict JSON with this shape:',
  '{"assistantMessage":"short helpful text","actions":[...]}',
  'Supported actions:',
  '- create_scalar_variable: {type,name,dataType,defaultValue,group}',
  '- create_table_variable: {type,name,columns,rows,group}',
  '- create_computed_variable: {type,name,formula,group}',
  '- insert_variable: {type,name}',
  '- insert_table: {type,name}',
  '- configure_dynamic_table: {type,shapeName,tableVar,fromRow,toRow}',
  '- configure_dynamic_chart: {type,shapeName,tableVar,labelCol,valueCol,titleVar}',
  '- configure_dynamic_image: {type,shapeName,sourceUrl,fitMode}',
  '- run_preview: {type}',
  'Rules:',
  '- Prefer one or more concrete actions when possible.',
  '- If required info is missing, ask via assistantMessage and return actions as [].',
  '- Use variable names that are valid identifiers (A-Za-z_ then A-Za-z0-9_).',
  '- For table rows, return a 2D array of strings/numbers/booleans.',
  '- Never return markdown fences, explanations, or extra keys outside the JSON object.'
].join('\n');

function readJsonBody(req) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    req.on('data', (chunk) => chunks.push(chunk));
    req.on('end', () => {
      try {
        const raw = Buffer.concat(chunks).toString('utf8') || '{}';
        resolve(JSON.parse(raw));
      } catch (e) {
        reject(new Error('Invalid JSON body.'));
      }
    });
    req.on('error', reject);
  });
}

function extractJsonFromText(text) {
  if (!text) return null;
  const trimmed = String(text).trim();
  try {
    return JSON.parse(trimmed);
  } catch (_) {}

  const fenceMatch = trimmed.match(/```(?:json)?\s*([\s\S]*?)\s*```/i);
  if (fenceMatch) {
    try {
      return JSON.parse(fenceMatch[1]);
    } catch (_) {}
  }

  const start = trimmed.indexOf('{');
  const end = trimmed.lastIndexOf('}');
  if (start >= 0 && end > start) {
    try {
      return JSON.parse(trimmed.slice(start, end + 1));
    } catch (_) {}
  }
  return null;
}

function postJson(urlString, payload) {
  return new Promise((resolve, reject) => {
    const url = new URL(urlString);
    const data = JSON.stringify(payload);
    const client = url.protocol === 'https:' ? https : http;

    const req = client.request({
      hostname: url.hostname,
      port: url.port || (url.protocol === 'https:' ? 443 : 80),
      path: `${url.pathname}${url.search}`,
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Content-Length': Buffer.byteLength(data),
      },
      timeout: 120000,
    }, (res) => {
      const chunks = [];
      res.on('data', (chunk) => chunks.push(chunk));
      res.on('end', () => {
        const raw = Buffer.concat(chunks).toString('utf8');
        if (res.statusCode < 200 || res.statusCode >= 300) {
          reject(new Error(`HTTP ${res.statusCode}: ${raw || 'request failed'}`));
          return;
        }
        try {
          resolve(JSON.parse(raw));
        } catch (e) {
          reject(new Error('Non-JSON response from model endpoint.'));
        }
      });
    });

    req.on('timeout', () => req.destroy(new Error('Model request timeout.')));
    req.on('error', reject);
    req.write(data);
    req.end();
  });
}

async function buildChatPlan(message, variables) {
  const userPayload = {
    message,
    existingVariables: Array.isArray(variables) ? variables.slice(0, 200) : [],
  };

  const modelResp = await postJson(OLLAMA_CHAT_URL, {
    model: OLLAMA_MODEL,
    stream: false,
    messages: [
      { role: 'system', content: CHAT_SYSTEM_PROMPT },
      { role: 'user', content: JSON.stringify(userPayload) },
    ],
  });

  const assistant = modelResp && modelResp.message ? modelResp.message : {};
  const parsed = extractJsonFromText(assistant.content);
  if (!parsed || typeof parsed !== 'object') {
    return {
      assistantMessage: 'I could not parse a structured plan. Please try a more specific command.',
      actions: [],
      raw: assistant.content || '',
    };
  }

  return {
    assistantMessage: String(parsed.assistantMessage || 'Planned requested operation.'),
    actions: Array.isArray(parsed.actions) ? parsed.actions : [],
  };
}

// ── PNG generator (no deps) ──────────────────────────────────────────────────
const CRC_TABLE = (() => {
  const t = new Uint32Array(256);
  for (let i = 0; i < 256; i++) {
    let c = i;
    for (let j = 0; j < 8; j++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
    t[i] = c;
  }
  return t;
})();

function crc32(buf) {
  let c = 0xFFFFFFFF;
  for (let i = 0; i < buf.length; i++) c = CRC_TABLE[(c ^ buf[i]) & 0xFF] ^ (c >>> 8);
  return (c ^ 0xFFFFFFFF) >>> 0;
}

function makePng(size, r, g, b) {
  const sig = Buffer.from([137,80,78,71,13,10,26,10]);

  const ihdrData = Buffer.alloc(13);
  ihdrData.writeUInt32BE(size, 0);
  ihdrData.writeUInt32BE(size, 4);
  ihdrData[8] = 8; ihdrData[9] = 2; // 8-bit RGB

  const rows = [];
  for (let y = 0; y < size; y++) {
    const row = Buffer.alloc(1 + size * 3);
    row[0] = 0; // filter: None
    for (let x = 0; x < size; x++) {
      row[1 + x * 3]     = r;
      row[1 + x * 3 + 1] = g;
      row[1 + x * 3 + 2] = b;
    }
    rows.push(row);
  }
  const idatData = zlib.deflateSync(Buffer.concat(rows));

  function chunk(type, data) {
    const t = Buffer.from(type, 'ascii');
    const len = Buffer.alloc(4); len.writeUInt32BE(data.length, 0);
    const crc = Buffer.alloc(4); crc.writeUInt32BE(crc32(Buffer.concat([t, data])), 0);
    return Buffer.concat([len, t, data, crc]);
  }

  return Buffer.concat([sig, chunk('IHDR', ihdrData), chunk('IDAT', idatData), chunk('IEND', Buffer.alloc(0))]);
}

function ensureIcons() {
  const dir = path.join(__dirname, 'assets');
  if (!fs.existsSync(dir)) fs.mkdirSync(dir);
  // Seismic blue #0078D4
  for (const size of [16, 32, 80]) {
    const file = path.join(dir, `icon-${size}.png`);
    if (!fs.existsSync(file)) fs.writeFileSync(file, makePng(size, 0, 120, 212));
  }
}

// ── Cert check ───────────────────────────────────────────────────────────────
function ensureCerts() {
  if (!fs.existsSync(CERT_FILE) || !fs.existsSync(KEY_FILE)) {
    console.log('Dev certs not found. Installing (may prompt for admin)...');
    execSync('npx office-addin-dev-certs install --machine', { stdio: 'inherit' });
  }
}

// ── Start ─────────────────────────────────────────────────────────────────────
function main() {
  ensureCerts();
  ensureIcons();

  const options = {
    cert: fs.readFileSync(CERT_FILE),
    key:  fs.readFileSync(KEY_FILE),
  };

  const server = https.createServer(options, async (req, res) => {
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

    if (req.method === 'OPTIONS') { res.writeHead(204); res.end(); return; }

    if (req.method === 'POST' && req.url === '/api/chat-intent') {
      try {
        const body = await readJsonBody(req);
        const message = String(body.message || '').trim();
        const variables = body.variables;
        if (!message) {
          res.writeHead(400, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify({ error: 'message is required' }));
          return;
        }

        const plan = await buildChatPlan(message, variables);
        res.writeHead(200, { 'Content-Type': 'application/json', 'Cache-Control': 'no-cache' });
        res.end(JSON.stringify({ model: OLLAMA_MODEL, ...plan }));
      } catch (e) {
        res.writeHead(500, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({
          error: 'chat-intent failed',
          detail: e && e.message ? e.message : String(e),
        }));
      }
      return;
    }

    if (req.method === 'GET' && req.url === '/api/health') {
      res.writeHead(200, { 'Content-Type': 'application/json', 'Cache-Control': 'no-cache' });
      res.end(JSON.stringify({ ok: true, model: OLLAMA_MODEL, ollamaUrl: OLLAMA_CHAT_URL }));
      return;
    }

    let urlPath = req.url.split('?')[0];
    if (urlPath === '/') urlPath = '/taskpane.html';

    const filePath = path.resolve(__dirname, '.' + urlPath);
    if (!filePath.startsWith(__dirname)) { res.writeHead(403); res.end('Forbidden'); return; }

    fs.readFile(filePath, (err, data) => {
      if (err) { res.writeHead(404); res.end(`Not found: ${urlPath}`); return; }
      const ct = MIME[path.extname(filePath)] || 'application/octet-stream';
      res.writeHead(200, { 'Content-Type': ct, 'Cache-Control': 'no-cache' });
      res.end(data);
    });
  });

  server.listen(PORT, () => {
    console.log(`\n✓  POC server: https://localhost:${PORT}`);
    console.log('\nSideload steps:');
    console.log('  1. Open PowerPoint');
    console.log('  2. Insert > Get Add-ins > My Add-ins > Upload My Add-in');
    console.log(`  3. Browse to: ${path.join(__dirname, 'manifest.xml')}`);
    console.log('  4. Click "Variables" button in the Home tab ribbon\n');
  });
}

main();
