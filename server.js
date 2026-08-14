/**
 * Local HTTPS dev server for the Office.js POC add-in.
 * Uses office-addin-dev-certs to generate trusted localhost certs.
 *
 * Run:  npm install && npm run install-certs && npm start
 */
'use strict';

const https = require('https');
const fs = require('fs');
const path = require('path');
const os = require('os');
const zlib = require('zlib');
const { execSync } = require('child_process');

const PORT = 3000;
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

  const server = https.createServer(options, (req, res) => {
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'GET, OPTIONS');

    if (req.method === 'OPTIONS') { res.writeHead(204); res.end(); return; }

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
