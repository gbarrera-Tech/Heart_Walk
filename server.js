'use strict';
const http = require('node:http');
const fs = require('node:fs');
const path = require('node:path');
const crypto = require('node:crypto');
const { DatabaseSync } = require('node:sqlite');
const ExcelJS = require('exceljs');

const PORT = process.env.PORT || 3000;
const DATA_DIR = process.env.DATA_DIR || path.join(__dirname, 'data');
const ADMIN_PASSWORD = process.env.ADMIN_PASSWORD;
const SIZES = ['XS', 'S', 'M', 'L', 'XL', '2XL', '3XL', '4XL'];

fs.mkdirSync(DATA_DIR, { recursive: true });
const db = new DatabaseSync(path.join(DATA_DIR, 'registrations.db'));
db.exec(`CREATE TABLE IF NOT EXISTS registrations (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  full_name TEXT NOT NULL,
  contact TEXT NOT NULL,
  shirt_size TEXT NOT NULL,
  guests INTEGER NOT NULL,
  created_at TEXT NOT NULL DEFAULT (strftime('%Y-%m-%dT%H:%M:%fZ','now'))
)`);
const insert = db.prepare('INSERT INTO registrations (full_name, contact, shirt_size, guests) VALUES (?,?,?,?)');
const selectAll = db.prepare('SELECT * FROM registrations ORDER BY id');

function validate(b) {
  const full_name = String(b.fullName ?? '').trim();
  const contact = String(b.contact ?? '').trim();
  const shirt_size = String(b.shirtSize ?? '').trim();
  const guests = Number(b.guests);
  if (!full_name || full_name.length > 120) return { error: 'Please enter your full name.' };
  if (!contact || contact.length > 60) return { error: 'Please enter your phone number or Employee ID.' };
  if (!SIZES.includes(shirt_size)) return { error: 'Please choose a shirt size.' };
  if (!Number.isInteger(guests) || guests < 0 || guests > 20) return { error: 'Guests must be a whole number from 0 to 20.' };
  return { values: [full_name, contact, shirt_size, guests] };
}

function authorized(req) {
  if (!ADMIN_PASSWORD) return false;
  const m = /^Basic (.+)$/.exec(req.headers.authorization || '');
  if (!m) return false;
  const pass = Buffer.from(m[1], 'base64').toString().split(':').slice(1).join(':');
  const a = crypto.createHash('sha256').update(pass).digest();
  const b = crypto.createHash('sha256').update(ADMIN_PASSWORD).digest();
  return crypto.timingSafeEqual(a, b);
}

function send(res, status, body, type = 'text/plain; charset=utf-8', headers = {}) {
  res.writeHead(status, { 'Content-Type': type, ...headers });
  res.end(body);
}
const json = (res, status, obj) => send(res, status, JSON.stringify(obj), 'application/json');
const esc = (s) => String(s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));

function readBody(req) {
  return new Promise((resolve, reject) => {
    let data = '';
    req.on('data', (c) => { data += c; if (data.length > 10_000) { reject(new Error('too large')); req.destroy(); } });
    req.on('end', () => resolve(data));
    req.on('error', reject);
  });
}

function requireAdmin(req, res) {
  if (!ADMIN_PASSWORD) { send(res, 503, 'Admin is disabled: set the ADMIN_PASSWORD environment variable.'); return false; }
  if (!authorized(req)) {
    send(res, 401, 'Authentication required', 'text/plain', { 'WWW-Authenticate': 'Basic realm="Heart Walk Admin"' });
    return false;
  }
  return true;
}

async function exportXlsx(res) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Registrations');
  ws.columns = [
    { header: 'Full Name', key: 'full_name', width: 30 },
    { header: 'Phone Number / Employee ID', key: 'contact', width: 28 },
    { header: 'Shirt Size', key: 'shirt_size', width: 12 },
    { header: 'Additional Guests', key: 'guests', width: 18 },
    { header: 'Registered (UTC)', key: 'created_at', width: 24 },
  ];
  ws.getRow(1).font = { bold: true };
  ws.views = [{ state: 'frozen', ySplit: 1 }];
  for (const r of selectAll.all()) {
    // contact is stored as text so leading zeros / Employee IDs aren't mangled
    const row = ws.addRow({ ...r, contact: String(r.contact), created_at: r.created_at.replace('T', ' ').slice(0, 19) });
    row.getCell('contact').numFmt = '@';
  }
  const buf = await wb.xlsx.writeBuffer();
  send(res, 200, Buffer.from(buf), 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    { 'Content-Disposition': 'attachment; filename="heart-walk-registrations.xlsx"' });
}

function adminPage() {
  const rows = selectAll.all();
  const guests = rows.reduce((n, r) => n + r.guests, 0);
  const bySize = SIZES.map((s) => [s, rows.filter((r) => r.shirt_size === s).length]).filter(([, n]) => n);
  return `<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Heart Walk Registrations</title><link rel="stylesheet" href="/style.css"></head><body><main class="wide">
<h1>Registrations</h1>
<p class="summary">${rows.length} registered · ${guests} additional guests · ${rows.length + guests} total attendees${bySize.length ? '<br>Shirts: ' + bySize.map(([s, n]) => `${s} × ${n}`).join(', ') : ''}</p>
<p><a class="button" href="/admin/export.xlsx">Download Excel</a></p>
<div class="table-wrap"><table><thead><tr><th>Full name</th><th>Phone / Employee ID</th><th>Shirt</th><th>Guests</th><th>Registered (UTC)</th></tr></thead><tbody>
${rows.map((r) => `<tr><td>${esc(r.full_name)}</td><td>${esc(r.contact)}</td><td>${esc(r.shirt_size)}</td><td>${r.guests}</td><td>${esc(r.created_at.replace('T', ' ').slice(0, 19))}</td></tr>`).join('') || '<tr><td colspan="5">No registrations yet.</td></tr>'}
</tbody></table></div></main></body></html>`;
}

const TYPES = { '.html': 'text/html; charset=utf-8', '.css': 'text/css; charset=utf-8', '.js': 'text/javascript; charset=utf-8' };

const server = http.createServer(async (req, res) => {
  try {
    const { pathname } = new URL(req.url, 'http://x');
    if (req.method === 'POST' && pathname === '/api/register') {
      let body;
      try { body = JSON.parse(await readBody(req)); } catch { return json(res, 400, { error: 'Invalid request.' }); }
      const v = validate(body);
      if (v.error) return json(res, 400, { error: v.error });
      insert.run(...v.values);
      return json(res, 201, { ok: true });
    }
    if (req.method === 'GET' && pathname === '/admin') return requireAdmin(req, res) && send(res, 200, adminPage(), 'text/html; charset=utf-8');
    if (req.method === 'GET' && pathname === '/admin/export.xlsx') return requireAdmin(req, res) && await exportXlsx(res);
    if (req.method === 'GET' && pathname === '/healthz') return send(res, 200, 'ok');
    if (req.method === 'GET') {
      const file = pathname === '/' ? 'index.html' : pathname.slice(1);
      const full = path.join(__dirname, 'public', file);
      if (full.startsWith(path.join(__dirname, 'public') + path.sep) && fs.existsSync(full) && fs.statSync(full).isFile()) {
        return send(res, 200, fs.readFileSync(full), TYPES[path.extname(full)] || 'application/octet-stream');
      }
    }
    send(res, 404, 'Not found');
  } catch (e) {
    console.error(e);
    send(res, 500, 'Server error');
  }
});

if (require.main === module) {
  server.listen(PORT, () => console.log(`Heart Walk registration listening on :${PORT}`));
}
module.exports = { server, validate };
