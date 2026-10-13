'use strict';
const test = require('node:test');
const assert = require('node:assert');
const os = require('node:os');
const fs = require('node:fs');
const path = require('node:path');
const ExcelJS = require('exceljs');

process.env.DATA_DIR = fs.mkdtempSync(path.join(os.tmpdir(), 'hw-'));
process.env.ADMIN_PASSWORD = 'secret';
const { server } = require('./server');

test('register, admin auth, excel export', async (t) => {
  await new Promise((r) => server.listen(0, r));
  t.after(() => server.close());
  const base = `http://127.0.0.1:${server.address().port}`;
  const post = (b) => fetch(base + '/api/register', { method: 'POST', body: JSON.stringify(b) });

  assert.strictEqual((await post({ fullName: '', contact: 'x', shirtSize: 'M', guests: 0 })).status, 400);
  assert.strictEqual((await post({ fullName: 'A', contact: 'x', shirtSize: 'huge', guests: 0 })).status, 400);
  assert.strictEqual((await post({ fullName: 'Ana <b>', contact: '007123', shirtSize: 'M', guests: 2 })).status, 201);

  assert.strictEqual((await fetch(base + '/admin')).status, 401);
  assert.strictEqual((await fetch(base + '/admin/export.xlsx')).status, 401);
  const auth = { Authorization: 'Basic ' + Buffer.from('admin:secret').toString('base64') };
  const html = await (await fetch(base + '/admin', { headers: auth })).text();
  assert.match(html, /Ana &lt;b&gt;/);

  const res = await fetch(base + '/admin/export.xlsx', { headers: auth });
  assert.strictEqual(res.status, 200);
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(Buffer.from(await res.arrayBuffer()));
  const ws = wb.getWorksheet('Registrations');
  assert.strictEqual(ws.getRow(2).getCell(1).value, 'Ana <b>');
  assert.strictEqual(ws.getRow(2).getCell(2).value, '007123');
  assert.strictEqual(ws.getRow(2).getCell(4).value, 2);
});
