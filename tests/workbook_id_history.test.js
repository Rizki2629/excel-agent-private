const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { Readable } = require('node:stream');
const { workbookIdKey, workbookHistoryKey, createRequestHandler } = require('../server');

const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'excelagent-idhist-'));
const CONFIG = path.join(tmp, 'config.json');

// Mulai dari config kosong agar readConfigFile tidak melempar ENOENT.
fs.writeFileSync(CONFIG, JSON.stringify({ chatHistoryByWorkbook: {} }));

function writeConfigFile(config) {
  fs.writeFileSync(CONFIG, JSON.stringify(config, null, 2));
}
function readConfigFile() {
  return JSON.parse(fs.readFileSync(CONFIG, 'utf8'));
}

function request(method, url, body) {
  const req = Readable.from(body === undefined ? [] : [Buffer.from(JSON.stringify(body))]);
  req.method = method;
  req.url = url;
  req.headers = { origin: 'https://localhost:3000' };
  return req;
}

function response() {
  return {
    statusCode: 0, headers: null, body: '',
    writeHead(code, headers) { this.statusCode = code; this.headers = headers; },
    write(chunk) { this.body += chunk; },
    end(chunk) { if (chunk) this.body += chunk; },
    on() { return this; }
  };
}

function handler() {
  return createRequestHandler({ readConfig: readConfigFile, writeConfig: writeConfigFile });
}

async function call(method, url, body) {
  const res = response();
  await handler()(request(method, url, body), res);
  return res;
}

const ID = 'ea_abc123_def456';

test('workbookIdKey stabil terhadap huruf besar/kecil dan spasi', () => {
  const a = workbookIdKey(ID);
  assert.equal(a, workbookIdKey('  ' + ID.toUpperCase() + '  '));
  assert.equal(workbookIdKey(null), null);
  assert.equal(workbookIdKey(''), null);
});

test('riwayat yang disimpan dengan ID terbaca lewat file hasil duplikat', async () => {
  const pathAsli = 'C:\\Data\\RKAS 2026.xlsm';
  const pathDuplikat = 'C:\\Data\\Salinan\\RKAS 2026 - CADANGAN.xlsm';

  // Simpan riwayat dari file asli (dengan ID tertanam).
  const res1 = await call('POST', '/api/history', {
    workbookPath: pathAsli,
    workbookId: ID,
    chatHistory: [
      { role: 'user', content: 'bantu rekap PPN' },
      { role: 'assistant', content: 'Siap, sudah saya hitung.' }
    ]
  });
  assert.equal(JSON.parse(res1.body).ok, true);

  // Buka file duplikat: path lain, ID sama -> riwayat muncul.
  const url = '/api/history?workbookPath=' + encodeURIComponent(pathDuplikat) +
    '&workbookId=' + encodeURIComponent(ID);
  const res2 = await call('GET', url);
  const hist = JSON.parse(res2.body).chatHistory;
  assert.equal(hist.length, 2, 'riwayat harus ikut ke file duplikat');
  assert.equal(hist[0].content, 'bantu rekap PPN');
});

test('Save As tanpa ID tetap memakai perilaku lama berbasis path', async () => {
  const p = 'C:\\Data\\Lama.xlsx';
  await call('POST', '/api/history', {
    workbookPath: p,
    chatHistory: [{ role: 'user', content: 'halo lama' }]
  });

  const res = await call('GET', '/api/history?workbookPath=' + encodeURIComponent(p));
  assert.equal(JSON.parse(res.body).chatHistory.length, 1, 'jalur lama harus tetap jalan');
});

test('riwayat lama tidak hilang saat file berikutnya menyimpan dengan ID', async () => {
  const lama = 'C:\\Data\\Lama.xlsx';
  const histories = {};
  histories[workbookHistoryKey(lama)] = [{ role: 'user', content: 'entri lama' }];
  writeConfigFile({ chatHistoryByWorkbook: histories });

  const baru = 'C:\\Data\\Baru.xlsx';
  await call('POST', '/api/history', {
    workbookPath: baru, workbookId: ID,
    chatHistory: [{ role: 'user', content: 'entri baru' }]
  });

  const stored = readConfigFile().chatHistoryByWorkbook;
  assert.equal(stored[workbookHistoryKey(lama)].length, 1, 'riwayat file lama utuh');
  assert.equal(stored[workbookIdKey(ID)].length, 1, 'riwayat file baru tercatat');
});

test('DELETE menghapus di kedua kunci (ID dan path)', async () => {
  const p = 'C:\\Data\\Hapus.xlsx';
  await call('POST', '/api/history', {
    workbookPath: p, workbookId: ID,
    chatHistory: [{ role: 'user', content: 'akan dihapus' }]
  });
  assert.ok(readConfigFile().chatHistoryByWorkbook[workbookIdKey(ID)]);
  assert.ok(readConfigFile().chatHistoryByWorkbook[workbookHistoryKey(p)]);

  await call('DELETE', '/api/history', { workbookPath: p, workbookId: ID });
  assert.equal(readConfigFile().chatHistoryByWorkbook[workbookIdKey(ID)], undefined);
  assert.equal(readConfigFile().chatHistoryByWorkbook[workbookHistoryKey(p)], undefined);
});

test('taskpane menanam ExcelAgentId lewat custom properties', () => {
  const js = fs.readFileSync(path.join(__dirname, '..', 'taskpane', 'taskpane.js'), 'utf8');
  assert.match(js, /EXCEL_AGENT_ID_PROP = 'ExcelAgentId'/);
  assert.match(js, /properties\.custom/);
  assert.match(js, /getItemOrNullObject\(EXCEL_AGENT_ID_PROP\)/);
  assert.match(js, /custom\.add\(EXCEL_AGENT_ID_PROP/);
  // ID ikut dikirim ke load / save / delete
  assert.match(js, /'&workbookId=' \+ encodeURIComponent\(workbookBinding\.workbookId\)/);
  assert.match(js, /workbookId: workbookBinding\.workbookId \|\| null/);
  assert.match(js, /payload\.workbookId = workbookBinding\.workbookId/);
});
