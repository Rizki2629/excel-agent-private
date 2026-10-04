const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const { EventEmitter } = require('node:events');
const { Readable } = require('node:stream');
const {
  createRequestHandler,
  registerExecution,
  unregisterExecution,
  cancelExecutions
} = require('../server');

function request(method, url, body) {
  const req = Readable.from(body === undefined ? [] : [Buffer.from(JSON.stringify(body))]);
  req.method = method;
  req.url = url;
  req.headers = {};
  return req;
}

function response() {
  return {
    statusCode: 0,
    headers: null,
    body: '',
    headersSent: false,
    writableEnded: false,
    writeHead(code, headers) { this.statusCode = code; this.headers = headers; this.headersSent = true; },
    write(chunk) { this.body += chunk; },
    end(chunk) { if (chunk) this.body += chunk; this.writableEnded = true; },
    on() { return this; }
  };
}

// Proses palsu yang mencatat apakah dirinya dibunuh.
function fakeChild(pid) {
  const c = new EventEmitter();
  c.pid = pid || 1234;
  c.killed = false;
  c.cancelled = false;
  c.kill = function () { c.killed = true; };
  c.stdin = { on() { }, write() { }, end() { } };
  return c;
}

test('cancel: registry menghentikan proses milik workbook yang dituju saja', () => {
  const a = fakeChild(101);
  const b = fakeChild(102);
  registerExecution('wb-satu', a);
  registerExecution('wb-dua', b);

  const stopped = cancelExecutions('wb-satu');
  assert.equal(stopped, 1);
  assert.equal(a.cancelled, true, 'proses target harus ditandai batal');
  assert.equal(b.cancelled, false, 'proses workbook lain tidak boleh tersentuh');

  // bersihkan
  cancelExecutions('wb-dua');
});

test('cancel: tanpa key menghentikan semua proses yang terdaftar', () => {
  const a = fakeChild(201);
  const b = fakeChild(202);
  registerExecution('wb-a', a);
  registerExecution('wb-b', b);

  const stopped = cancelExecutions('');
  assert.equal(stopped, 2);
  assert.equal(a.cancelled, true);
  assert.equal(b.cancelled, true);
});

test('cancel: unregister mencegah proses selesai ikut dibunuh', () => {
  const c = fakeChild(301);
  registerExecution('wb-x', c);
  unregisterExecution('wb-x', c);

  const stopped = cancelExecutions('wb-x');
  assert.equal(stopped, 0);
  assert.equal(c.cancelled, false);
});

test('cancel: membatalkan saat tidak ada proses tidak error', () => {
  const stopped = cancelExecutions('workbook-tidak-ada');
  assert.equal(stopped, 0);
});

test('cancel: registry mengabaikan argumen kosong', () => {
  assert.doesNotThrow(() => registerExecution('', fakeChild()));
  assert.doesNotThrow(() => registerExecution('k', null));
  assert.doesNotThrow(() => unregisterExecution('tidak-ada', fakeChild()));
});

test('endpoint /api/cancel melaporkan tidak ada proses berjalan', async () => {
  const handler = createRequestHandler();
  const res = response();
  await handler(request('POST', '/api/cancel', {}), res);
  assert.equal(res.statusCode, 200);
  const data = JSON.parse(res.body);
  assert.equal(data.ok, true);
  assert.equal(data.stopped, 0);
  assert.match(data.message, /Tidak ada eksekusi/);
});

test('endpoint /api/cancel menghentikan proses milik workbook tertentu', async () => {
  const { workbookHistoryKey } = require('../server');
  const wp = 'C:\\Data\\Batal.xlsm';
  const child = fakeChild(401);
  registerExecution(workbookHistoryKey(wp), child);

  const handler = createRequestHandler();
  const res = response();
  await handler(request('POST', '/api/cancel', { workbookPath: wp }), res);
  assert.equal(res.statusCode, 200);
  const data = JSON.parse(res.body);
  assert.equal(data.ok, true);
  assert.equal(data.stopped, 1);
  assert.equal(child.cancelled, true);
  assert.match(data.message, /dihentikan/i);
});

test('endpoint /api/cancel menerima body kosong tanpa crash', async () => {
  const handler = createRequestHandler();
  const res = response();
  await handler(request('POST', '/api/cancel'), res);
  assert.equal(res.statusCode, 200);
  assert.equal(JSON.parse(res.body).ok, true);
});

test('exec-python mengembalikan status cancelled, bukan error, saat dihentikan', async () => {
  // execFile palsu: simulasikan proses yang dibunuh user di tengah jalan.
  const execFile = (cmd, args, opts, cb) => {
    const child = fakeChild(501);
    setTimeout(() => {
      child.cancelled = true;
      cb(new Error('Command failed: killed'), '', '');
    }, 5);
    return child;
  };
  const handler = createRequestHandler({ execFile });
  const res = response();
  await handler(request('POST', '/api/exec-python', {
    code: 'pass',
    target: { sessionId: 's1', workbookName: 'Nilai.xlsm', workbookPath: 'C:\\Data\\Nilai.xlsm' }
  }), res);

  const data = JSON.parse(res.body);
  assert.equal(data.cancelled, true, 'harus ditandai cancelled');
  assert.equal(data.ok, false);
  assert.match(data.error, /dihentikan/i);
});

test('server menyediakan endpoint cancel dan registry eksekusi', () => {
  const source = fs.readFileSync(path.join(__dirname, '..', 'server.js'), 'utf8');
  assert.match(source, /req\.url === '\/api\/cancel'/);
  assert.match(source, /registerExecution\(progressKey, child\)/);
  assert.match(source, /unregisterExecution\(progressKey, child\)/);
  // OpenCode juga ikut dihentikan
  assert.match(source, /activeOpencodeCancel\.cancel/);
});

test('taskpane: tombol Esc dan tombol Hentikan memicu stopProcessing', () => {
  const js = fs.readFileSync(path.join(__dirname, '..', 'taskpane', 'taskpane.js'), 'utf8');
  const html = fs.readFileSync(path.join(__dirname, '..', 'taskpane', 'taskpane.html'), 'utf8');

  assert.match(html, /id="stopBtn"/);
  assert.match(js, /function stopProcessing/);
  // keyCode 27 = Esc
  assert.match(js, /e\.keyCode === 27/);
  assert.match(js, /stopBtn\.onclick/);
  // memanggil endpoint cancel di server
  assert.match(js, /\/api\/cancel/);
});

test('taskpane: XHR dilacak agar bisa di-abort saat dibatalkan', () => {
  const js = fs.readFileSync(path.join(__dirname, '..', 'taskpane', 'taskpane.js'), 'utf8');
  assert.match(js, /function trackXhr/);
  assert.match(js, /function untrackXhr/);
  assert.match(js, /function abortAllXhrs/);
  // streaming turn ikut dilacak
  assert.match(js, /trackXhr\(xhr\);[\s\S]{0,200}xhr\.open\('POST', SERVER \+ '\/api\/proxy'/);
});

test('taskpane: rantai eksekusi berhenti saat cancelRequested aktif', () => {
  const js = fs.readFileSync(path.join(__dirname, '..', 'taskpane', 'taskpane.js'), 'utf8');
  const guards = js.match(/if \(cancelRequested\) \{ isProcessing = false; return; \}/g) || [];
  // minimal terpasang di 6 fungsi rantai agent
  assert.ok(guards.length >= 6, 'guard cancelRequested kurang, ditemukan ' + guards.length);
  // flag direset saat giliran baru
  assert.match(js, /cancelRequested = false;/);
  // watchdog tidak memunculkan pesan palsu setelah dibatalkan
  assert.match(js, /if \(cancelRequested\) return;[\s\S]{0,300}Proses memakan waktu sangat lama/);
});
