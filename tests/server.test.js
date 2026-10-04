const test = require('node:test');
const assert = require('node:assert/strict');
const { EventEmitter } = require('node:events');
const http = require('node:http');
const { Readable } = require('node:stream');
const os = require('node:os');
const fs = require('node:fs');
const path = require('node:path');
const {
  cleanHistory,
  createRequestHandler,
  enqueueCom,
  isAllowedOrigin,
  SERVER_HOST,
  validateExecutionEnvelope,
  workbookHistoryKey,
  publicSettings,
  resolveProxyApiKey,
  saveProvider,
  deleteProvider,
  activateProvider,
  sanitizeProviderName,
  MAX_REQUEST_BODY_BYTES,
  MAX_PENDING_COM,
  PROXY_TIMEOUT_MS,
  EXEC_TIMEOUT_MS,
  readBody
} = require('../server');

const target = {
  sessionId: 's1',
  workbookName: 'Nilai.xlsm',
  workbookPath: 'C:\\Data\\Nilai.xlsm'
};

function request(method, url, body) {
  const req = Readable.from(body === undefined ? [] : [Buffer.from(JSON.stringify(body))]);
  req.method = method;
  req.url = url;
  req.headers = {};
  return req;
}

function response() {
  return {
    statusCode: undefined,
    headers: undefined,
    body: '',
    writeHead(statusCode, headers) {
      this.statusCode = statusCode;
      this.headers = headers;
    },
    end(body = '') {
      this.body = body;
    }
  };
}

test('validasi envelope mewajibkan kode dan target lengkap', () => {
  assert.deepEqual(validateExecutionEnvelope({ code: 'wb.Save()', target }), { code: 'wb.Save()', target });
  assert.throws(() => validateExecutionEnvelope({ code: 'pass' }), /target/i);
  assert.throws(() => validateExecutionEnvelope({ code: '', target }), /kosong/i);
});

test('key riwayat stabil terhadap kapital path', () => {
  assert.equal(workbookHistoryKey('C:\\Data\\Nilai.xlsm'), workbookHistoryKey('c:\\data\\NILAI.xlsm'));
  assert.match(workbookHistoryKey(target.workbookPath), /^[a-f0-9]{64}$/);
});

test('key riwayat menerima URL OneDrive dan stabil terhadap host kapital', () => {
  assert.equal(
    workbookHistoryKey('https://D.DOCS.LIVE.NET/A652/Documents/Daftar%20hadir.xlsx'),
    workbookHistoryKey('https://d.docs.live.net/a652/Documents/Daftar hadir.xlsx')
  );
});

test('origin hanya mengizinkan add-in lokal dan request tanpa Origin', () => {
  assert.equal(isAllowedOrigin(undefined), true);
  assert.equal(isAllowedOrigin('https://localhost:3000'), true);
  assert.equal(isAllowedOrigin('https://evil.example'), false);
});

test('antrean COM menjalankan pekerjaan berurutan', async () => {
  const events = [];
  const first = enqueueCom(async () => { events.push('mulai-1'); await new Promise(r => setTimeout(r, 20)); events.push('selesai-1'); });
  const second = enqueueCom(async () => { events.push('mulai-2'); events.push('selesai-2'); });
  await Promise.all([first, second]);
  assert.deepEqual(events, ['mulai-1', 'selesai-1', 'mulai-2', 'selesai-2']);
});

test('antrean COM menolak pekerjaan di atas batas dengan 429 dan pulih setelah selesai', async () => {
  let release;
  const gate = new Promise(resolve => { release = resolve; });
  const queued = [];
  for (let index = 0; index < MAX_PENDING_COM; index += 1) queued.push(enqueueCom(() => gate));
  await assert.rejects(enqueueCom(() => Promise.resolve()), error => error.statusCode === 429);
  release();
  await Promise.all(queued);
  await enqueueCom(() => Promise.resolve());
});

test('server HTTPS dibatasi ke loopback IPv4', () => {
  assert.equal(SERVER_HOST, '127.0.0.1');
});

test('settings tidak mengekspos API key, hanya status keberadaan key', async () => {
  const handler = createRequestHandler({
    readConfig: () => ({
      apiKey: 'secret',
      model: 'm1',
      chatHistory: [{ role: 'user', content: 'legacy' }],
      chatHistoryByWorkbook: { hash: [{ role: 'user', content: 'rahasia' }] }
    })
  });
  const res = response();

  await handler(request('GET', '/api/settings'), res);

  assert.equal(res.statusCode, 200);
  assert.deepEqual(JSON.parse(res.body), { hasApiKey: true, model: 'm1' });
  assert.equal(publicSettings({ apiKey: 'secret' }).apiKey, undefined);
});

test('settings POST mempertahankan API key saat incoming kosong', async () => {
  let saved;
  const config = { apiKey: 'lama', chatHistory: [{ role: 'user', content: 'lama' }] };
  const handler = createRequestHandler({
    readConfig: () => config,
    writeConfig: (partial) => {
      saved = partial;
      Object.assign(config, partial);
      return Object.assign({}, config, { chatHistory: [{ role: 'user', content: 'lama' }] });
    }
  });
  const res = response();

  await handler(request('POST', '/api/settings', {
    apiKey: '',
    model: 'baru',
    chatHistory: [{ role: 'user', content: 'jangan simpan' }]
  }), res);

  assert.equal(res.statusCode, 200);
  assert.deepEqual(saved, { model: 'baru' });
  assert.deepEqual(JSON.parse(res.body), { ok: true, config: { hasApiKey: true, model: 'baru' } });
});

test('publicSettings menyertakan daftar provider tanpa membocorkan API key', () => {
  const out = publicSettings({
    apiKey: 'aktif-secret',
    providers: {
      'Hallo Router': { apiEndpoint: 'https://api.hallorouter.com/v1', model: 'hr/hallo/gemini-3.1-pro', apiKey: 'rahasia-hallo' },
      '9router lokal': { apiEndpoint: 'http://localhost:20128/v1', model: 'qd/ultimate', apiKey: '' }
    }
  });
  // Key mentah TIDAK boleh muncul di mana pun
  assert.equal(JSON.stringify(out).includes('rahasia-hallo'), false);
  assert.equal(out.providers['Hallo Router'].hasKey, true);
  assert.equal(out.providers['Hallo Router'].apiKey, undefined);
  assert.equal(out.providers['9router lokal'].hasKey, false);
  assert.equal(out.providers['Hallo Router'].model, 'hr/hallo/gemini-3.1-pro');
});

test('saveProvider menyimpan entry dan tidak menghapus key saat disimpan ulang tanpa key', () => {
  // Test lewat handler agar memakai jalur nyata (read/write in-memory).
  let config = { providers: {} };
  const read = () => config;
  const write = (p) => { config = Object.assign({}, config, p); return config; };
  return (async () => {
    const handler = createRequestHandler({ readConfig: read, writeConfig: write });
    // 1. Simpan provider dengan key
    let res = response();
    await handler(request('POST', '/api/providers', {
      name: 'Hallo', apiEndpoint: 'https://api.hallorouter.com/v1', model: 'hr/hallo/gemini-3.1-pro', apiKey: 'k-hallo-123'
    }), res);
    assert.equal(res.statusCode, 200);
    assert.equal(config.providers['Hallo'].apiKey, 'k-hallo-123');
    // 2. Simpan ulang TANPA key (mis. hanya ganti model) — key lama harus bertahan
    res = response();
    await handler(request('POST', '/api/providers', { name: 'Hallo', model: 'hr/hallo/gpt-5.6-sol' }), res);
    assert.equal(res.statusCode, 200);
    assert.equal(config.providers['Hallo'].apiKey, 'k-hallo-123');
    assert.equal(config.providers['Hallo'].model, 'hr/hallo/gpt-5.6-sol');
  })();
});

test('activateProvider menyalin endpoint/model/key provider ke setting aktif', async () => {
  let config = {
    apiKey: 'lama', apiEndpoint: 'http://old/v1', model: 'old-model', activeProvider: '',
    providers: { 'Hallo': { apiEndpoint: 'https://api.hallorouter.com/v1', model: 'hr/hallo/gemini-3.1-pro', apiKey: 'k-hallo' } }
  };
  const handler = createRequestHandler({
    readConfig: () => config,
    writeConfig: (p) => { config = Object.assign({}, config, p); return config; }
  });
  const res = response();
  await handler(request('POST', '/api/providers/activate', { name: 'Hallo' }), res);
  assert.equal(res.statusCode, 200);
  assert.equal(config.apiEndpoint, 'https://api.hallorouter.com/v1');
  assert.equal(config.model, 'hr/hallo/gemini-3.1-pro');
  assert.equal(config.apiKey, 'k-hallo');
  assert.equal(config.activeProvider, 'Hallo');
  // Respons ke client tidak membocorkan key
  assert.equal(res.body.includes('k-hallo'), false);
});

test('activateProvider mengembalikan 404 untuk provider tak dikenal', async () => {
  let config = { providers: {} };
  const handler = createRequestHandler({
    readConfig: () => config,
    writeConfig: (p) => { config = Object.assign({}, config, p); return config; }
  });
  const res = response();
  await handler(request('POST', '/api/providers/activate', { name: 'TidakAda' }), res);
  assert.equal(res.statusCode, 404);
});

test('deleteProvider menghapus entry dan mereset activeProvider bila cocok', async () => {
  let config = {
    activeProvider: 'Hallo',
    providers: { 'Hallo': { apiEndpoint: 'e', model: 'm', apiKey: 'k' }, 'Lain': { apiEndpoint: 'e2', model: 'm2', apiKey: 'k2' } }
  };
  const handler = createRequestHandler({
    readConfig: () => config,
    writeConfig: (p) => { config = Object.assign({}, config, p); return config; }
  });
  const res = response();
  await handler(request('DELETE', '/api/providers', { name: 'Hallo' }), res);
  assert.equal(res.statusCode, 200);
  assert.equal(config.providers['Hallo'], undefined);
  assert.equal(config.providers['Lain'].model, 'm2');
  assert.equal(config.activeProvider, '');
});

test('sanitizeProviderName menolak nama kosong atau kepanjangan', () => {
  assert.equal(sanitizeProviderName('  Hallo  '), 'Hallo');
  assert.equal(sanitizeProviderName(''), '');
  assert.equal(sanitizeProviderName('   '), '');
  assert.equal(sanitizeProviderName('x'.repeat(61)), '');
  assert.equal(sanitizeProviderName(123), '');
});

test('settings POST menyimpan modelList (daftar model bisa ganti-ganti)', async () => {
  let config = { apiKey: 'k', model: 'm0', modelList: [] };
  const handler = createRequestHandler({
    readConfig: () => config,
    writeConfig: (p) => { config = Object.assign({}, config, p); return config; }
  });
  const res = response();
  await handler(request('POST', '/api/settings', {
    model: 'stealth/ox-alpha',
    modelList: ['stealth/ox-alpha', 'hr/hallo/gemini-3.1-pro', 'qd/ultimate']
  }), res);
  assert.equal(res.statusCode, 200);
  assert.deepEqual(config.modelList, ['stealth/ox-alpha', 'hr/hallo/gemini-3.1-pro', 'qd/ultimate']);
  // modelList ikut dikembalikan ke client (bukan field rahasia)
  const body = JSON.parse(res.body);
  assert.deepEqual(body.config.modelList, ['stealth/ox-alpha', 'hr/hallo/gemini-3.1-pro', 'qd/ultimate']);
});



test('proxy memilih API key tersimpan saat request tidak membawa key', () => {
  assert.equal(resolveProxyApiKey(undefined, 'stored-secret'), 'stored-secret');
  assert.equal(resolveProxyApiKey('', 'stored-secret'), 'stored-secret');
  assert.equal(resolveProxyApiKey('request-secret', 'stored-secret'), 'request-secret');
});

test('batas body request ditentukan satu MiB', () => {
  assert.equal(MAX_REQUEST_BODY_BYTES, 1024 * 1024);
});

test('timeout proxy memberi waktu cukup untuk tugas Excel kompleks', () => {
  assert.equal(PROXY_TIMEOUT_MS, 180000);
});

test('timeout executor Python tidak lebih pendek dari timeout gateway', () => {
  assert.equal(EXEC_TIMEOUT_MS, 210000);
});

test('reader menolak body melebihi satu MiB dengan status 413', async () => {
  const req = Readable.from([Buffer.alloc(MAX_REQUEST_BODY_BYTES + 1)]);
  await assert.rejects(readBody(req), error => error.statusCode === 413);
});

test('handler mengembalikan JSON 413 dan server tetap dapat dipakai setelah POST terlalu besar', async (t) => {
  const config = { apiKey: 'secret-test', apiEndpoint: '', model: '', autoRun: true, chatHistoryByWorkbook: {} };
  const server = http.createServer(createRequestHandler({
    readConfig: () => config,
    writeConfig: partial => Object.assign(config, partial)
  }));
  await new Promise((resolve, reject) => {
    server.once('error', reject);
    server.listen(0, '127.0.0.1', resolve);
  });
  t.after(() => new Promise(resolve => server.close(resolve)));
  const port = server.address().port;

  function send(method, url, body) {
    return new Promise((resolve, reject) => {
      const req = http.request({ hostname: '127.0.0.1', port, method, path: url, headers: { 'Content-Type': 'application/json' } }, res => {
        const chunks = [];
        res.on('data', chunk => chunks.push(chunk));
        res.on('end', () => resolve({ status: res.statusCode, body: Buffer.concat(chunks).toString() }));
      });
      req.on('error', reject);
      req.end(body);
    });
  }

  const oversized = await send('POST', '/api/settings', Buffer.alloc(MAX_REQUEST_BODY_BYTES + 1, 0x20));
  assert.equal(oversized.status, 413);
  assert.deepEqual(JSON.parse(oversized.body), { ok: false, error: 'Body request terlalu besar.' });

  const next = await send('POST', '/api/settings', JSON.stringify({ model: 'masih-hidup' }));
  assert.equal(next.status, 200);
  assert.equal(JSON.parse(next.body).ok, true);
});

test('static server menolak traversal terenkripsi dan tetap melayani taskpane', async (t) => {
  const server = http.createServer(createRequestHandler());
  await new Promise((resolve, reject) => {
    server.once('error', reject);
    server.listen(0, '127.0.0.1', resolve);
  });
  t.after(() => new Promise(resolve => server.close(resolve)));
  const port = server.address().port;

  function get(url) {
    return new Promise((resolve, reject) => {
      http.get({ hostname: '127.0.0.1', port, path: url }, res => {
        const chunks = [];
        res.on('data', chunk => chunks.push(chunk));
        res.on('end', () => resolve({
          status: res.statusCode,
          headers: res.headers,
          body: Buffer.concat(chunks).toString()
        }));
      }).on('error', reject);
    });
  }

  for (const url of ['/%2e%2e/CLAUDE.md', '/..%2fCLAUDE.md', '/..%5cCLAUDE.md']) {
    const result = await get(url);
    assert.notEqual(result.status, 200, url);
    assert.doesNotMatch(result.body, /What this is|graphify/, url);
  }

  const malformed = await get('/%E0%A4%A');
  assert.ok(malformed.status === 400 || malformed.status === 404);

  const taskpane = await get('/taskpane.html');
  assert.equal(taskpane.status, 200);
  assert.equal(taskpane.headers['access-control-allow-origin'], undefined);
  assert.match(taskpane.body, /Excel Agent/);
});

async function withHistoryServer(t) {
  const config = { chatHistoryByWorkbook: {} };
  const server = http.createServer(createRequestHandler({
    readConfig: () => config,
    writeConfig: (partial) => Object.assign(config, partial)
  }));
  await new Promise((resolve, reject) => {
    server.once('error', reject);
    server.listen(0, '127.0.0.1', resolve);
  });
  t.after(() => new Promise(resolve => server.close(resolve)));
  const port = server.address().port;
  return function send(method, body) {
    return new Promise((resolve, reject) => {
      const payload = Buffer.isBuffer(body) ? body : Buffer.from(body);
      const req = http.request({
        hostname: '127.0.0.1',
        port,
        method,
        path: '/api/history',
        headers: { 'Content-Type': 'application/json', 'Content-Length': payload.length }
      }, res => {
        const chunks = [];
        res.on('data', chunk => chunks.push(chunk));
        res.on('end', () => resolve({ status: res.statusCode, body: Buffer.concat(chunks).toString() }));
      });
      req.on('error', reject);
      req.end(payload);
    });
  };
}

test('history POST mempertahankan JSON 413 lalu menerima request kecil berikutnya', async (t) => {
  const send = await withHistoryServer(t);
  const oversized = await send('POST', Buffer.alloc(MAX_REQUEST_BODY_BYTES + 1, 0x20));
  assert.equal(oversized.status, 413);
  assert.deepEqual(JSON.parse(oversized.body), { ok: false, error: 'Body request terlalu besar.' });

  const next = await send('POST', JSON.stringify({
    workbookPath: target.workbookPath,
    chatHistory: [{ role: 'user', content: 'kecil' }]
  }));
  assert.equal(next.status, 200);
  assert.equal(JSON.parse(next.body).ok, true);
});

test('history DELETE mempertahankan JSON 413 lalu menerima request kecil berikutnya', async (t) => {
  const send = await withHistoryServer(t);
  const oversized = await send('DELETE', Buffer.alloc(MAX_REQUEST_BODY_BYTES + 1, 0x20));
  assert.equal(oversized.status, 413);
  assert.deepEqual(JSON.parse(oversized.body), { ok: false, error: 'Body request terlalu besar.' });

  const next = await send('DELETE', JSON.stringify({ workbookPath: target.workbookPath }));
  assert.equal(next.status, 200);
  assert.equal(JSON.parse(next.body).ok, true);
});

test('history membersihkan role dan menyimpan maksimum 50 pesan', async () => {
  let saved;
  const original = Array.from({ length: 55 }, (_, index) => ({ role: index % 2 ? 'assistant' : 'user', content: String(index) }));
  original.splice(10, 0, { role: 'system', content: 'tolak' });
  const handler = createRequestHandler({
    readConfig: () => ({ chatHistoryByWorkbook: {} }),
    writeConfig: (partial) => { saved = partial; return partial; }
  });
  const res = response();

  await handler(request('POST', '/api/history', { workbookPath: target.workbookPath, chatHistory: original }), res);

  const history = JSON.parse(res.body).chatHistory;
  assert.equal(res.statusCode, 200);
  assert.equal(history.length, 50);
  assert.equal(history[0].content, '5');
  assert.deepEqual(saved.chatHistoryByWorkbook[workbookHistoryKey(target.workbookPath)], history);
  assert.equal(cleanHistory(original).length, 50);
});

test('executor mengirim envelope JSON utuh ke stdin', async () => {
  let stdinPayload;
  const execFile = (pythonPath, args, options, callback) => {
    const stdin = new EventEmitter();
    stdin.write = (value) => { stdinPayload = value; };
    stdin.end = () => callback(null, JSON.stringify({ ok: true }), '');
    return { stdin };
  };
  const handler = createRequestHandler({ execFile });
  const res = response();
  const envelope = { code: 'wb.Save()', target };

  await handler(request('POST', '/api/exec-python', envelope), res);

  assert.equal(res.statusCode, 200);
  const parsed = JSON.parse(stdinPayload);
  assert.equal(parsed.code, envelope.code);
  assert.deepEqual(parsed.target, envelope.target);
  assert.match(parsed.progressFile, /progress_[a-f0-9]{64}\.txt$/);
  assert.deepEqual(JSON.parse(res.body), { ok: true });
});

test('executor meneruskan target error dengan status 409', async () => {
  const execFile = (pythonPath, args, options, callback) => {
    const stdin = new EventEmitter();
    stdin.write = () => {};
    stdin.end = () => callback(null, JSON.stringify({ ok: false, targetError: true, error: 'target berubah' }), '');
    return { stdin };
  };
  const handler = createRequestHandler({ execFile });
  const res = response();

  await handler(request('POST', '/api/exec-python', { code: 'pass', target }), res);

  assert.equal(res.statusCode, 409);
  assert.deepEqual(JSON.parse(res.body), { ok: false, targetError: true, error: 'target berubah' });
});

test('executor mempertahankan error kode biasa sebagai non-target', async () => {
  const execFile = (pythonPath, args, options, callback) => {
    const stdin = new EventEmitter();
    stdin.write = () => {};
    stdin.end = () => callback(null, JSON.stringify({ ok: false, error: 'kode gagal' }), '');
    return { stdin };
  };
  const handler = createRequestHandler({ execFile });
  const res = response();

  await handler(request('POST', '/api/exec-python', { code: 'raise RuntimeError()', target }), res);

  assert.equal(res.statusCode, 200);
  assert.equal(JSON.parse(res.body).targetError, undefined);
});

test('validation target executor gagal sebagai target error HTTP 400', async () => {
  const handler = createRequestHandler();
  for (const envelope of [
    { code: 'pass' },
    { code: 'pass', target: { sessionId: 's1', workbookName: 'Nilai.xlsm' } },
    { code: 'pass', target: Object.assign({}, target, { workbookPath: 'relative.xlsm' }) }
  ]) {
    const res = response();
    await handler(request('POST', '/api/exec-python', envelope), res);
    assert.equal(res.statusCode, 400);
    assert.equal(JSON.parse(res.body).targetError, true);
  }
});

test('validation menolak nama workbook yang tidak sama dengan basename sebelum spawn', async () => {
  let spawned = false;
  const handler = createRequestHandler({ execFile: () => { spawned = true; } });
  const res = response();
  await handler(request('POST', '/api/exec-python', {
    code: 'pass',
    target: Object.assign({}, target, { workbookName: 'Lain.xlsm' })
  }), res);
  assert.equal(res.statusCode, 400);
  assert.equal(JSON.parse(res.body).targetError, true);
  assert.equal(spawned, false);
});

test('taskpane memblokir binding pada targetError di semua jalur eksekusi', () => {
  const source = require('node:fs').readFileSync(require('node:path').join(__dirname, '..', 'taskpane', 'taskpane.js'), 'utf8');
  assert.match(source, /if \(result\.targetError\)[\s\S]{0,300}setBindingBlocked/);
  assert.ok((source.match(/setBindingBlocked\('Target workbook tidak valid\. Buka ulang task pane untuk melanjutkan\.'\)/g) || []).length >= 3);
});

test('taskpane tidak menyimpan atau mengirim API key tersimpan dari state client', () => {
  const source = require('node:fs').readFileSync(require('node:path').join(__dirname, '..', 'taskpane', 'taskpane.js'), 'utf8');
  assert.doesNotMatch(source, /settings\.apiKey/);
  assert.match(source, /delete parsed\.apiKey/);
  assert.doesNotMatch(source, /localStorage\.setItem\([^\n]+JSON\.stringify\(settings\)/);
});

test('validation kode kosong executor tetap non-target 400', async () => {
  const handler = createRequestHandler();
  const res = response();

  await handler(request('POST', '/api/exec-python', { code: '', target }), res);

  assert.equal(res.statusCode, 400);
  assert.deepEqual(JSON.parse(res.body), { ok: false, error: 'Kode Python kosong.' });
});

test('opencode: validasi model/session menolak karakter shell', async () => {
  const { OPENCODE_MODEL_RE, OPENCODE_SESSION_RE, runOpencode } = require('../server');
  assert.equal(OPENCODE_MODEL_RE.test('9router/arap/claude-opus-5'), true);
  assert.equal(OPENCODE_MODEL_RE.test('model & calc'), false);
  assert.equal(OPENCODE_MODEL_RE.test('a;whoami'), false);
  assert.equal(OPENCODE_SESSION_RE.test('ses_abc123-X'), true);
  assert.equal(OPENCODE_SESSION_RE.test('ses; rm'), false);
  await assert.rejects(runOpencode('halo', 'model & jahat', '', '', ''), /tidak valid/);
  await assert.rejects(runOpencode('halo', '', 'ses; rm', '', ''), /tidak valid/);
});

test('opencode endpoint menolak prompt kosong dengan 400', async () => {
  const handler = createRequestHandler();
  const res = response();
  await handler(request('POST', '/api/opencode', { prompt: '   ' }), res);
  assert.equal(res.statusCode, 400);
  assert.match(JSON.parse(res.body).error, /kosong/i);
});

test('timeout opencode lebih panjang dari timeout executor', () => {
  const { OPENCODE_TIMEOUT_MS } = require('../server');
  assert.ok(OPENCODE_TIMEOUT_MS > EXEC_TIMEOUT_MS);
});

test('lampiran OpenCode membuat file gambar temporary dan cleanup menghapusnya', () => {
  const { prepareOpencodeImages } = require('../server');
  const root = fs.mkdtempSync(path.join(os.tmpdir(), 'excelagent-test-'));
  const prepared = prepareOpencodeImages([{
    name: 'contoh.png',
    type: 'image/png',
    data: 'data:image/png;base64,iVBORw0KGgo='
  }], root);

  assert.equal(prepared.paths.length, 1);
  assert.equal(fs.existsSync(prepared.paths[0]), true);
  assert.equal(fs.readFileSync(prepared.paths[0]).toString('base64'), 'iVBORw0KGgo=');

  const requestDir = path.dirname(prepared.paths[0]);
  prepared.cleanup();
  assert.equal(fs.existsSync(requestDir), false);
  fs.rmSync(root, { recursive: true, force: true });
});

test('lampiran OpenCode menolak format, ukuran, dan jumlah tidak valid', () => {
  const { prepareOpencodeImages } = require('../server');
  const png = { name: 'x.png', type: 'image/png', data: 'data:image/png;base64,iVBORw0KGgo=' };

  assert.throws(() => prepareOpencodeImages(Array(6).fill(png)), /maksimal/i);
  assert.throws(() => prepareOpencodeImages([{ name: 'x.svg', type: 'image/svg+xml', data: 'data:image/svg+xml;base64,PHN2Zz4=' }]), /format/i);
  assert.throws(() => prepareOpencodeImages([{ name: 'x.png', type: 'image/png', data: 'bukan-data-url' }]), /base64/i);
  const tooLarge = Buffer.alloc(5 * 1024 * 1024 + 1).toString('base64');
  assert.throws(() => prepareOpencodeImages([{ name: 'x.png', type: 'image/png', data: 'data:image/png;base64,' + tooLarge }]), /5 MB/i);
});

test('endpoint OpenCode memakai batas body gambar, meneruskan path, dan cleanup', () => {
  const source = fs.readFileSync(path.join(__dirname, '..', 'server.js'), 'utf8');
  assert.match(source, /readBody\(req, MAX_OPENCODE_BODY_BYTES\)/);
  assert.match(source, /runOpencode\(prompt, model, sessionId, dir, progressFile, cancelRef, preparedImages\.paths\)/);
  assert.match(source, /finally \{\s*preparedImages\.cleanup\(\);/);
});

test('health endpoint mengembalikan status server, router, dan python', async () => {
  const handler = createRequestHandler();
  const res = response();
  await handler(request('GET', '/api/health'), res);
  assert.equal(res.statusCode, 200);
  const data = JSON.parse(res.body);
  assert.equal(data.ok, true);
  assert.equal(data.server.status, 'online');
  assert.ok(typeof data.server.port === 'number');
  assert.ok(data.router !== undefined);
  assert.ok(data.python !== undefined);
});

test('checkRouterHealth menangani URL kosong atau tidak valid', async () => {
  const { checkRouterHealth } = require('../server');
  const resEmpty = await checkRouterHealth('');
  assert.equal(resEmpty.ok, false);
  assert.equal(resEmpty.status, 'unconfigured');

  const resInvalid = await checkRouterHealth('bukan-url');
  assert.equal(resInvalid.ok, false);
  assert.equal(resInvalid.status, 'invalid_url');
});

// ---- Integrasi Hermes ----

function writeHermesHome(configYaml, envText) {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'hermes-test-'));
  fs.writeFileSync(path.join(dir, 'config.yaml'), configYaml, 'utf8');
  if (envText !== undefined) fs.writeFileSync(path.join(dir, '.env'), envText, 'utf8');
  return dir;
}

test('hermes: endpoint chat completions dibentuk dari base_url', () => {
  const h = require('../hermes_providers');
  assert.equal(h.toChatCompletionsEndpoint('https://x.test/v1'), 'https://x.test/v1/chat/completions');
  assert.equal(h.toChatCompletionsEndpoint('https://x.test/v1/'), 'https://x.test/v1/chat/completions');
  assert.equal(h.toChatCompletionsEndpoint('https://x.test/v1/chat/completions'), 'https://x.test/v1/chat/completions');
  assert.equal(h.toChatCompletionsEndpoint(''), '');
});

test('hermes: custom_providers dibaca lengkap dengan daftar model dan API key dari .env', () => {
  const h = require('../hermes_providers');
  const dir = writeHermesHome([
    'model:',
    '  default: FreeCombo',
    '  provider: custom',
    '  base_url: https://router.test/v1',
    '  api_key: ${HERMES_ROUTER_KEY}',
    'custom_providers:',
    '  - name: RouterTest',
    '    base_url: https://router.test/v1',
    '    key_env: HERMES_ROUTER_KEY',
    '    model: FreeCombo',
    '    models:',
    '      FreeCombo: {}',
    '      ag/gemini-pro: {}'
  ].join('\n'), 'HERMES_ROUTER_KEY=rahasia-123\n');

  const result = h.readHermesProviders(dir);
  assert.equal(result.ok, true);
  assert.equal(result.providers.length, 1);
  const p = result.providers[0];
  assert.equal(p.name, 'RouterTest');
  assert.equal(p.apiEndpoint, 'https://router.test/v1/chat/completions');
  assert.equal(p.apiKey, 'rahasia-123');
  assert.deepEqual(p.models, ['FreeCombo', 'ag/gemini-pro']);
  fs.rmSync(dir, { recursive: true, force: true });
});

test('hermes: fallback_providers ikut terbaca dan model aktif digabung ke provider ber-base_url sama', () => {
  const h = require('../hermes_providers');
  const dir = writeHermesHome([
    'model:',
    '  default: model-utama',
    '  provider: custom',
    '  base_url: https://router.test/v1',
    '  api_key: ${K1}',
    'custom_providers:',
    '  - name: RouterTest',
    '    base_url: https://router.test/v1',
    '    key_env: K1',
    '    model: lain',
    'fallback_providers:',
    '  - name: Lokal',
    '    base_url: http://localhost:20128/v1',
    '    api_key: sk-lokal',
    '    model: qd/ultimate'
  ].join('\n'), 'K1=key-satu\n');

  const result = h.readHermesProviders(dir);
  assert.equal(result.ok, true);
  const names = result.providers.map(p => p.name);
  assert.deepEqual(names, ['RouterTest', 'Lokal']);
  // model aktif Hermes digabung ke provider dengan base_url sama, bukan provider baru
  assert.equal(result.providers[0].model, 'model-utama');
  assert.ok(result.providers[0].models.indexOf('model-utama') !== -1);
  assert.equal(result.providers[1].apiKey, 'sk-lokal');
  fs.rmSync(dir, { recursive: true, force: true });
});

test('hermes: publicHermesProviders tidak pernah membocorkan API key', () => {
  const h = require('../hermes_providers');
  const dir = writeHermesHome([
    'custom_providers:',
    '  - name: RouterTest',
    '    base_url: https://router.test/v1',
    '    key_env: K1',
    '    model: m1'
  ].join('\n'), 'K1=super-rahasia\n');

  const publicView = h.publicHermesProviders(h.readHermesProviders(dir));
  assert.equal(JSON.stringify(publicView).includes('super-rahasia'), false);
  assert.equal(publicView.providers[0].hasKey, true);
  fs.rmSync(dir, { recursive: true, force: true });
});

test('hermes: konfigurasi tidak ditemukan dilaporkan sebagai gagal, bukan crash', () => {
  const h = require('../hermes_providers');
  const result = h.readHermesProviders(path.join(os.tmpdir(), 'hermes-tidak-ada-' + Date.now()));
  assert.equal(result.ok, false);
  assert.ok(result.error);
  assert.deepEqual(result.providers, []);
});

test('hermes: endpoint /api/hermes/providers melayani katalog tanpa key', async () => {
  const handler = createRequestHandler();
  const res = response();
  await handler(request('GET', '/api/hermes/providers'), res);
  assert.ok(res.statusCode === 200 || res.statusCode === 404);
  const data = JSON.parse(res.body);
  assert.ok(Array.isArray(data.providers));
  for (const p of data.providers) {
    assert.equal(p.apiKey, undefined);
    assert.equal(typeof p.hasKey, 'boolean');
  }
});

test('hermes: import menolak nama provider yang tidak ada', async () => {
  const handler = createRequestHandler();
  const res = response();
  await handler(request('POST', '/api/hermes/import', { name: 'Provider-Tidak-Ada-xyz' }), res);
  assert.ok(res.statusCode >= 400);
  assert.equal(JSON.parse(res.body).ok, false);
});

test('taskpane menyediakan kontrol impor provider Hermes', () => {
  const source = fs.readFileSync(path.join(__dirname, '..', 'taskpane', 'taskpane.js'), 'utf8');
  const html = fs.readFileSync(path.join(__dirname, '..', 'taskpane', 'taskpane.html'), 'utf8');
  assert.match(html, /id="hermesProviderSelect"/);
  assert.match(html, /id="importHermesBtn"/);
  assert.match(source, /\/api\/hermes\/providers/);
  assert.match(source, /\/api\/hermes\/import/);
});

