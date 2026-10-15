const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

function loadTaskpane() {
  const elements = {};
  function element(id) {
    if (!elements[id]) {
      elements[id] = {
        id,
        className: '',
        textContent: '',
        value: '',
        disabled: false,
        style: {},
        innerHTML: '',
        appendChild() {},
        querySelector() { return null; }
      };
    }
    return elements[id];
  }
  const requests = [];
  function Xhr() {
    this.readyState = 0;
    this.status = 0;
    requests.push(this);
  }
  Xhr.prototype.open = function(method, url) { this.method = method; this.url = url; };
  Xhr.prototype.setRequestHeader = function() {};
  Xhr.prototype.send = function(body) { this.body = body; };
  Xhr.prototype.respond = function(status, body) {
    this.status = status;
    this.responseText = body || '{}';
    this.readyState = 4;
    if (this.onreadystatechange) this.onreadystatechange();
  };

  const context = {
    console,
    JSON,
    Math,
    encodeURIComponent,
    decodeURIComponent,
    XMLHttpRequest: Xhr,
    localStorage: { getItem() { return null; }, setItem() {} },
    document: {
      readyState: 'loading',
      addEventListener() {},
      getElementById: element,
      querySelectorAll() { return []; },
      createElement() { return element('created'); },
      body: { appendChild() {}, removeChild() {} }
    },
    Office: {
      AsyncResultStatus: { Succeeded: 'succeeded' },
      context: { document: { getFilePropertiesAsync() {} } },
      onReady() {}
    },
    window: {},
    setTimeout(fn) { fn(); },
    clearTimeout() {},
    setInterval(fn, ms) { return 1; },
    clearInterval() {}
  };
  vm.createContext(context);
  vm.runInContext(fs.readFileSync(path.join(__dirname, '..', 'taskpane', 'taskpane.js'), 'utf8'), context);
  return { context, elements, requests };
}

test('targetError memblokir binding permanen', () => {
  const { context, elements } = loadTaskpane();
  const messages = [];
  context.workbookBinding = { sessionId: 's1', workbookName: 'Nilai.xlsm', workbookPath: 'C:\\Data\\Nilai.xlsm' };
  context.addMessage = (role, content) => messages.push({ role, content });
  context.getSheetSnapshot = callback => callback({ empty: false });
  context.runPythonCode = (code, callback) => callback({ ok: false, targetError: true, error: 'ROT gagal' });
  context.executeBoundCodeWithRetry('pass', 0, '', '');
  assert.equal(context.bindingBlocked, true);
  assert.equal(context.workbookBinding, null);
  assert.equal(elements.userInput.disabled, true);
  assert.match(messages[0].content, /Target workbook tidak valid/);
});

test('normalizeFileUrl menerima path Windows langsung dari Excel desktop', () => {
  const { context } = loadTaskpane();
  assert.equal(
    context.normalizeFileUrl('C:\\Users\\acer\\Documents\\Daftar hadir.xlsx'),
    'C:\\Users\\acer\\Documents\\Daftar hadir.xlsx'
  );
});

test('normalizeFileUrl tetap menolak URL cloud non-OneDrive', () => {
  const { context } = loadTaskpane();
  assert.equal(context.normalizeFileUrl('https://storage.example.com/Daftar%20hadir.xlsx'), null);
});

test('normalizeFileUrl menerima URL workbook OneDrive', () => {
  const { context } = loadTaskpane();
  assert.equal(
    context.normalizeFileUrl('https://d.docs.live.net/a652565c59200a84/Documents/Daftar%20hadir.xlsx'),
    'https://d.docs.live.net/a652565c59200a84/Documents/Daftar hadir.xlsx'
  );
});

test('normalizeFileUrl menerima URL workbook SharePoint OneDrive bisnis', () => {
  const { context } = loadTaskpane();
  assert.equal(
    context.normalizeFileUrl('https://sekolah.sharepoint.com/sites/Guru/Daftar%20hadir.xlsx'),
    'https://sekolah.sharepoint.com/sites/Guru/Daftar hadir.xlsx'
  );
});

test('binding OneDrive memakai nama workbook yang sudah didekode', () => {
  const { context } = loadTaskpane();
  let bound;
  context.Office.context.document.getFilePropertiesAsync = callback => callback({
    status: context.Office.AsyncResultStatus.Succeeded,
    value: { url: 'https://d.docs.live.net/a652/Documents/Daftar%20hadir.xlsx' }
  });
  context.bindTaskPane(ok => { bound = ok; });
  assert.equal(bound, true);
  assert.equal(context.workbookBinding.workbookName, 'Daftar hadir.xlsx');
  assert.equal(context.workbookBinding.workbookPath, 'https://d.docs.live.net/a652/Documents/Daftar hadir.xlsx');
});

test('taskpane menampilkan versi di pojok kanan bawah', () => {
  const html = fs.readFileSync(path.join(__dirname, '..', 'taskpane', 'taskpane.html'), 'utf8');
  const css = fs.readFileSync(path.join(__dirname, '..', 'taskpane', 'taskpane.css'), 'utf8');
  assert.match(html, /id="taskpaneVersion"[^>]*>v1\.7\.0</);
  assert.match(css, /\.taskpane-version[\s\S]*text-align:\s*right/);
});

test('timeout turn taskpane lebih panjang dari timeout gateway', () => {
  const { context } = loadTaskpane();
  assert.equal(context.AGENT_TURN_TIMEOUT_MS, 210000);
});

test('xhrPost hanya melaporkan satu hasil setelah timeout', () => {
  const { context, requests } = loadTaskpane();
  const results = [];

  context.xhrPost('/api/opencode', { prompt: 'uji' }, (status, data) => {
    results.push({ status, data });
  }, 440000);

  const xhr = requests[0];
  xhr.ontimeout();
  xhr.respond(0, '');
  xhr.onerror();

  assert.equal(results.length, 1);
  assert.equal(results[0].status, 0);
  assert.match(results[0].data.error.message, /timeout/i);
});

test('composer mengizinkan kirim gambar tanpa teks dan memasukkan images ke payload OpenCode', () => {
  const { context, elements } = loadTaskpane();
  let payload;
  context.workbookBinding = { sessionId: 's1', workbookName: 'Nilai.xlsm', workbookPath: 'C:\\Data\\Nilai.xlsm' };
  context.settings.hasApiKey = true;
  context.settings.useOpencode = true;
  context.pendingImages = [{ name: 'referensi.png', type: 'image/png', size: 12, data: 'data:image/png;base64,iVBORw0KGgo=' }];
  context.assertBindingCurrent = callback => callback(null);
  context.saveHistory = () => {};
  context.getWorkbookSchema = callback => callback(null, { sheets: [] });
  context.buildContextMessage = () => 'KONTEKS';
  context.getMcpWorkbookInfo = callback => callback(null);
  context.refreshSettingsFromServer = done => done();
  context.sendOpencodeRequest = (body) => { payload = body; };
  context.document.getElementById('userInput').value = '';

  context.sendMessage();

  assert.equal(payload.prompt.includes('Analisis gambar ini sebagai referensi workbook.'), true);
  assert.equal(payload.images.length, 1);
  assert.equal(payload.images[0].name, 'referensi.png');
  assert.equal(context.pendingImages.length, 0);
});

test('formatApiError membaca error string dan tidak pernah menghasilkan unknown', () => {
  const { context } = loadTaskpane();
  assert.equal(context.formatApiError(500, { error: 'Executor gagal' }), 'Executor gagal');
  assert.equal(context.formatApiError(502, {}), 'Server error (HTTP 502). Coba lagi.');
  assert.equal(context.formatApiError(429, {}), 'Rate limit tercapai. Tunggu sebentar lalu coba lagi.');
  assert.equal(context.formatApiError(0, null), 'Koneksi ke server lokal gagal.');
});

test('multi-step final identity failure tidak melaporkan sukses', () => {
  const { context } = loadTaskpane();
  const messages = [];
  context.workbookBinding = { sessionId: 's1', workbookName: 'Nilai.xlsm', workbookPath: 'C:\\Data\\Nilai.xlsm' };
  context.addMessage = (role, content) => messages.push(content);
  context.Office.context.document.getFilePropertiesAsync = callback => callback({
    status: context.Office.AsyncResultStatus.Succeeded,
    value: { url: 'file:///C:/Data/Lain.xlsm' }
  });
  context.executeMultiStep(['pass'], 2, '', '');
  assert.equal(context.bindingBlocked, true);
  assert.equal(messages.includes('Semua langkah selesai.'), false);
});

test('multi-step targetError menghentikan langkah dan memblokir binding permanen', () => {
  const { context, elements } = loadTaskpane();
  const messages = [];
  const executed = [];
  context.workbookBinding = { sessionId: 's1', workbookName: 'Nilai.xlsm', workbookPath: 'C:\\Data\\Nilai.xlsm' };
  context.addMessage = (role, content) => messages.push(content);
  context.Office.context.document.getFilePropertiesAsync = callback => callback({
    status: context.Office.AsyncResultStatus.Succeeded,
    value: { url: 'file:///C:/Data/Nilai.xlsm' }
  });
  context.getSheetSnapshot = callback => callback({ empty: false });
  context.runPythonCode = (code, callback) => {
    executed.push(code);
    callback({ ok: false, targetError: true, error: 'target berubah' });
  };

  context.executeMultiStep(['langkah-1', 'langkah-2'], 0, '', '');

  // Multi-step sekarang menggabungkan semua code blocks ke satu eksekusi
  assert.equal(executed.length, 1);
  assert.ok(executed[0].includes('langkah-1'));
  assert.ok(executed[0].includes('langkah-2'));
  assert.equal(context.bindingBlocked, true);
  assert.equal(context.workbookBinding, null);
  assert.equal(elements.userInput.disabled, true);
  assert.equal(elements.sendBtn.disabled, true);
  assert.equal(context.isProcessing, false);
});

test('executeJS manual targetError memblokir binding permanen', () => {
  const { context, elements } = loadTaskpane();
  const messages = [];
  const codeBlock = { querySelector: () => ({ textContent: 'pass' }) };
  const btn = {
    disabled: false,
    textContent: 'Jalankan',
    parentNode: { parentNode: { parentNode: codeBlock } }
  };
  context.workbookBinding = { sessionId: 's1', workbookName: 'Nilai.xlsm', workbookPath: 'C:\\Data\\Nilai.xlsm' };
  context.addMessage = (role, content) => messages.push(content);
  context.Office.context.document.getFilePropertiesAsync = callback => callback({
    status: context.Office.AsyncResultStatus.Succeeded,
    value: { url: 'file:///C:/Data/Nilai.xlsm' }
  });
  context.runPythonCode = (code, callback) => callback({ ok: false, targetError: true, error: 'target berubah' });

  context.executeJS(btn);

  assert.equal(context.bindingBlocked, true);
  assert.equal(context.workbookBinding, null);
  assert.equal(elements.userInput.disabled, true);
  assert.equal(elements.sendBtn.disabled, true);
  assert.equal(btn.disabled, true);
  assert.equal(btn.textContent, 'Gagal');
  assert.match(messages[0], /target berubah/);
});

test('executeJS sukses executor tetapi identitas akhir gagal tidak melaporkan sukses', () => {
  const { context, elements } = loadTaskpane();
  const messages = [];
  let identityChecks = 0;
  const codeBlock = { querySelector: () => ({ textContent: 'pass' }) };
  const btn = {
    disabled: false,
    textContent: 'Jalankan',
    parentNode: { parentNode: { parentNode: codeBlock } }
  };
  context.workbookBinding = { sessionId: 's1', workbookName: 'Nilai.xlsm', workbookPath: 'C:\\Data\\Nilai.xlsm' };
  context.addMessage = (role, content) => messages.push(content);
  context.Office.context.document.getFilePropertiesAsync = callback => {
    identityChecks += 1;
    callback({
      status: context.Office.AsyncResultStatus.Succeeded,
      value: { url: identityChecks === 1 ? 'file:///C:/Data/Nilai.xlsm' : 'file:///C:/Data/Lain.xlsm' }
    });
  };
  context.runPythonCode = (code, callback) => callback({ ok: true, message: 'OK' });

  context.executeJS(btn);

  assert.equal(identityChecks, 2);
  assert.equal(context.bindingBlocked, true);
  assert.equal(context.workbookBinding, null);
  assert.equal(elements.userInput.disabled, true);
  assert.equal(elements.sendBtn.disabled, true);
  assert.equal(messages.some(message => /Berhasil/.test(message)), false);
  assert.ok(messages.some(message => /workbook/i.test(message)));
});

test('save selama DELETE sukses tidak membuat ulang riwayat', () => {
  const { context, requests } = loadTaskpane();
  context.workbookBinding = { sessionId: 's1', workbookName: 'Nilai.xlsm', workbookPath: 'C:\\Data\\Nilai.xlsm' };
  context.chatHistory = [{ role: 'user', content: 'lama' }];
  let deleted;
  context.deleteHistory(ok => { deleted = ok; });
  assert.equal(requests.length, 1);
  assert.equal(requests[0].method, 'DELETE');

  context.chatHistory.push({ role: 'assistant', content: 'incidental' });
  context.saveHistory();
  assert.equal(requests.length, 1);

  requests[0].respond(200);
  assert.equal(deleted, true);
  assert.equal(requests.filter(request => request.method === 'POST').length, 0);
  assert.equal(context.pendingHistorySnapshot, null);
});

test('DELETE gagal melepas lock dan menyimpan snapshot terbaru', () => {
  const { context, requests } = loadTaskpane();
  context.workbookBinding = { sessionId: 's1', workbookName: 'Nilai.xlsm', workbookPath: 'C:\\Data\\Nilai.xlsm' };
  context.chatHistory = [{ role: 'user', content: 'lama' }];
  context.deleteHistory(() => {});
  context.chatHistory.push({ role: 'assistant', content: 'terbaru' });
  context.saveHistory();

  requests[0].respond(500);
  assert.equal(context.historyDeleteInFlight, false);
  assert.equal(requests.length, 2);
  assert.equal(requests[1].method, 'POST');
  assert.match(requests[1].body, /terbaru/);
});

test('taskpane: deteksi aksi vs pertanyaan penjelasan', () => {
  const { context } = loadTaskpane();
  // Tanya penjelasan -> looksLikeActionRequest harus false
  context.chatHistory = [{ role: 'user', content: 'apa bedanya format xlsm dan xlsx?' }];
  assert.equal(context.window.looksLikeActionRequest(), false);

  context.chatHistory = [{ role: 'user', content: 'kenapa rumus ini menghasilkan #REF!?' }];
  assert.equal(context.window.looksLikeActionRequest(), false);

  context.chatHistory = [{ role: 'user', content: 'jelaskan cara kerja macro di sheet rekap' }];
  assert.equal(context.window.looksLikeActionRequest(), false);

  // Aksi -> looksLikeActionRequest harus true
  context.chatHistory = [{ role: 'user', content: 'jumlahkan kolom D baris 5 sampai 20' }];
  assert.equal(context.window.looksLikeActionRequest(), true);

  context.chatHistory = [{ role: 'user', content: 'buat sheet baru bernama Analisis' }];
  assert.equal(context.window.looksLikeActionRequest(), true);

  context.chatHistory = [{ role: 'user', content: 'rapikan format tanggal di kolom B' }];
  assert.equal(context.window.looksLikeActionRequest(), true);
});
