const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { Readable } = require('node:stream');
const {
  createRequestHandler,
  publicSettings,
  migrateProviderModels
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

// Config store palsu supaya test tidak menyentuh config.json milik user.
function fakeStore(initial) {
  let state = JSON.parse(JSON.stringify(initial));
  return {
    read: () => JSON.parse(JSON.stringify(state)),
    write: (partial) => { state = Object.assign(state, JSON.parse(JSON.stringify(partial))); return JSON.parse(JSON.stringify(state)); },
    peek: () => state
  };
}

test('publicSettings memaparkan models per provider tanpa API key', () => {
  const out = publicSettings({
    apiKey: 'sk-rahasia',
    providers: {
      A: { apiEndpoint: 'https://a.test/v1', model: 'm1', models: ['m1', 'm2'], apiKey: 'sk-a' },
      B: { apiEndpoint: 'https://b.test/v1', model: 'n1', models: ['n1'], apiKey: '' }
    }
  });
  assert.deepEqual(out.providers.A.models, ['m1', 'm2']);
  assert.deepEqual(out.providers.B.models, ['n1']);
  assert.equal(out.providers.A.hasKey, true);
  assert.equal(out.providers.B.hasKey, false);
  assert.equal(JSON.stringify(out).includes('sk-rahasia'), false);
  assert.equal(JSON.stringify(out).includes('sk-a'), false);
});

test('publicSettings memberi array models kosong bila provider lama belum punya field itu', () => {
  const out = publicSettings({ providers: { Lama: { apiEndpoint: 'x', model: 'm' } } });
  assert.deepEqual(out.providers.Lama.models, []);
});

test('migrasi mengisi models dari model aktif provider lama tanpa menghapus modelList global', () => {
  const store = fakeStore({
    modelList: ['arsip-1', 'arsip-2', 'arsip-3'],
    activeProvider: 'Lama',
    chatHistoryByWorkbook: { 'c:/a.xlsx': [{ role: 'user', content: 'halo' }] },
    providers: {
      Lama: { apiEndpoint: 'https://lama.test/v1', model: 'model-lama', apiKey: 'sk-lama' },
      Baru: { apiEndpoint: 'https://baru.test/v1', model: 'model-baru', models: ['model-baru', 'ekstra'] }
    }
  });

  const result = migrateProviderModels(store.read, store.write);
  assert.equal(result.migrated, 1);

  const after = store.peek();
  // provider lama kini punya daftar sendiri
  assert.deepEqual(after.providers.Lama.models, ['model-lama']);
  // provider yang sudah punya daftar tidak diubah
  assert.deepEqual(after.providers.Baru.models, ['model-baru', 'ekstra']);
  // DATA LAMA TIDAK HILANG
  assert.deepEqual(after.modelList, ['arsip-1', 'arsip-2', 'arsip-3']);
  assert.equal(after.providers.Lama.apiKey, 'sk-lama');
  assert.equal(after.providers.Lama.apiEndpoint, 'https://lama.test/v1');
  assert.equal(after.activeProvider, 'Lama');
  assert.deepEqual(after.chatHistoryByWorkbook['c:/a.xlsx'], [{ role: 'user', content: 'halo' }]);
});

test('migrasi idempoten: dijalankan dua kali tidak mengubah apa pun lagi', () => {
  const store = fakeStore({
    modelList: ['a'],
    providers: { P: { apiEndpoint: 'e', model: 'm' } }
  });
  const first = migrateProviderModels(store.read, store.write);
  assert.equal(first.migrated, 1);
  const snapshot = JSON.stringify(store.peek());
  const second = migrateProviderModels(store.read, store.write);
  assert.equal(second.migrated, 0);
  assert.equal(JSON.stringify(store.peek()), snapshot);
});

test('migrasi provider tanpa model menghasilkan daftar kosong, bukan crash', () => {
  const store = fakeStore({ providers: { Kosong: { apiEndpoint: 'e', model: '' } } });
  migrateProviderModels(store.read, store.write);
  assert.deepEqual(store.peek().providers.Kosong.models, []);
});

test('migrasi aman saat belum ada provider sama sekali', () => {
  const store = fakeStore({ modelList: ['x'], providers: {} });
  const result = migrateProviderModels(store.read, store.write);
  assert.equal(result.migrated, 0);
  assert.deepEqual(store.peek().modelList, ['x']);
});

test('endpoint tambah model provider hanya menyentuh provider yang dituju', async () => {
  const store = fakeStore({
    apiKey: '',
    modelList: ['arsip-global'],
    providers: {
      Satu: { apiEndpoint: 'https://satu.test/v1', model: 'a1', models: ['a1'], apiKey: 'sk-1' },
      Dua: { apiEndpoint: 'https://dua.test/v1', model: 'b1', models: ['b1'], apiKey: 'sk-2' }
    }
  });
  const handler = createRequestHandler({ readConfig: store.read, writeConfig: store.write });
  const res = response();
  await handler(request('POST', '/api/providers/models', { name: 'Satu', model: 'a2' }), res);
  assert.equal(res.statusCode, 200);
  const data = JSON.parse(res.body);
  assert.deepEqual(data.models, ['a1', 'a2']);

  const after = store.peek();
  assert.deepEqual(after.providers.Satu.models, ['a1', 'a2']);
  // provider lain TIDAK tersentuh
  assert.deepEqual(after.providers.Dua.models, ['b1']);
  assert.equal(after.providers.Dua.apiKey, 'sk-2');
  // arsip global tetap
  assert.deepEqual(after.modelList, ['arsip-global']);
});

test('endpoint hapus model provider tidak menyentuh provider lain atau arsip global', async () => {
  const store = fakeStore({
    modelList: ['arsip-global'],
    providers: {
      Satu: { apiEndpoint: 'e1', model: 'a1', models: ['a1', 'a2'], apiKey: 'sk-1' },
      Dua: { apiEndpoint: 'e2', model: 'a2', models: ['a2'], apiKey: 'sk-2' }
    }
  });
  const handler = createRequestHandler({ readConfig: store.read, writeConfig: store.write });
  const res = response();
  await handler(request('DELETE', '/api/providers/models', { name: 'Satu', model: 'a2' }), res);
  assert.equal(res.statusCode, 200);

  const after = store.peek();
  assert.deepEqual(after.providers.Satu.models, ['a1']);
  // model bernama sama di provider lain tetap ada
  assert.deepEqual(after.providers.Dua.models, ['a2']);
  assert.deepEqual(after.modelList, ['arsip-global']);
});

test('tambah model duplikat tidak menggandakan entri', async () => {
  const store = fakeStore({ providers: { P: { apiEndpoint: 'e', model: 'm', models: ['m'] } } });
  const handler = createRequestHandler({ readConfig: store.read, writeConfig: store.write });
  const res = response();
  await handler(request('POST', '/api/providers/models', { name: 'P', model: 'm' }), res);
  assert.equal(res.statusCode, 200);
  assert.deepEqual(store.peek().providers.P.models, ['m']);
});

test('endpoint model provider menolak provider tak dikenal dan model kosong', async () => {
  const store = fakeStore({ providers: { P: { apiEndpoint: 'e', model: 'm', models: ['m'] } } });
  const handler = createRequestHandler({ readConfig: store.read, writeConfig: store.write });

  const res404 = response();
  await handler(request('POST', '/api/providers/models', { name: 'TidakAda', model: 'x' }), res404);
  assert.equal(res404.statusCode, 404);
  assert.equal(JSON.parse(res404.body).ok, false);

  const res400 = response();
  await handler(request('POST', '/api/providers/models', { name: 'P', model: '   ' }), res400);
  assert.equal(res400.statusCode, 400);
  assert.equal(JSON.parse(res400.body).ok, false);
});

test('simpan provider mempertahankan daftar model lama bila models tidak dikirim', async () => {
  const store = fakeStore({
    providers: { P: { apiEndpoint: 'e', model: 'm1', models: ['m1', 'm2', 'm3'], apiKey: 'sk-p' } }
  });
  const handler = createRequestHandler({ readConfig: store.read, writeConfig: store.write });
  const res = response();
  // Simpan ulang hanya mengubah endpoint; daftar model & key wajib utuh.
  await handler(request('POST', '/api/providers', { name: 'P', apiEndpoint: 'e-baru' }), res);
  assert.equal(res.statusCode, 200);

  const after = store.peek();
  assert.deepEqual(after.providers.P.models, ['m1', 'm2', 'm3']);
  assert.equal(after.providers.P.apiKey, 'sk-p');
  assert.equal(after.providers.P.apiEndpoint, 'e-baru');
});

test('simpan provider dengan model baru otomatis memasukkannya ke daftar provider itu', async () => {
  const store = fakeStore({ providers: {} });
  const handler = createRequestHandler({ readConfig: store.read, writeConfig: store.write });
  const res = response();
  await handler(request('POST', '/api/providers', { name: 'Baru', apiEndpoint: 'e', model: 'model-x' }), res);
  assert.equal(res.statusCode, 200);
  assert.deepEqual(store.peek().providers.Baru.models, ['model-x']);
});

test('aktivasi provider tidak mengubah daftar model provider mana pun', async () => {
  const store = fakeStore({
    activeProvider: 'Satu',
    providers: {
      Satu: { apiEndpoint: 'e1', model: 'a1', models: ['a1', 'a2'], apiKey: 'k1' },
      Dua: { apiEndpoint: 'e2', model: 'b1', models: ['b1', 'b2'], apiKey: 'k2' }
    }
  });
  const handler = createRequestHandler({ readConfig: store.read, writeConfig: store.write });
  const res = response();
  await handler(request('POST', '/api/providers/activate', { name: 'Dua' }), res);
  assert.equal(res.statusCode, 200);

  const after = store.peek();
  assert.equal(after.activeProvider, 'Dua');
  assert.equal(after.model, 'b1');
  assert.deepEqual(after.providers.Satu.models, ['a1', 'a2']);
  assert.deepEqual(after.providers.Dua.models, ['b1', 'b2']);
});

test('hapus satu provider tidak menghapus provider lain maupun arsip global', async () => {
  const store = fakeStore({
    modelList: ['arsip'],
    activeProvider: 'Satu',
    providers: {
      Satu: { apiEndpoint: 'e1', model: 'a1', models: ['a1'] },
      Dua: { apiEndpoint: 'e2', model: 'b1', models: ['b1'] }
    }
  });
  const handler = createRequestHandler({ readConfig: store.read, writeConfig: store.write });
  const res = response();
  await handler(request('DELETE', '/api/providers', { name: 'Satu' }), res);
  assert.equal(res.statusCode, 200);

  const after = store.peek();
  assert.equal(after.providers.Satu, undefined);
  assert.deepEqual(after.providers.Dua.models, ['b1']);
  assert.deepEqual(after.modelList, ['arsip']);
  assert.equal(after.activeProvider, '');
});

test('server menjalankan migrasi model-per-provider saat start', () => {
  const source = fs.readFileSync(path.join(__dirname, '..', 'server.js'), 'utf8');
  assert.match(source, /migrateProviderModels\(/);
  // modelList global tidak pernah dihapus dari config
  assert.doesNotMatch(source, /delete\s+\w*[Cc]onfig\.modelList/);
});
