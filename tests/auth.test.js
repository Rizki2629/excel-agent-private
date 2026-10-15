const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { Readable } = require('node:stream');

// Isolasi APPDATA sebelum server di-load.
const tmpAppData = fs.mkdtempSync(path.join(os.tmpdir(), 'excelagent-auth-'));
process.env.APPDATA = tmpAppData;

const { createRequestHandler } = require('../server');

function request(method, url, body, token) {
  const req = Readable.from(body === undefined ? [] : [Buffer.from(JSON.stringify(body))]);
  req.method = method;
  req.url = url;
  req.headers = {};
  if (token) req.headers.authorization = 'Bearer ' + token;
  return req;
}

function response() {
  return {
    statusCode: undefined,
    headers: undefined,
    body: '',
    writeHead(statusCode, headers) { this.statusCode = statusCode; this.headers = headers; },
    end(body = '') { this.body = body; },
  };
}

function json(res) {
  return JSON.parse(res.body);
}

test('auth: status awal mode tunggal, tanpa user', async () => {
  const handler = createRequestHandler();
  const res = response();
  await handler(request('GET', '/api/auth/status'), res);
  assert.equal(res.statusCode, 200);
  const data = json(res);
  assert.equal(data.multiUser, false);
  assert.equal(data.loggedIn, false);
});

test('auth: /api/settings bebas di mode tunggal', async () => {
  const handler = createRequestHandler();
  const res = response();
  await handler(request('GET', '/api/settings'), res);
  assert.equal(res.statusCode, 200);
});

test('auth: setup user pertama mengaktifkan multi-user', async () => {
  const handler = createRequestHandler();
  const res = response();
  await handler(request('POST', '/api/auth/setup', { username: 'guru1', password: 'pass1234' }), res);
  assert.equal(res.statusCode, 200);
  const data = json(res);
  assert.equal(data.username, 'guru1');
  assert.match(data.token, /^[0-9a-f]{64}$/);

  // setup kedua ditolak
  const res2 = response();
  await handler(request('POST', '/api/auth/setup', { username: 'guru2', password: 'pass1234' }), res2);
  assert.equal(res2.statusCode, 403);

  const res3 = response();
  await handler(request('GET', '/api/auth/status'), res3);
  assert.equal(json(res3).multiUser, true);
});

test('auth: endpoint sensitif butuh token setelah multi-user aktif', async () => {
  const handler = createRequestHandler();
  const res = response();
  await handler(request('GET', '/api/settings'), res);
  assert.equal(res.statusCode, 401);
  assert.equal(json(res).needLogin, true);

  const res2 = response();
  await handler(request('GET', '/api/memori'), res2);
  assert.equal(res2.statusCode, 401);
});

test('auth: login benar dapat token, login salah ditolak', async () => {
  const handler = createRequestHandler();
  const resBad = response();
  await handler(request('POST', '/api/auth/login', { username: 'guru1', password: 'salah' }), resBad);
  assert.equal(resBad.statusCode, 401);

  const res = response();
  await handler(request('POST', '/api/auth/login', { username: 'guru1', password: 'pass1234' }), res);
  assert.equal(res.statusCode, 200);
  const token = json(res).token;

  const res2 = response();
  await handler(request('GET', '/api/settings', undefined, token), res2);
  assert.equal(res2.statusCode, 200);

  // token ngawur ditolak
  const res3 = response();
  await handler(request('GET', '/api/settings', undefined, '0'.repeat(64)), res3);
  assert.equal(res3.statusCode, 401);
});

test('auth: data terisolasi antar user via API', async () => {
  const handler = createRequestHandler();
  // login sebagai guru1
  const loginRes = response();
  await handler(request('POST', '/api/auth/login', { username: 'guru1', password: 'pass1234' }), loginRes);
  const token1 = json(loginRes).token;

  // tambah user guru2
  const addRes = response();
  await handler(request('POST', '/api/auth/users', { username: 'guru2', password: 'pass5678' }, token1), addRes);
  assert.equal(addRes.statusCode, 200);

  // daftar user terlihat
  const listRes = response();
  await handler(request('GET', '/api/auth/users', undefined, token1), listRes);
  const usernames = json(listRes).users.map((u) => u.username).sort();
  assert.deepEqual(usernames, ['guru1', 'guru2']);

  // guru1 simpan memori
  const memRes = response();
  await handler(request('POST', '/api/memori', { kategori: 'fakta', isi: 'Milik guru1.' }, token1), memRes);
  assert.equal(memRes.statusCode, 200);

  // login sebagai guru2 -> memori kosong (terisolasi)
  const login2 = response();
  await handler(request('POST', '/api/auth/login', { username: 'guru2', password: 'pass5678' }), login2);
  const token2 = json(login2).token;
  const memRes2 = response();
  await handler(request('GET', '/api/memori', undefined, token2), memRes2);
  assert.deepEqual(json(memRes2).memori, []);
});

test('auth: logout membatalkan token', async () => {
  const handler = createRequestHandler();
  const loginRes = response();
  await handler(request('POST', '/api/auth/login', { username: 'guru1', password: 'pass1234' }), loginRes);
  const token = json(loginRes).token;

  const outRes = response();
  await handler(request('POST', '/api/auth/logout', undefined, token), outRes);
  assert.equal(outRes.statusCode, 200);

  const res = response();
  await handler(request('GET', '/api/settings', undefined, token), res);
  assert.equal(res.statusCode, 401);
});

test('auth: ganti password', async () => {
  const handler = createRequestHandler();
  const loginRes = response();
  await handler(request('POST', '/api/auth/login', { username: 'guru2', password: 'pass5678' }), loginRes);
  const token = json(loginRes).token;

  const cpRes = response();
  await handler(request('POST', '/api/auth/change-password',
    { oldPassword: 'pass5678', newPassword: 'baru9999' }, token), cpRes);
  assert.equal(cpRes.statusCode, 200);

  const loginNew = response();
  await handler(request('POST', '/api/auth/login', { username: 'guru2', password: 'baru9999' }), loginNew);
  assert.equal(loginNew.statusCode, 200);

  const loginOld = response();
  await handler(request('POST', '/api/auth/login', { username: 'guru2', password: 'pass5678' }), loginOld);
  assert.equal(loginOld.statusCode, 401);
});
