const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');

// Isolasi: arahkan APPDATA ke folder sementara SEBELUM server di-load,
// supaya test tidak menyentuh %APPDATA%\ExcelAgent asli.
const tmpAppData = fs.mkdtempSync(path.join(os.tmpdir(), 'excelagent-test-'));
process.env.APPDATA = tmpAppData;

const {
  readProfil,
  writeProfil,
  readMemori,
  addMemori,
  deleteMemori,
  MEMORI_KATEGORI,
  sanitizeUsername,
  createUser,
  verifyUser,
  changePassword,
  createSession,
  getSessionUser,
  destroySession,
  multiUserEnabled,
} = require('../server');

test('profil default terisi identitas pemilik (mode tunggal)', () => {
  const p = readProfil(null);
  assert.equal(p.nama, 'Muhammad Rizki Pratama, S.Pd');
  assert.ok(p.peran.includes('Grogol Utara'));
  assert.ok(p.zonaWaktu.includes('Jakarta'));
});

test('profil bisa diubah dan tersimpan di disk', () => {
  const saved = writeProfil(null, { nama: 'Nama Tes', peran: 'Tester' });
  assert.equal(saved.nama, 'Nama Tes');
  assert.equal(readProfil(null).nama, 'Nama Tes');
  writeProfil(null, { nama: 'Muhammad Rizki Pratama, S.Pd', peran: 'Guru kelas IV, SDN Grogol Utara 09' });
});

test('profil menolak field kosong dan membatasi panjang', () => {
  const before = readProfil(null).nama;
  const saved = writeProfil(null, { nama: '   ', peran: 'x'.repeat(500) });
  assert.equal(saved.nama, before);
  assert.equal(saved.peran.length, 200);
});

test('memori: tambah, tolak duplikat, hapus', () => {
  assert.deepEqual(readMemori(null), []);
  const list = addMemori(null, 'preferensi', 'Format tanggal DD-MM-YYYY.');
  assert.equal(list.length, 1);
  assert.equal(list[0].kategori, 'preferensi');
  assert.ok(list[0].id);
  assert.ok(list[0].tanggal);
  assert.throws(() => addMemori(null, 'preferensi', 'Format tanggal DD-MM-YYYY.'), /sudah tercatat/);
  const afterDelete = deleteMemori(null, list[0].id);
  assert.deepEqual(afterDelete, []);
});

test('memori: kategori tidak dikenal dinormalisasi jadi fakta', () => {
  const list = addMemori(null, 'ngawur', 'Sesuatu.');
  assert.equal(list[0].kategori, 'fakta');
  deleteMemori(null, list[0].id);
});

test('memori: isi kosong ditolak', () => {
  assert.throws(() => addMemori(null, 'pola', '   '), /kosong/);
});

test('daftar kategori memori lengkap', () => {
  assert.deepEqual(MEMORI_KATEGORI, ['preferensi', 'koreksi', 'pola', 'fakta']);
});

test('sanitizeUsername menolak yang tidak valid', () => {
  assert.equal(sanitizeUsername('Budi Santoso'), '');
  assert.equal(sanitizeUsername('ab'), '');
  assert.equal(sanitizeUsername('budi_santoso-123'), 'budi_santoso-123');
  assert.equal(sanitizeUsername('  RIZKI  '), 'rizki');
});

test('multi-user: awalnya mati, buat user pertama mengaktifkannya', () => {
  assert.equal(multiUserEnabled(), false);
  const u = createUser('rizki', 'rahasia123');
  assert.equal(u, 'rizki');
  assert.equal(multiUserEnabled(), true);
  assert.throws(() => createUser('rizki', 'lain123'), /sudah dipakai/);
  assert.throws(() => createUser('ab', 'rahasia123'), /Username/);
  assert.throws(() => createUser('budi', '123'), /Password/);
});

test('multi-user: verifikasi password dan ganti password', () => {
  assert.equal(verifyUser('rizki', 'rahasia123'), 'rizki');
  assert.equal(verifyUser('rizki', 'salah'), null);
  assert.equal(verifyUser('tidakada', 'rahasia123'), null);
  changePassword('rizki', 'rahasia123', 'baru4567');
  assert.equal(verifyUser('rizki', 'baru4567'), 'rizki');
  assert.equal(verifyUser('rizki', 'rahasia123'), null);
  assert.throws(() => changePassword('rizki', 'salah', 'zzz9999'), /Password lama salah/);
});

test('multi-user: sesi token jalan dan bisa dihapus', () => {
  const token = createSession('rizki');
  assert.match(token, /^[0-9a-f]{64}$/);
  const req = { headers: { authorization: 'Bearer ' + token } };
  assert.equal(getSessionUser(req), 'rizki');
  assert.equal(getSessionUser({ headers: {} }), null);
  assert.equal(getSessionUser({ headers: { authorization: 'Bearer xxx' } }), null);
  destroySession(req);
  assert.equal(getSessionUser(req), null);
});

test('multi-user: data terisolasi per user', () => {
  createUser('budi', 'budi1234');
  writeProfil('rizki', { nama: 'Rizki Asli' });
  writeProfil('budi', { nama: 'Budi Asli' });
  assert.equal(readProfil('rizki').nama, 'Rizki Asli');
  assert.equal(readProfil('budi').nama, 'Budi Asli');
  addMemori('rizki', 'fakta', 'Rizki suka kopi.');
  assert.equal(readMemori('rizki').length, 1);
  assert.equal(readMemori('budi').length, 0);
  // file tersimpan di folder masing-masing
  assert.ok(fs.existsSync(path.join(tmpAppData, 'ExcelAgent', 'users', 'rizki', 'profil.json')));
  assert.ok(fs.existsSync(path.join(tmpAppData, 'ExcelAgent', 'users', 'budi', 'profil.json')));
  assert.ok(fs.existsSync(path.join(tmpAppData, 'ExcelAgent', 'users', 'rizki', 'memori.json')));
});
