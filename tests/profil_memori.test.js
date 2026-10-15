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
} = require('../server');

test('profil default terisi identitas pemilik', () => {
  const p = readProfil();
  assert.equal(p.nama, 'Muhammad Rizki Pratama, S.Pd');
  assert.ok(p.peran.includes('Grogol Utara'));
  assert.ok(p.zonaWaktu.includes('Jakarta'));
});

test('profil bisa diubah dan tersimpan di disk', () => {
  const saved = writeProfil({ nama: 'Nama Tes', peran: 'Tester' });
  assert.equal(saved.nama, 'Nama Tes');
  assert.equal(readProfil().nama, 'Nama Tes');
  const raw = JSON.parse(fs.readFileSync(path.join(tmpAppData, 'ExcelAgent', 'profil.json'), 'utf8'));
  assert.equal(raw.nama, 'Nama Tes');
  writeProfil({ nama: 'Muhammad Rizki Pratama, S.Pd', peran: 'Guru kelas IV, SDN Grogol Utara 09' });
});

test('profil menolak field kosong dan membatasi panjang', () => {
  const before = readProfil().nama;
  const saved = writeProfil({ nama: '   ', peran: 'x'.repeat(500) });
  assert.equal(saved.nama, before);
  assert.equal(saved.peran.length, 200);
});

test('memori: tambah, tolak duplikat, hapus', () => {
  assert.deepEqual(readMemori(), []);
  const list = addMemori('preferensi', 'Format tanggal DD-MM-YYYY.');
  assert.equal(list.length, 1);
  assert.equal(list[0].kategori, 'preferensi');
  assert.ok(list[0].id);
  assert.ok(list[0].tanggal);
  assert.throws(() => addMemori('preferensi', 'Format tanggal DD-MM-YYYY.'), /sudah tercatat/);
  const afterDelete = deleteMemori(list[0].id);
  assert.deepEqual(afterDelete, []);
});

test('memori: kategori tidak dikenal dinormalisasi jadi fakta', () => {
  const list = addMemori('ngawur', 'Sesuatu.');
  assert.equal(list[0].kategori, 'fakta');
  deleteMemori(list[0].id);
});

test('memori: isi kosong ditolak', () => {
  assert.throws(() => addMemori('pola', '   '), /kosong/);
});

test('daftar kategori memori lengkap', () => {
  assert.deepEqual(MEMORI_KATEGORI, ['preferensi', 'koreksi', 'pola', 'fakta']);
});
