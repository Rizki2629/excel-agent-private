# Identitas & Memori Jangka Panjang

Mulai versi ini, Excel Agent mengenali pemiliknya dan belajar dari waktu ke waktu.
Keduanya tersimpan di `%APPDATA%\ExcelAgent\` (di luar repo) sehingga tidak ikut
ter-commit, tidak hilang saat update kode, dan konsisten di semua komputer
setelah SETUP.bat dijalankan.

## Identitas (`profil.json`)

Berisi nama, peran, zona waktu, bahasa, dan gaya jawab. Disuntik ke system
prompt setiap chat sebagai bagian `### PROFIL PENGGUNA ###`, jadi model tahu
siapa yang dilayani tanpa perlu dijelaskan ulang.

- Nilai default diisi saat SETUP.bat pertama kali (tidak ditimpa bila sudah ada).
- Bisa diubah kapan saja lewat Pengaturan (bagian "Identitas Pengguna") di taskpane,
  atau manual dengan mengedit `%APPDATA%\ExcelAgent\profil.json`.
- API: `GET /api/profil`, `POST /api/profil`.

## Memori (`memori.json`)

Daftar pelajaran dari sesi sebelumnya: preferensi, koreksi, pola sukses, dan
fakta. Disuntik ke system prompt sebagai `### MEMORI JANGKA PANJANG ###`
sehingga agent makin pintar tanpa perlu dilatih ulang.

### Cara mencatat

1. **Otomatis oleh agent.** Jika kamu bilang "ingat ini", "catat", "jangan lupa",
   atau mengoreksi agent ("bukan begitu, yang benar..."), agent menyimpan
   pelajarannya sendiri lewat blok ` ```memory ` pada jawabannya:

   ````
   ```memory
   kategori: preferensi
   isi: Format tanggal di daftar hadir selalu DD-MM-YYYY.
   ```
   ````
   Kategori: `preferensi`, `koreksi`, `pola`, `fakta`.
2. **Manual.** Di Pengaturan ada daftar memori; entri bisa dihapus satu per satu.
3. **API.** `GET /api/memori`, `POST /api/memori` (`{kategori, isi}`),
   `DELETE /api/memori` (`{id}`).

### Aturan main

- Maksimal 200 entri; yang terlama dibuang otomatis bila penuh.
- Duplikat persis ditolak.
- **Dilarang menyimpan API key, password, token, atau secret apa pun.**
  Server menolak request yang mengandung kata kunci kredensial.

## Reproducibility (hasil sama di semua komputer)

Keluhan "install di komputer lain hasilnya beda" diatasi dengan:

- `requirements.txt` mengunci `pywin32==312` (bukan `>=` tanpa batas).
- SETUP.bat memakai `npm ci` (bukan `npm install`) sehingga versi paket
  persis mengikuti `package-lock.json`.
- SETUP.bat mengecek versi minimum: Node.js 18+, Python 3.10+.
- Tahap 8 SETUP.bat adalah self-test: syntax check semua file JS utama,
  keberadaan `manifest.xml`, dan import `win32com.client`. Setup hanya
  dinyatakan berhasil bila semua cek lolos.
