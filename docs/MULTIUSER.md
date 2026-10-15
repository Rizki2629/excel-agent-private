# Multi-user (login per pengguna)

Excel Agent mendukung banyak pengguna di satu komputer. Setiap user login
dengan username + password sendiri dan mendapat:

- **Identitas sendiri** (`profil.json` miliknya) — model mengenalinya.
- **Memori sendiri** (`memori.json` miliknya) — belajar dari kebiasaannya.
- **Pengaturan sendiri** (`config.json` miliknya) — endpoint, model, API key,
  riwayat chat, dan daftar provider terpisah per user.

Data tersimpan di `%APPDATA%\ExcelAgent\users\<username>\`. Password di-hash
dengan scrypt (Node.js built-in, tanpa dependensi tambahan).

## Cara mengaktifkan

1. Buka taskpane, masuk ke **Pengaturan > Multi-user**.
2. Isi username + password pertama, klik **Aktifkan Multi-user**.
3. Sejak itu setiap membuka taskpane harus login dulu.
4. Tambah user lain di **Pengaturan > Multi-user > Tambah user**.

Catatan: data lama (sebelum multi-user) otomatis dipindahkan menjadi milik
user pertama, jadi tidak ada yang hilang.

## Tanpa multi-user

Selama belum ada akun yang dibuat, server berjalan dalam **mode tunggal**:
semua endpoint bebas diakses tanpa login, persis seperti versi lama.
Cocok untuk PC pribadi yang memang dipakai satu orang.

## Keamanan

- Server hanya mendengar di `127.0.0.1` (localhost): login hanya berlaku
  untuk pemakai komputer itu sendiri.
- Token sesi acak 256-bit, kedaluwarsa 7 hari, disimpan di `localStorage`
  (login tetap berlaku walau Excel ditutup; logout via badge user di header).
- Semua endpoint `/api/*` (kecuali `/api/auth/*` dan health) menolak request
  tanpa token begitu multi-user aktif (HTTP 401 + `needLogin`).
- File `users.json` diberi izin akses pemilik saja (chmod 600, best-effort).

## API

- `GET /api/auth/status` → `{ multiUser, loggedIn, username }` (publik)
- `POST /api/auth/setup` → buat user pertama `{ username, password }`
- `POST /api/auth/login` → `{ username, password }` → `{ token }`
- `POST /api/auth/logout` → batalkan token ini
- `POST /api/auth/change-password` → `{ oldPassword, newPassword }` (login)
- `GET /api/auth/users` → daftar user (login)
- `POST /api/auth/users` → tambah user (login)

Request terotentikasi memakai header `Authorization: Bearer <token>`.
