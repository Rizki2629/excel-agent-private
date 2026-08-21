# Desain Lampiran Gambar Taskpane

## Tujuan

Taskpane menerima maksimal lima screenshot per pesan melalui pemilih file atau paste clipboard. OpenCode menganalisis gambar sebagai referensi bersama prompt teks dan konteks workbook.

## Antarmuka

- Tombol paperclip membuka input file tersembunyi dengan format PNG, JPEG, dan WEBP.
- Paste gambar pada textarea menambahkan gambar tanpa menghapus teks clipboard lain.
- Preview thumbnail tampil di atas textarea. Setiap thumbnail memiliki tombol hapus.
- Batas: lima gambar per pesan dan 5 MB per gambar.
- Tombol kirim aktif bila ada teks atau gambar.
- Pesan tanpa teks memakai prompt `Analisis gambar ini sebagai referensi workbook.`

## Aliran Data

Taskpane membaca gambar memakai `FileReader.readAsDataURL()` dan mengirim array `images` berisi `name`, `type`, dan `data` Base64 ke `/api/opencode`. Server memvalidasi jumlah, MIME type, Base64, dan ukuran hasil decode. Server menulis file ke direktori temporary per request, menambahkan setiap path sebagai argumen `--file` pada `opencode run`, lalu menghapus seluruh file pada sukses, error, timeout, atau koneksi client terputus.

## Riwayat

Riwayat permanen hanya menyimpan teks dan metadata nama/jumlah gambar. Data Base64 dan file tidak disimpan setelah analisis. Thumbnail pesan memakai data URL selama taskpane tetap terbuka.

## Error

- Format tidak didukung: gambar ditolak sebelum preview.
- Ukuran di atas 5 MB: gambar ditolak sebelum preview.
- Lebih dari lima gambar: gambar tambahan ditolak.
- Payload server tidak valid: HTTP 400 dengan pesan Indonesia.
- Cleanup file temporary selalu berjalan melalui `finally`.

## Verifikasi

- Unit test taskpane: validasi, paste, tombol kirim, payload, dan pembersihan preview.
- Unit test server: decode, batas, MIME, argumen `--file`, dan cleanup.
- End-to-end: kirim gambar kecil ke endpoint dan pastikan OpenCode menyebut isi visual.
