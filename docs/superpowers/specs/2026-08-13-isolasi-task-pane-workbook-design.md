# Desain Isolasi Task Pane per Workbook

## Tujuan

Setiap instance task pane hanya boleh membaca, mengubah, dan memverifikasi workbook yang membuka task pane tersebut. Perpindahan fokus ke workbook atau instance Excel lain tidak boleh mengubah target. Sistem harus mendukung beberapa instance Excel yang berjalan bersamaan.

## Batasan

- Workbook target wajib sudah pernah disimpan dan memiliki path penuh.
- Workbook yang belum disimpan tidak dapat menjalankan agent.
- Pengaturan API tetap global per pengguna.
- Riwayat percakapan dipisahkan per workbook.
- Tidak ada fallback ke `ActiveWorkbook` jika target tidak ditemukan atau identitas berubah.
- Perubahan hanya mencakup task pane Office.js dan jalur eksekusi Python yang dipakai task pane modern.

## Identitas Sesi dan Workbook

Saat task pane selesai diinisialisasi, task pane mengambil nama dan URL/path workbook dari konteks Office.js. Task pane membuat `sessionId` acak untuk umur instance pane tersebut dan menyimpan binding tetap berikut di memori:

```json
{
  "sessionId": "UUID",
  "workbookName": "Nilai Kelas 4.xlsm",
  "workbookPath": "C:\\Data\\Nilai Kelas 4.xlsm"
}
```

`workbookPath` dinormalisasi sebelum dibandingkan: path absolut Windows, separator konsisten, dan perbandingan tidak peka huruf besar-kecil. `workbookName` menjadi pemeriksaan tambahan, bukan identitas utama.

Task pane tidak memperbarui binding saat workbook lain menjadi aktif. Perubahan nama, lokasi, atau penutupan workbook setelah binding membuat sesi tidak valid. Pengguna harus membuka ulang task pane untuk membuat binding baru.

## Inisialisasi

1. Office.js menunggu `Office.onReady`.
2. Task pane mengambil properti file workbook.
3. Jika path kosong atau hanya nama sementara seperti `Book1`, task pane masuk status terblokir.
4. Status terblokir menonaktifkan pengiriman pesan dan menampilkan instruksi untuk menyimpan workbook lalu membuka ulang task pane.
5. Jika path valid, task pane membuat binding dan memuat riwayat milik workbook tersebut.

Kegagalan membaca properti file juga menghasilkan status terblokir. Agent tidak boleh menebak target dari workbook aktif.

## Alur Eksekusi

Setiap panggilan `/api/exec-python` mengirim envelope JSON:

```json
{
  "code": "...",
  "target": {
    "sessionId": "UUID",
    "workbookName": "Nilai Kelas 4.xlsm",
    "workbookPath": "C:\\Data\\Nilai Kelas 4.xlsm"
  }
}
```

Server memvalidasi struktur request, lalu memasukkan pekerjaan ke antrean eksekusi COM tunggal. Serialisasi diperlukan karena beberapa task pane dapat mengirim pekerjaan bersamaan dan automasi COM Excel tidak aman dijalankan tanpa koordinasi.

Server meneruskan envelope lengkap ke `exec_python.py` melalui stdin. Timeout dan respons tetap memakai format JSON terstruktur.

## Resolusi Multi-Instance Excel

`exec_python.py` tidak memakai `Dispatch("Excel.Application")` untuk memilih target. Proses mencari objek Excel yang sedang berjalan melalui Windows Running Object Table, memeriksa koleksi `Workbooks` pada setiap instance, lalu mencocokkan `Workbook.FullName` dengan `target.workbookPath`.

Hasil resolusi harus tepat satu workbook:

- Nol kecocokan: gagal dengan pesan bahwa workbook target ditutup, dipindah, atau diganti nama.
- Lebih dari satu kecocokan: gagal karena identitas ambigu.
- Satu kecocokan: sediakan `xl` dan `wb` pada lingkungan eksekusi.

Prompt agent diubah agar semua operasi memakai `wb`, misalnya `ws = wb.ActiveSheet`. Referensi dan contoh `xl.ActiveWorkbook` dihapus. Variabel `xl` tetap tersedia untuk API tingkat aplikasi yang diperlukan, tetapi bukan untuk memilih workbook.

Sebelum dan sesudah kode agent berjalan, executor memeriksa bahwa `wb.FullName` masih sama dengan binding. Perbedaan membatalkan hasil sebagai kegagalan target. Tidak ada percobaan ulang ke workbook aktif.

## Verifikasi

Verifikasi setelah eksekusi dijalankan melalui konteks Office.js milik task pane. Verifikasi tidak boleh mengambil konteks dari COM `ActiveWorkbook`.

Sebelum menampilkan status sukses, task pane membandingkan kembali properti file konteksnya dengan binding. Jika path berubah, sesi dianggap tidak valid dan hasil tidak dilaporkan sebagai sukses terverifikasi.

## Riwayat Percakapan

Konfigurasi server menyimpan riwayat sebagai map berdasarkan hash path workbook yang sudah dinormalisasi, bukan satu array `chatHistory` global. Path mentah tidak perlu menjadi nama properti JSON.

```json
{
  "chatHistoryByWorkbook": {
    "<hash-path>": []
  }
}
```

Task pane mengirim key workbook saat memuat, menyimpan, atau menghapus riwayat. Batas 50 pesan tetap berlaku untuk setiap workbook. API key, endpoint, model, dan `autoRun` tetap global.

## Penanganan Error

Semua kegagalan target bersifat fail-closed:

- Workbook belum disimpan: tombol kirim dinonaktifkan.
- Workbook target tidak ditemukan: kode tidak dijalankan.
- Identitas ambigu: kode tidak dijalankan.
- Workbook berubah path atau nama: sesi dihentikan dan pane harus dibuka ulang.
- Request tanpa metadata target: server mengembalikan HTTP 400.
- Eksekusi lain sedang berjalan: request menunggu antrean, bukan berjalan paralel.

Pesan error memakai bahasa Indonesia dan menyebut tindakan yang diperlukan tanpa menampilkan detail sensitif seperti API key.

## Keamanan

Endpoint eksekusi menerima kode arbitrer dan hanya ditujukan untuk task pane lokal. Perubahan ini harus membatasi origin API ke origin add-in lokal, bukan `Access-Control-Allow-Origin: *`, serta memvalidasi bentuk request sebelum membuat child process. Penguatan ini tidak menggantikan sandbox; kode agent tetap memiliki hak akses proses pengguna.

## Skenario Uji

1. Buka Workbook A dan task pane A, lalu aktifkan Workbook B. Perintah pane A hanya mengubah A.
2. Buka pane A dan pane B. Eksekusi dan riwayat masing-masing tetap terpisah.
3. Buka A dan B pada dua instance Excel berbeda. Setiap pane menemukan target berdasarkan `FullName`.
4. Buka workbook baru yang belum disimpan. Agent terblokir sampai workbook disimpan dan task pane dibuka ulang.
5. Tutup target setelah pane terikat. Eksekusi gagal tanpa menyentuh workbook lain.
6. Ganti nama atau pindahkan target setelah pane terikat. Eksekusi gagal dan meminta pane dibuka ulang.
7. Kirim request dari dua pane bersamaan. Server menjalankannya berurutan terhadap target masing-masing.
8. Kirim request tanpa `target` atau dengan path tidak valid. Server menolak sebelum menjalankan Python.
9. Cari seluruh jalur task pane modern. Tidak ada pemilihan target melalui `ActiveWorkbook` atau fallback berdasarkan nama saja.

## Di Luar Cakupan

- Mendukung workbook yang belum disimpan melalui marker tersembunyi.
- Memigrasikan jalur VBA UserForm lama.
- Memisahkan pengaturan API per workbook.
- Menjalankan beberapa mutasi COM secara paralel.
- Mempertahankan sesi setelah workbook dipindah atau diganti nama.
