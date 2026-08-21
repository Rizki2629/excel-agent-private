# Laporan Final Fix Wave

## Perbaikan

- API key tidak lagi keluar dari `GET /api/settings`, tidak disimpan di `localStorage`, dan proxy memakai key server saat request tidak mengirim key.
- Penyimpanan setting kosong mempertahankan key; `clearApiKey: true` menghapus key secara eksplisit.
- Semua `targetError` menghentikan eksekusi task pane dan memblokir sesi sampai pane dibuka ulang.
- Kegagalan membaca koleksi COM `Workbooks` diklasifikasikan sebagai `TargetWorkbookError`.
- Save history satu pane diserialkan dan snapshot terbaru dikoaleskan; delete menunggu save aktif dan membatalkan snapshot pending.
- Body request dibatasi 1 MiB dan antrean COM dibatasi 20 pekerjaan dengan respons 429 saat penuh.
- Nama workbook divalidasi terhadap basename path sebelum child process dibuat.
- Dokumentasi menjelaskan executor bukan sandbox.

## Risiko Tersisa

- Dua task pane untuk workbook yang sama masih memakai full-snapshot last-write-wins. Konflik lintas pane belum diberi merge/revision server karena snapshot dari sesi berbeda tidak punya urutan global yang aman.
- Kode Python hasil model tetap memiliki izin proses pengguna. Binding `wb` hanya menjaga identitas target, bukan membatasi kemampuan kode berbahaya.
