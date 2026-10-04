# Excel Agent

Asisten AI di dalam Microsoft Excel. Ketik perintah bahasa Indonesia di panel chat — agent menulis dan menjalankan kode Python (win32com) langsung ke workbook yang terbuka: rekap data, format, chart, analisis gambar, dan lainnya. Semua berjalan lokal di komputermu.

## Syarat

- Windows + Microsoft Excel desktop
- Node.js 18+
- Python 3.12+ (centang **"Add Python to PATH"** saat install)
- OpenCode CLI (untuk mode analisis mendalam)

## Instalasi pertama kali

Double-click **`SETUP.bat`**, ikuti sampai selesai. Add-in terdaftar permanen di Excel.

## Pakai harian

1. Double-click **`START.bat`** → pilih **`[1] Jalankan Semua`**.
2. Di Excel: tab **HOME** → grup **Excel Agent** → **Buka Asisten AI**.
3. Ketik perintah di chat, contoh:
   - "Buat chart penjualan bulanan dari data di Sheet1"
   - "Format sel A1:F10 dengan warna biru dan tebal"
   - "Hitung rata-rata kolom B"
4. Selesai kerja: **`STOP.bat`**.

## Menu START.bat

| Opsi | Fungsi |
|------|--------|
| 1 | Jalankan Semua (server + buka Excel) |
| 2 | Server saja |
| 3 | Load add-in ke Excel |
| 4 | Bersihkan cache (kalau taskpane tidak muncul) |
| 5 | Cek status server / Excel / Node.js |

## Troubleshooting

- **Taskpane tidak muncul / masih versi lama** → `START.bat` → `[4] Bersihkan cache`, lalu `[1]`.
- **"Session workbook is invalid"** → tunggu 5–10 detik (auto-recovery), atau cek `[5]`.
- **Server mati** → `START.bat` → `[2]`.

## Tab di taskpane

💬 **Chat** (utama) · ⚡ **Cepat** (aksi sekali klik) · 🖼️ **Gambar** (analisis screenshot, maks 5 gambar) · ⚙️ **Pengaturan** (provider, model, API key)

## Pasang di komputer lain (clone)

Repo ini **private** — siapkan Personal Access Token GitHub (classic, scope `repo`) untuk login saat clone.

```powershell
git clone https://github.com/Rizki2629/excel-agent-private.git
cd excel-agent-private
```

Lalu double-click **`SETUP.bat`** (sekali saja). Selanjutnya pakai harian seperti di atas (`START.bat` → `[1]`).

## Dokumentasi teknis

Lihat folder [`docs/`](docs/): arsitektur (`TECH_ARCH.md`), deployment (`DEPLOYMENT.md`), referensi cepat (`QUICK_REFERENCE.md`), panduan file .bat (`BATCH_FILES_GUIDE.md`), dan arsip pengembangan.
