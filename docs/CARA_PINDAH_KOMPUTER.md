# 💻 Pasang Excel Agent di Komputer Lain

## 1. Clone repo

Repo ini **private**. Siapkan Personal Access Token GitHub (classic, scope `repo`) — dipakai sebagai password saat git meminta login.

```powershell
git clone https://github.com/Rizki2629/excel-agent-private.git
cd excel-agent-private
```

## 2. Setup sekali jalan

Double-click **`SETUP.bat`** (klik kanan → *Run as Administrator* lebih baik). Otomatis mengerjakan:

- ✅ Cek Node.js & Python (+ install `pywin32` bila belum ada)
- ✅ `npm install`
- ✅ Sertifikat HTTPS lokal
- ✅ Daftarkan add-in ke Registry Office Windows (permanen)
- ✅ Shortcut **Excel Agent Server** di Desktop

## 3. Pakai harian

`START.bat` → `[1] Jalankan Semua` → di Excel tab **HOME** → **Buka Asisten AI**. Selesai kerja: `STOP.bat`.

## Troubleshooting

| Masalah | Solusi |
|---|---|
| Tombol tidak muncul di Excel | Tutup Excel, jalankan `SETUP.bat` sekali lagi |
| Badge "Server Offline" | `START.bat` → `[2]` (server harus jalan di port 3000) |
| Error pywin32 / win32com | Buka Command Prompt: `pip install pywin32` |
| Peringatan sertifikat SSL | `npx office-addin-dev-certs install --machine` |
| Taskpane tidak muncul / versi lama | `START.bat` → `[4] Bersihkan cache`, lalu `[1]` |
