# Panduan File .bat — Excel Agent

> Diperbarui Okt 2026: 8 file .bat lama digabung menjadi 3 file.

## File yang ada

| File | Fungsi | Kapan dipakai |
|------|--------|---------------|
| `START.bat` | Menu utama: [1] Jalankan Semua · [2] Server saja · [3] Load add-in · [4] Bersihkan cache · [5] Cek status | Setiap hari |
| `STOP.bat` | Hentikan server + Node.js, opsi tutup Excel | Selesai kerja |
| `SETUP.bat` | Setup komputer baru sekali jalan (Node/Python check, npm install, sertifikat HTTPS, registrasi Registry Office, config, shortcut Desktop) | Sekali saja di komputer baru |

## Detail menu START.bat

- **[1] Jalankan Semua** — hentikan proses lama → bersihkan cache Office WEF → jalankan `node server.js` (port 3000) → sideload add-in → buka Excel + workbook (jika path `C:\Users\acer\agent_excel\ADMINISTRSAI KELAS.xlsm` ada).
- **[2] Server saja** — cek Node.js, jalankan `node server.js` di window ini (CTRL+C untuk berhenti).
- **[3] Load Add-in ke Excel** — `npx office-addin-debugging start manifest.xml` + buka workbook. Server harus sudah jalan.
- **[4] Bersihkan cache** — tutup Excel, hapus cache Office WEF + Chrome. Pakai kalau taskpane tidak muncul / versi lama.
- **[5] Cek status** — cek port 3000, proses Excel, proses Node.js.

## Riwayat

File lama (dihapus Okt 2026, riwayat ada di git): `LAUNCH_ALL.bat`, `MENU_CONTROL.bat`, `START_SERVER.bat`, `LOAD_ADDIN.bat`, `CLEAR_CACHE.bat`, `STOP_ALL.bat`, `INSTALL_PERMANENT.bat`, `SETUP_KOMPUTER_BARU.bat`.
`SETUP.bat` adalah `SETUP_KOMPUTER_BARU.bat` yang diganti nama (langkah `INSTALL_PERMANENT.bat` sudah tercakup di dalamnya).
