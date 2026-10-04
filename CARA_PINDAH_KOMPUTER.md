# 💻 Panduan Instalasi Excel Agent di Komputer Lain

Panduan memindahkan dan menjalankan **Excel Agent** di komputer Windows lain dengan mudah dalam 3 langkah.

---

## 📋 Prasyarat Minimal di Komputer Baru
Sebelum menjalankan setup, pastikan komputer baru sudah memiliki:
1. **Node.js (LTS)** → Download dari [nodejs.org](https://nodejs.org/) (Centang *Add to PATH*).
2. **Python 3** (Disarankan 3.11 atau lebih baru) → Download dari [python.org](https://python.org/) (Centang *Add Python to PATH*).
3. **Microsoft Excel** (Excel 2016 / 2019 / 2021 / Office 365).

---

## 🚀 Langkah Instalasi (Hanya 1x Saja)

1. **Salin Folder `ExcelAgent`** ke komputer baru (misalnya ke `C:\ExcelAgent` atau `D:\ExcelAgent`).
2. **Klik Kanan `SETUP_KOMPUTER_BARU.bat`** -> pilih **Run as Administrator** (atau Double Click langsung).
3. Tunggu proses otomatis selesai:
   - ✅ Memeriksa Node.js & Python
   - ✅ Menginstall library `pywin32` otomatis
   - ✅ Menginstall package `npm install`
   - ✅ Memasang sertifikat HTTPS lokal
   - ✅ Mendaftarkan Add-in ke Registry Office Windows secara permanen
   - ✅ Membuat shortcut **Excel Agent Server** di Desktop

---

## 🎯 Cara Pakai Sehari-Hari

1. **Jalankan Server**:
   Double click shortcut **`Excel Agent Server`** di Desktop (atau file `START_SERVER.bat`).
2. **Buka Excel**:
   Buka file Excel Anda seperti biasa.
3. **Gunakan Asisten AI**:
   - Di tab **HOME (Beranda)** bagian paling kanan, klik tombol **"Buka Asisten AI"**.
   - Task pane asisten AI akan langsung muncul di sebelah kanan dan siap menerima perintah!

---

## ❓ Troubleshooting di Komputer Baru

| Masalah | Solusi |
|---|---|
| **Tombol tidak muncul di Excel** | Jalankan ulang `SETUP_KOMPUTER_BARU.bat` satu kali dengan Excel tertutup. |
| **Status badge menunjukkan Server Offline** | Pastikan `START_SERVER.bat` sedang berjalan di latar belakang (Port 3000). |
| **Error pywin32 / win32com** | Buka Command Prompt dan ketik: `pip install pywin32` |
| **Peringatan Sertifikat SSL** | Jalankan `npx office-addin-dev-certs install --machine` di terminal. |
