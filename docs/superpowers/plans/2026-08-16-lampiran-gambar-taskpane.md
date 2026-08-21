# Lampiran Gambar Taskpane Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Menambahkan upload dan paste maksimal lima screenshot agar OpenCode memakai gambar sebagai referensi analisis workbook.

**Architecture:** Taskpane mengirim data URL Base64 bersama prompt. Server memvalidasi dan menulis file temporary, meneruskannya ke `opencode run --file`, lalu menghapus file pada semua jalur selesai.

**Tech Stack:** Office taskpane ES5/IE11, Node.js HTTPS server, OpenCode CLI, `node:test`.

## Global Constraints

- Maksimal 5 gambar per pesan.
- Maksimal 5 MB per gambar.
- Format PNG, JPEG, WEBP.
- Teks prompt opsional.
- File hanya hidup sampai analisis selesai.
- Tidak menambah dependency.

---

### Task 1: Kontrak Lampiran Server

**Files:**
- Modify: `server.js`
- Test: `tests/server.test.js`

**Interfaces:**
- Produces: `prepareOpencodeImages(images): { paths, cleanup }`
- Consumes: request body `/api/opencode.images`

- [ ] Tulis test gagal untuk jumlah, MIME, ukuran, Base64, dan cleanup.
- [ ] Jalankan `node --test tests/server.test.js` dan pastikan test baru gagal.
- [ ] Implementasikan validasi serta file temporary.
- [ ] Tambahkan path ke `runOpencode(..., imagePaths)` sebagai argumen `--file`.
- [ ] Jalankan test sampai lulus.

### Task 2: Composer Multi-Gambar

**Files:**
- Modify: `taskpane/taskpane.html`
- Modify: `taskpane/taskpane.css`
- Modify: `taskpane/taskpane.js`
- Test: `tests/taskpane.test.js`

**Interfaces:**
- Produces: `pendingImages`, `addImageFiles(files)`, `removePendingImage(index)`, `renderImagePreviews()`
- Consumes: `File`, clipboard `items`, dan payload `/api/opencode.images`

- [ ] Tulis test gagal untuk batas gambar, tombol kirim, dan payload.
- [ ] Jalankan `node --test tests/taskpane.test.js` dan pastikan gagal.
- [ ] Tambahkan tombol paperclip, input file, area preview, dan tombol hapus.
- [ ] Tambahkan upload/paste, validasi client, default prompt, dan payload Base64.
- [ ] Bersihkan composer setelah pesan dikirim dan pertahankan thumbnail pesan lokal.
- [ ] Jalankan test sampai lulus.

### Task 3: Versioning dan Verifikasi

**Files:**
- Modify: `taskpane/taskpane.html`
- Modify: `taskpane/taskpane.js`
- Test: `tests/taskpane.test.js`

**Interfaces:**
- Produces: bundle taskpane `v1.5.0`

- [ ] Naikkan versi dan cache-buster.
- [ ] Jalankan `node --check server.js` dan parse `taskpane.js` memakai `new Function`.
- [ ] Jalankan seluruh suite JS dan Python.
- [ ] Restart server lokal.
- [ ] Kirim fixture gambar ke `/api/opencode` dan pastikan respons visual diterima.
