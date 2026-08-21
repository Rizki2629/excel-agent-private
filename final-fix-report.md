# Laporan Perbaikan Final

## 14 Agustus 2026

- Enumerasi ROT: kegagalan `GetRunningObjectTable`, `EnumRunning`, dan setiap `Next` sekarang menjadi `TargetWorkbookError` berpesan Indonesia. `execute_envelope` meneruskannya sebagai `targetError`.
- Batas body HTTP: catch `POST` dan `DELETE /api/history` sekarang menghormati `e.statusCode || 400`. Tes HTTP nyata membuktikan body di atas 1 MiB mendapat JSON HTTP 413 dan request kecil berikutnya tetap berhasil untuk kedua metode.
- Penyelesaian multi-step: identitas workbook diperiksa lagi dan perubahan sheet diverifikasi sebelum status akhir ditampilkan. Perubahan identitas memblokir binding tanpa pesan sukses.
- Race hapus riwayat: `historyDeleteInFlight` mencegah POST selama DELETE. DELETE sukses membersihkan riwayat dan snapshot tertunda; DELETE gagal melepas lock dan mengirim snapshot terbaru.
- Tes perilaku task pane: mencakup blok permanen `targetError` pada retry biasa, multi-step, dan tombol manual `executeJS`; kegagalan identitas pada akhir multi-step; serta hasil sukses dan gagal race DELETE.
- Static server: path kini fail-closed ke root allowlist `taskpane` dan `assets`. URL malformed, path absolut, segmen `..`, backslash, traversal encoded, dan kandidat di luar root ditolak sebelum `fs.readFile`. Respons static tidak lagi mengirim wildcard CORS.
- Tombol manual `executeJS`: respons executor sukses diperiksa lagi dengan `assertBindingCurrent` sebelum pesan sukses. Perubahan identitas memblokir binding dan hanya menampilkan kegagalan berbahasa Indonesia.
- Tes regresi baru: HTTP server nyata menolak `/%2e%2e/CLAUDE.md`, `/..%2fCLAUDE.md`, `/..%5cCLAUDE.md`, menangani percent encoding malformed tanpa crash, dan tetap melayani `/taskpane.html`; VM membuktikan sukses executor diikuti kegagalan identitas tidak pernah menampilkan sukses. Tes manual `targetError` tetap dipertahankan.

## Verifikasi

- `npm test`: lulus, 31 tes Node dan 17 tes Python.
- `npm run validate`: lulus, manifest valid.
- `python -m py_compile workbook_target.py exec_python.py`: lulus.
- `node --check server.js`: lulus.
- `node --check taskpane/taskpane.js`: lulus.
- `node --test --test-name-pattern="static server menolak traversal" tests/server.test.js`: lulus, 1 tes focused security probe.
- `git diff --check -- workbook_target.py exec_python.py server.js taskpane/taskpane.js package.json tests/test_workbook_target.py tests/test_exec_python.py tests/server.test.js tests/taskpane.test.js final-fix-report.md`: lulus.
- `git diff --check`: gagal hanya karena trailing whitespace yang sudah ada di `src/frmChat.frm` dan `src/frmSettings.frm`. File VBA tidak diubah.

## Residual

- Tidak ada residual blocker pada dua temuan ini.
- Static roots sengaja terbatas ke `taskpane` dan `assets`; file root repo lain tidak dapat dilayani.
- Verifikasi runtime Office/Excel tidak dijalankan; perubahan tercakup tes HTTP dan VM tanpa memerlukan Excel aktif.
