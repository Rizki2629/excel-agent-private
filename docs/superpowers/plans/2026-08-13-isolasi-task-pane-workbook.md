# Isolasi Task Pane per Workbook Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Mengikat setiap instance task pane ke workbook tersimpan yang membukanya, termasuk saat beberapa instance Excel berjalan bersamaan.

**Architecture:** Office.js membuat binding tetap berisi `sessionId`, `workbookName`, dan path lokal saat pane dibuka. Server memvalidasi envelope, menyimpan riwayat per hash path, dan menyerialkan proses COM; executor Python mencari workbook melalui Windows Running Object Table dan hanya menyediakan `wb` yang cocok tepat satu kali.

**Tech Stack:** Office.js JavaScript kompatibel IE11, Node.js HTTPS server dan `node:test`, Python 3 `unittest`, `pywin32`, Microsoft Excel desktop Windows.

## Global Constraints

- Workbook target wajib sudah pernah disimpan dan memiliki path penuh.
- Binding tidak berubah ketika fokus berpindah ke workbook lain.
- Beberapa instance Excel harus didukung.
- Tidak ada fallback ke `ActiveWorkbook`, nama workbook saja, atau instance Excel aktif.
- Path, nama, atau status buka target berubah setelah binding membuat sesi gagal tertutup.
- Pengaturan API tetap global; riwayat maksimal 50 pesan dipisahkan per workbook.
- Jalur VBA UserForm lama tetap di luar cakupan.
- Semua UI, error, komentar baru, dan dokumentasi baru memakai bahasa Indonesia.
- Jangan mengubah atau menghapus perubahan worktree yang tidak terkait.

---

## Struktur File

- `workbook_target.py`: normalisasi target, enumerasi Running Object Table, resolusi workbook tepat satu, dan validasi ulang.
- `exec_python.py`: baca envelope JSON, resolve `xl`/`wb`, eksekusi kode, validasi target sebelum/sesudah.
- `server.js`: validasi origin/request, antrean COM tunggal, endpoint riwayat per workbook, dan spawn executor.
- `taskpane/taskpane.js`: binding tetap Office.js, status terblokir, payload target, prompt berbasis `wb`, dan riwayat scoped.
- `taskpane/taskpane.html` dan `taskpane/taskpane.css`: banner target atau error binding yang ringkas.
- `tests/test_workbook_target.py`: unit test resolver COM memakai objek palsu.
- `tests/test_exec_python.py`: unit test envelope dan lingkungan eksekusi.
- `tests/server.test.js`: unit test validasi target, hash riwayat, origin, dan antrean.
- `tests/manual_multi_instance_excel.py`: probe integrasi terhadap Excel nyata tanpa mengubah data pengguna.
- `README_TASKPANE.md`: prasyarat workbook tersimpan dan perilaku binding.

---

### Task 1: Resolver Workbook COM yang Fail-Closed

**Files:**
- Create: `workbook_target.py`
- Create: `tests/test_workbook_target.py`

**Interfaces:**
- Consumes: target dict dengan string `sessionId`, `workbookName`, `workbookPath`.
- Produces: `normalize_windows_path(path: str) -> str`, `validate_target(target: dict) -> dict`, `iter_excel_applications() -> iterator`, `resolve_workbook(target: dict, applications=None) -> tuple[xl, wb]`, `assert_same_workbook(wb, normalized_path: str) -> None`, dan `TargetWorkbookError`.

- [ ] **Step 1: Tulis unit test target dan pencocokan workbook palsu**

```python
# tests/test_workbook_target.py
import unittest

from workbook_target import (
    TargetWorkbookError,
    assert_same_workbook,
    normalize_windows_path,
    resolve_workbook,
    validate_target,
)


class FakeWorkbook:
    def __init__(self, full_name):
        self.FullName = full_name
        self.Name = full_name.rsplit("\\", 1)[-1]


class FakeApplication:
    def __init__(self, hwnd, paths):
        self.Hwnd = hwnd
        self.Workbooks = [FakeWorkbook(path) for path in paths]


TARGET = {
    "sessionId": "sesi-1",
    "workbookName": "Nilai.xlsm",
    "workbookPath": r"C:\Data\Nilai.xlsm",
}


class WorkbookTargetTests(unittest.TestCase):
    def test_validate_target_normalizes_path(self):
        target = validate_target(TARGET)
        self.assertEqual(target["workbookPath"], r"c:\data\nilai.xlsm")

    def test_validate_target_rejects_relative_or_missing_values(self):
        for bad in ({}, {**TARGET, "sessionId": ""}, {**TARGET, "workbookPath": "Book1"}):
            with self.subTest(bad=bad), self.assertRaises(TargetWorkbookError):
                validate_target(bad)

    def test_resolve_workbook_across_applications(self):
        apps = [FakeApplication(10, [r"C:\Data\Lain.xlsx"]), FakeApplication(20, [r"C:\Data\Nilai.xlsm"])]
        xl, wb = resolve_workbook(TARGET, applications=apps)
        self.assertEqual(xl.Hwnd, 20)
        self.assertEqual(wb.Name, "Nilai.xlsm")

    def test_resolve_rejects_zero_or_duplicate_matches(self):
        with self.assertRaisesRegex(TargetWorkbookError, "tidak ditemukan"):
            resolve_workbook(TARGET, applications=[])
        duplicate = [FakeApplication(10, [TARGET["workbookPath"]]), FakeApplication(20, [TARGET["workbookPath"]])]
        with self.assertRaisesRegex(TargetWorkbookError, "lebih dari satu"):
            resolve_workbook(TARGET, applications=duplicate)

    def test_assert_same_workbook_detects_rename(self):
        wb = FakeWorkbook(r"C:\Data\Baru.xlsm")
        with self.assertRaisesRegex(TargetWorkbookError, "berubah"):
            assert_same_workbook(wb, normalize_windows_path(TARGET["workbookPath"]))


if __name__ == "__main__":
    unittest.main()
```

- [ ] **Step 2: Jalankan test dan pastikan gagal karena modul belum ada**

Run: `python -m unittest tests.test_workbook_target -v`

Expected: FAIL dengan `ModuleNotFoundError: No module named 'workbook_target'`.

- [ ] **Step 3: Implementasikan validasi dan resolver ROT minimal**

```python
# workbook_target.py
import ntpath
from urllib.parse import unquote, urlparse


class TargetWorkbookError(RuntimeError):
    pass


def normalize_windows_path(value):
    if not isinstance(value, str) or not value.strip():
        raise TargetWorkbookError("Path workbook target kosong.")
    value = value.strip()
    if value.lower().startswith("file:"):
        parsed = urlparse(value)
        value = unquote(parsed.path).lstrip("/")
    normalized = ntpath.normcase(ntpath.normpath(value))
    if not ntpath.isabs(normalized):
        raise TargetWorkbookError("Workbook harus disimpan sebelum agent dijalankan.")
    return normalized


def validate_target(target):
    if not isinstance(target, dict):
        raise TargetWorkbookError("Metadata target workbook tidak valid.")
    result = {}
    for key in ("sessionId", "workbookName", "workbookPath"):
        value = target.get(key)
        if not isinstance(value, str) or not value.strip():
            raise TargetWorkbookError("Metadata target workbook tidak lengkap: " + key)
        result[key] = value.strip()
    result["workbookPath"] = normalize_windows_path(result["workbookPath"])
    if ntpath.basename(result["workbookPath"]).lower() != result["workbookName"].lower():
        raise TargetWorkbookError("Nama dan path workbook target tidak cocok.")
    return result


def iter_excel_applications():
    import pythoncom
    import win32com.client

    rot = pythoncom.GetRunningObjectTable()
    bind_ctx = pythoncom.CreateBindCtx(0)
    enum = rot.EnumRunning()
    seen = set()
    while True:
        monikers = enum.Next(1)
        if not monikers:
            break
        moniker = monikers[0]
        try:
            running = rot.GetObject(moniker)
            obj = win32com.client.Dispatch(running.QueryInterface(pythoncom.IID_IDispatch))
            app = obj if getattr(obj, "Name", "") == "Microsoft Excel" else obj.Application
            hwnd = int(app.Hwnd)
            if hwnd not in seen:
                seen.add(hwnd)
                yield app
        except Exception:
            continue


def _workbooks(application):
    collection = application.Workbooks
    try:
        return [collection.Item(index) for index in range(1, collection.Count + 1)]
    except AttributeError:
        return list(collection)


def resolve_workbook(target, applications=None):
    target = validate_target(target)
    matches = []
    for application in applications if applications is not None else iter_excel_applications():
        for workbook in _workbooks(application):
            try:
                if normalize_windows_path(str(workbook.FullName)) == target["workbookPath"]:
                    matches.append((application, workbook))
            except Exception:
                continue
    if not matches:
        raise TargetWorkbookError("Workbook target tidak ditemukan. Buka ulang task pane pada workbook tersebut.")
    if len(matches) != 1:
        raise TargetWorkbookError("Workbook target ditemukan lebih dari satu kali. Tutup salinan duplikat.")
    return matches[0]


def assert_same_workbook(workbook, normalized_path):
    try:
        current = normalize_windows_path(str(workbook.FullName))
    except Exception as exc:
        raise TargetWorkbookError("Workbook target sudah ditutup.") from exc
    if current != normalized_path:
        raise TargetWorkbookError("Identitas workbook target berubah. Buka ulang task pane.")
```

- [ ] **Step 4: Jalankan test resolver**

Run: `python -m unittest tests.test_workbook_target -v`

Expected: 5 tests PASS.

- [ ] **Step 5: Commit unit resolver**

```bash
git add workbook_target.py tests/test_workbook_target.py
git commit -m "feat: resolve bound workbook across Excel instances"
```

---

### Task 2: Executor Python dengan Envelope Target

**Files:**
- Modify: `exec_python.py:1-32`
- Create: `tests/test_exec_python.py`

**Interfaces:**
- Consumes: JSON stdin `{"code": str, "target": Target}`.
- Produces: stdout JSON `{"ok": true, "message": str}` atau `{"ok": false, "error": str, "traceback"?: str}`; kode agent menerima globals `xl` dan `wb`.

- [ ] **Step 1: Tulis test executor tanpa Excel nyata**

```python
# tests/test_exec_python.py
import unittest
from unittest.mock import patch

import exec_python


TARGET = {"sessionId": "s1", "workbookName": "Nilai.xlsm", "workbookPath": r"C:\Data\Nilai.xlsm"}


class ExecPythonTests(unittest.TestCase):
    @patch("exec_python.assert_same_workbook")
    @patch("exec_python.resolve_workbook")
    def test_execute_envelope_exposes_bound_workbook(self, resolve, assert_same):
        xl, wb = object(), object()
        resolve.return_value = (xl, wb)
        result = exec_python.execute_envelope({"code": "result = (xl, wb)", "target": TARGET})
        self.assertTrue(result["ok"])
        resolve.assert_called_once_with(TARGET)
        self.assertEqual(assert_same.call_count, 2)

    def test_execute_envelope_rejects_missing_code_or_target(self):
        for payload in ({}, {"code": "pass"}, {"code": "", "target": TARGET}):
            with self.subTest(payload=payload):
                result = exec_python.execute_envelope(payload)
                self.assertFalse(result["ok"])

    @patch("exec_python.resolve_workbook")
    def test_execute_returns_agent_error_without_fallback(self, resolve):
        resolve.side_effect = RuntimeError("target hilang")
        result = exec_python.execute_envelope({"code": "pass", "target": TARGET})
        self.assertFalse(result["ok"])
        self.assertIn("target hilang", result["error"])


if __name__ == "__main__":
    unittest.main()
```

- [ ] **Step 2: Jalankan test dan pastikan gagal pada interface baru**

Run: `python -m unittest tests.test_exec_python -v`

Expected: FAIL dengan `AttributeError: module 'exec_python' has no attribute 'execute_envelope'`.

- [ ] **Step 3: Ganti executor menjadi parser envelope fail-closed**

```python
# exec_python.py
import json
import sys
import traceback

from workbook_target import assert_same_workbook, resolve_workbook, validate_target


def execute_envelope(payload):
    try:
        if not isinstance(payload, dict):
            raise ValueError("Payload eksekusi harus berupa objek JSON.")
        code = payload.get("code")
        if not isinstance(code, str) or not code.strip():
            raise ValueError("Kode Python kosong.")
        target = validate_target(payload.get("target"))
        xl, wb = resolve_workbook(target)
        assert_same_workbook(wb, target["workbookPath"])
        env = {"xl": xl, "wb": wb, "__builtins__": __builtins__}
        exec(compile(code, "<agent>", "exec"), env)
        assert_same_workbook(wb, target["workbookPath"])
        return {"ok": True, "message": "Berhasil dieksekusi pada " + target["workbookName"]}
    except Exception as exc:
        return {"ok": False, "error": str(exc), "traceback": traceback.format_exc()}


def main():
    try:
        payload = json.loads(sys.stdin.read())
    except Exception as exc:
        print(json.dumps({"ok": False, "error": "Envelope JSON tidak valid: " + str(exc)}))
        return
    print(json.dumps(execute_envelope(payload)))


if __name__ == "__main__":
    main()
```

- [ ] **Step 4: Jalankan seluruh test Python**

Run: `python -m unittest discover -s tests -p "test_*.py" -v`

Expected: 8 tests PASS dan tidak membuka instance Excel baru.

- [ ] **Step 5: Commit executor**

```bash
git add exec_python.py tests/test_exec_python.py
git commit -m "feat: execute Python against bound workbook"
```

---

### Task 3: Server API, Riwayat Scoped, Origin, dan Antrean COM

**Files:**
- Modify: `server.js:1-320`
- Modify: `package.json:6-12`
- Create: `tests/server.test.js`

**Interfaces:**
- Consumes: `POST /api/exec-python` envelope Task 2; endpoint history menerima `workbookPath` dan opsional `chatHistory`.
- Produces: `validateExecutionEnvelope(body)`, `workbookHistoryKey(path)`, `enqueueCom(work)`, `isAllowedOrigin(origin)`, `GET /api/history?workbookPath=<encoded>`, `POST /api/history`, dan `DELETE /api/history`.

- [ ] **Step 1: Tulis test helper server**

```javascript
// tests/server.test.js
const test = require('node:test');
const assert = require('node:assert/strict');
const {
  enqueueCom,
  isAllowedOrigin,
  validateExecutionEnvelope,
  workbookHistoryKey
} = require('../server');

const target = {
  sessionId: 's1',
  workbookName: 'Nilai.xlsm',
  workbookPath: 'C:\\Data\\Nilai.xlsm'
};

test('validasi envelope mewajibkan kode dan target lengkap', () => {
  assert.deepEqual(validateExecutionEnvelope({ code: 'wb.Save()', target }), { code: 'wb.Save()', target });
  assert.throws(() => validateExecutionEnvelope({ code: 'pass' }), /target/i);
  assert.throws(() => validateExecutionEnvelope({ code: '', target }), /kosong/i);
});

test('key riwayat stabil terhadap kapital path', () => {
  assert.equal(workbookHistoryKey('C:\\Data\\Nilai.xlsm'), workbookHistoryKey('c:\\data\\NILAI.xlsm'));
  assert.match(workbookHistoryKey(target.workbookPath), /^[a-f0-9]{64}$/);
});

test('origin hanya mengizinkan add-in lokal dan request tanpa Origin', () => {
  assert.equal(isAllowedOrigin(undefined), true);
  assert.equal(isAllowedOrigin('https://localhost:3000'), true);
  assert.equal(isAllowedOrigin('https://evil.example'), false);
});

test('antrean COM menjalankan pekerjaan berurutan', async () => {
  const events = [];
  const first = enqueueCom(async () => { events.push('mulai-1'); await new Promise(r => setTimeout(r, 20)); events.push('selesai-1'); });
  const second = enqueueCom(async () => { events.push('mulai-2'); events.push('selesai-2'); });
  await Promise.all([first, second]);
  assert.deepEqual(events, ['mulai-1', 'selesai-1', 'mulai-2', 'selesai-2']);
});
```

- [ ] **Step 2: Tambahkan mode import dan jalankan test gagal**

Pindahkan pemanggilan `start()` ke guard berikut agar test tidak membuka port:

```javascript
if (require.main === module) {
    start().catch((err) => {
        console.error('Gagal start server:', err.message);
        process.exit(1);
    });
}
```

Run: `node --test tests/server.test.js`

Expected: FAIL karena helper belum diekspor.

- [ ] **Step 3: Implementasikan helper validasi, hash, origin, dan antrean**

Tambahkan `const crypto = require('crypto');`, ubah default history, dan ekspor helper:

```javascript
const ALLOWED_ORIGIN = `https://localhost:${PORT}`;
let comQueue = Promise.resolve();

function isAllowedOrigin(origin) {
    return !origin || origin === ALLOWED_ORIGIN;
}

function normalizeWorkbookPath(value) {
    if (typeof value !== 'string' || !path.win32.isAbsolute(value.trim())) {
        throw new Error('Path target workbook tidak valid.');
    }
    return path.win32.normalize(value.trim()).toLowerCase();
}

function workbookHistoryKey(workbookPath) {
    return crypto.createHash('sha256').update(normalizeWorkbookPath(workbookPath)).digest('hex');
}

function validateExecutionEnvelope(body) {
    if (!body || typeof body.code !== 'string' || !body.code.trim()) throw new Error('Kode Python kosong.');
    if (!body.target || typeof body.target !== 'object') throw new Error('Metadata target workbook tidak tersedia.');
    ['sessionId', 'workbookName', 'workbookPath'].forEach((key) => {
        if (typeof body.target[key] !== 'string' || !body.target[key].trim()) throw new Error('Target workbook tidak lengkap: ' + key);
    });
    normalizeWorkbookPath(body.target.workbookPath);
    return { code: body.code, target: body.target };
}

function enqueueCom(work) {
    const queued = comQueue.then(work, work);
    comQueue = queued.catch(() => undefined);
    return queued;
}

module.exports = { enqueueCom, isAllowedOrigin, normalizeWorkbookPath, validateExecutionEnvelope, workbookHistoryKey };
```

Ubah `DEFAULT_CONFIG.chatHistory` menjadi `chatHistoryByWorkbook: {}`. Saat membaca config lama, jangan memasukkan array global ke pane baru; pertahankan field lama tanpa memakainya agar tidak salah mengaitkan history ke workbook tertentu.

- [ ] **Step 4: Implementasikan endpoint history per workbook**

Gunakan key hash, maksimum 50 pesan, dan allowlist role:

```javascript
function cleanHistory(value) {
    if (!Array.isArray(value)) throw new Error('Riwayat chat harus berupa array.');
    return value.filter(item => item && (item.role === 'user' || item.role === 'assistant') && typeof item.content === 'string').slice(-50);
}

function readWorkbookHistory(workbookPath) {
    const config = readConfig();
    return (config.chatHistoryByWorkbook || {})[workbookHistoryKey(workbookPath)] || [];
}

function writeWorkbookHistory(workbookPath, history) {
    const config = readConfig();
    const histories = Object.assign({}, config.chatHistoryByWorkbook || {});
    histories[workbookHistoryKey(workbookPath)] = cleanHistory(history);
    writeConfig({ chatHistoryByWorkbook: histories });
    return histories[workbookHistoryKey(workbookPath)];
}

function deleteWorkbookHistory(workbookPath) {
    const config = readConfig();
    const histories = Object.assign({}, config.chatHistoryByWorkbook || {});
    delete histories[workbookHistoryKey(workbookPath)];
    writeConfig({ chatHistoryByWorkbook: histories });
}
```

Tambahkan routing `GET /api/history?workbookPath=...`, `POST /api/history` dengan `{workbookPath, chatHistory}`, dan `DELETE /api/history` dengan `{workbookPath}`. Jangan mengembalikan map semua workbook melalui `/api/settings`.

- [ ] **Step 5: Serialisasikan endpoint executor dan kirim JSON utuh ke stdin**

Ganti isi route `/api/exec-python` agar `validateExecutionEnvelope` dijalankan sebelum spawn dan seluruh pekerjaan dibungkus `enqueueCom`. `child.stdin.write(JSON.stringify(envelope))`; respons dimulai hanya setelah child selesai. Timeout tetap 60 detik. Error validasi menghasilkan HTTP 400; error spawn/timeout menghasilkan HTTP 500.

- [ ] **Step 6: Batasi origin API lokal**

Pada awal callback server, tolak request API dengan origin selain `https://localhost:3000`:

```javascript
const origin = req.headers.origin;
if (req.url.startsWith('/api/') && !isAllowedOrigin(origin)) {
    res.writeHead(403, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify({ ok: false, error: 'Origin tidak diizinkan.' }));
    return;
}
```

Ganti header CORS `*` pada API menjadi `ALLOWED_ORIGIN`. Request same-origin tanpa header `Origin` tetap diterima untuk probe lokal dan test.

- [ ] **Step 7: Tambahkan script test dan jalankan**

```json
"scripts": {
  "test": "node --test tests/server.test.js && python -m unittest discover -s tests -p \"test_*.py\" -v"
}
```

Run: `npm test`

Expected: semua test Node dan 8 test Python PASS.

- [ ] **Step 8: Commit server**

```bash
git add server.js package.json tests/server.test.js
git commit -m "feat: serialize scoped workbook executions"
```

---

### Task 4: Binding Tetap dan Riwayat per Workbook di Task Pane

**Files:**
- Modify: `taskpane/taskpane.js:1-1062`
- Modify: `taskpane/taskpane.html:38-100,137`
- Modify: `taskpane/taskpane.css:76-105,380-400`

**Interfaces:**
- Consumes: `Office.context.document.getFilePropertiesAsync`, API history Task 3, dan envelope executor Task 2.
- Produces: global `workbookBinding`; `createSessionId()`, `normalizeFileUrl(url)`, `bindTaskPane(callback)`, `assertBindingCurrent(callback)`, `setBindingBlocked(message)`, dan semua request eksekusi menyertakan `target: workbookBinding`.

- [ ] **Step 1: Tambahkan status binding ke HTML**

Di bawah header dan sebelum `welcomeScreen`, tambahkan:

```html
<div id="workbookStatus" class="workbook-status hidden" role="status" aria-live="polite"></div>
```

Naikkan cache query script menjadi `taskpane.js?v=28`.

- [ ] **Step 2: Tambahkan style status ringkas tanpa menggeser kontrol**

```css
.workbook-status {
    margin: 8px 12px 0;
    padding: 8px 10px;
    border: 1px solid #3b4452;
    border-radius: 6px;
    color: #d7dde8;
    background: #20252d;
    font-size: 12px;
    line-height: 1.4;
    overflow-wrap: anywhere;
}

.workbook-status.error {
    border-color: #8f3b3b;
    color: #ffd8d8;
    background: #351f22;
}

.workbook-status.hidden { display: none; }
```

- [ ] **Step 3: Implementasikan binding satu kali sebelum settings/history**

Tambahkan state dan helper kompatibel IE11:

```javascript
var workbookBinding = null;
var bindingBlocked = false;

function createSessionId() {
    return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function(c) {
        var r = Math.random() * 16 | 0;
        return (c === 'x' ? r : (r & 3 | 8)).toString(16);
    });
}

function normalizeFileUrl(url) {
    if (typeof url !== 'string' || url.toLowerCase().indexOf('file:') !== 0) return null;
    var path = decodeURIComponent(url.replace(/^file:\/\/+/, ''));
    path = path.replace(/\//g, '\\');
    return /^[A-Za-z]:\\/.test(path) ? path : null;
}

function setBindingBlocked(message) {
    bindingBlocked = true;
    workbookBinding = null;
    var status = document.getElementById('workbookStatus');
    status.className = 'workbook-status error';
    status.textContent = message;
    document.getElementById('userInput').disabled = true;
    document.getElementById('sendBtn').disabled = true;
}

function bindTaskPane(done) {
    Office.context.document.getFilePropertiesAsync(function(result) {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
            setBindingBlocked('Workbook tidak dapat diidentifikasi. Simpan workbook lalu buka ulang task pane.');
            done(false);
            return;
        }
        var workbookPath = normalizeFileUrl(result.value.url);
        if (!workbookPath) {
            setBindingBlocked('Simpan workbook ke drive lokal lalu buka ulang task pane.');
            done(false);
            return;
        }
        var parts = workbookPath.split('\\');
        workbookBinding = { sessionId: createSessionId(), workbookName: parts[parts.length - 1], workbookPath: workbookPath };
        var status = document.getElementById('workbookStatus');
        status.className = 'workbook-status';
        status.textContent = 'Target: ' + workbookBinding.workbookName;
        done(true);
    });
}
```

Ubah `initApp()` agar memanggil `bindTaskPane`, lalu `loadSettings`, `loadHistory`, dan `restoreHistory` hanya setelah binding sukses. Settings tetap dapat dibuka ketika binding gagal.

- [ ] **Step 4: Pisahkan load/save/delete history dari settings**

Hapus pembacaan `cfg.chatHistory` dari `loadSettings`. Gunakan endpoint baru:

```javascript
function loadHistory(done) {
    var url = SERVER + '/api/history?workbookPath=' + encodeURIComponent(workbookBinding.workbookPath);
    var xhr = new XMLHttpRequest();
    xhr.open('GET', url, true);
    xhr.onreadystatechange = function() {
        if (xhr.readyState !== 4) return;
        if (xhr.status === 200) {
            try { chatHistory = JSON.parse(xhr.responseText).chatHistory || []; } catch (e) { chatHistory = []; }
        }
        if (done) done();
    };
    xhr.onerror = function() { if (done) done(); };
    xhr.send();
}

function saveHistory() {
    if (!workbookBinding) return;
    xhrPost(SERVER + '/api/history', {
        workbookPath: workbookBinding.workbookPath,
        chatHistory: chatHistory.slice(-50)
    }, function() {});
}
```

Ubah aksi hapus chat agar memakai `DELETE /api/history` dengan path binding, lalu kosongkan UI hanya setelah status sukses.

- [ ] **Step 5: Validasi binding sebelum membaca konteks dan mengeksekusi**

```javascript
function assertBindingCurrent(callback) {
    if (!workbookBinding || bindingBlocked) {
        callback({ error: 'Sesi workbook tidak valid. Buka ulang task pane.' });
        return;
    }
    Office.context.document.getFilePropertiesAsync(function(result) {
        var current = result.status === Office.AsyncResultStatus.Succeeded ? normalizeFileUrl(result.value.url) : null;
        if (!current || current.toLowerCase() !== workbookBinding.workbookPath.toLowerCase()) {
            setBindingBlocked('Identitas workbook berubah. Buka ulang task pane.');
            callback({ error: 'Identitas workbook target berubah.' });
            return;
        }
        callback(null);
    });
}
```

Panggil helper ini pada awal `sendMessage`, sebelum setiap retry, dan sebelum `verifySheetChanged`. Error binding tidak boleh masuk retry AI karena retry tidak dapat memperbaiki target.

- [ ] **Step 6: Ganti prompt dan payload menjadi `wb` eksplisit**

Ganti seluruh instruksi/contoh `xl.ActiveWorkbook` pada prompt dengan `wb`. Tambahkan aturan:

```text
Variabel `wb` adalah satu-satunya workbook target untuk sesi task pane ini.
JANGAN gunakan `xl.ActiveWorkbook`, `xl.Workbooks(...)`, atau membuka workbook lain untuk memilih target.
Gunakan `ws = wb.ActiveSheet` atau `ws = wb.Worksheets("Nama Sheet")`.
```

Ubah executor client:

```javascript
function runPythonCode(code, callback) {
    if (!workbookBinding) {
        callback({ ok: false, targetError: true, error: 'Sesi workbook tidak valid.' });
        return;
    }
    xhrPost(SERVER + '/api/exec-python', { code: code, target: workbookBinding }, function(status, data) {
        if (status >= 200 && status < 300 && data && data.ok) callback({ ok: true, message: data.message || 'OK' });
        else callback({ ok: false, targetError: status === 400 || status === 409, error: data && data.error ? data.error : 'HTTP ' + status });
    });
}
```

Jika `targetError` benar, tampilkan error dan hentikan `isProcessing` tanpa meminta model memperbaiki kode.

- [ ] **Step 7: Periksa tidak ada pemilihan target COM lama dalam task pane modern**

Run: `rg -n "xl\.ActiveWorkbook|Dispatch\(\"Excel.Application\"\)|\{ code: code \}" taskpane/taskpane.js exec_python.py`

Expected: tidak ada hasil. Penyebutan larangan string `xl.ActiveWorkbook` dalam prompt boleh tersisa satu kali; verifikasi manual bahwa itu aturan larangan, bukan contoh operasi.

- [ ] **Step 8: Jalankan test dan validasi manifest**

Run: `npm test`

Expected: semua test PASS.

Run: `npm run validate`

Expected: manifest valid tanpa error.

- [ ] **Step 9: Commit task pane**

```bash
git add taskpane/taskpane.js taskpane/taskpane.html taskpane/taskpane.css
git commit -m "feat: bind task pane to opening workbook"
```

---

### Task 5: Probe Excel Nyata, Dokumentasi, dan Verifikasi Akhir

**Files:**
- Create: `tests/manual_multi_instance_excel.py`
- Modify: `README_TASKPANE.md:13-20,61-71,114-141`

**Interfaces:**
- Consumes: `iter_excel_applications`, `resolve_workbook`, workbook tersimpan yang sedang terbuka.
- Produces: laporan JSON read-only berisi HWND, workbook path, dan hasil resolusi; tidak menulis workbook.

- [ ] **Step 1: Buat probe read-only multi-instance**

```python
# tests/manual_multi_instance_excel.py
import argparse
import json
import ntpath

from workbook_target import iter_excel_applications, resolve_workbook


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("workbook_path")
    args = parser.parse_args()
    apps = list(iter_excel_applications())
    inventory = []
    for app in apps:
        inventory.append({
            "hwnd": int(app.Hwnd),
            "workbooks": [str(app.Workbooks.Item(i).FullName) for i in range(1, app.Workbooks.Count + 1)],
        })
    target = {
        "sessionId": "probe-manual",
        "workbookName": ntpath.basename(args.workbook_path),
        "workbookPath": args.workbook_path,
    }
    app, workbook = resolve_workbook(target, applications=apps)
    print(json.dumps({"inventory": inventory, "resolvedHwnd": int(app.Hwnd), "resolvedPath": str(workbook.FullName)}, indent=2))


if __name__ == "__main__":
    main()
```

- [ ] **Step 2: Dokumentasikan binding dan error workbook belum disimpan**

Tambahkan ke `README_TASKPANE.md`:

```markdown
### Target Workbook

Task pane terikat ke workbook yang membukanya. Mengaktifkan workbook atau jendela Excel lain tidak mengubah target agent.

- Simpan workbook ke drive lokal sebelum membuka task pane.
- Setelah workbook diganti nama atau dipindah, buka ulang task pane.
- Setiap workbook memiliki riwayat percakapan sendiri.
- Beberapa instance Excel didukung selama setiap target memiliki path penuh yang unik.
```

Perbarui bagian usage agar menyebut Python COM, bukan VBA, sesuai implementasi aktif.

- [ ] **Step 3: Jalankan verifikasi otomatis lengkap**

Run: `npm test`

Expected: seluruh test Node dan Python PASS.

Run: `npm run validate`

Expected: manifest valid.

Run: `git diff --check`

Expected: tidak ada whitespace error.

- [ ] **Step 4: Jalankan probe pada dua instance Excel**

Prasyarat: buka dua workbook tersimpan berbeda pada dua proses Excel. Jalankan untuk masing-masing path:

```powershell
python tests/manual_multi_instance_excel.py "C:\Data\WorkbookA.xlsx"
python tests/manual_multi_instance_excel.py "C:\Data\WorkbookB.xlsx"
```

Expected: `resolvedPath` sama persis dengan argumen ternormalisasi dan `resolvedHwnd` menunjukkan instance pemilik masing-masing workbook.

- [ ] **Step 5: Uji task pane end-to-end**

1. Jalankan `npm start`.
2. Buka Workbook A dan Workbook B tersimpan, masing-masing dengan task pane.
3. Dari pane A, minta menulis nilai unik ke sel kosong yang disetujui untuk pengujian, lalu aktifkan B sebelum eksekusi selesai.
4. Pastikan hanya A berubah dan history A tidak muncul di pane B.
5. Tutup A, lalu kirim lagi dari pane A. Pastikan error target muncul dan B tidak berubah.
6. Buka workbook baru belum disimpan. Pastikan input agent dinonaktifkan dan instruksi simpan muncul.

Expected: seluruh enam kondisi sesuai, tanpa proses Excel baru yang dibuat oleh executor.

- [ ] **Step 6: Commit dokumentasi dan probe**

```bash
git add tests/manual_multi_instance_excel.py README_TASKPANE.md
git commit -m "docs: verify isolated workbook sessions"
```

---

## Pemeriksaan Akhir

- Resolver hanya menerima path absolut dan tepat satu kecocokan `Workbook.FullName`.
- Executor menyediakan `wb`, memvalidasi sebelum/sesudah, dan tidak memakai `Dispatch` atau `ActiveWorkbook`.
- Server menolak envelope tanpa target, menjalankan COM berurutan, dan membatasi origin API lokal.
- Task pane mengunci binding saat init, memblokir workbook belum disimpan, dan menghentikan retry pada error target.
- Riwayat tersimpan per hash path; settings API tetap global.
- Uji otomatis, manifest, probe dua instance, dan skenario end-to-end lulus.
