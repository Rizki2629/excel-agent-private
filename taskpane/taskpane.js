// Excel Agent - Task Pane JavaScript (IE11 Compatible)
// No arrow functions, no const/let, no fetch, no template literals, no async/await
try { console.log('[ExcelAgent] taskpane v1.5.1 loaded'); } catch (e) { }

var settings = {
    hasApiKey: false,
    apiEndpoint: '',
    model: '',
    autoRun: true,
    useOpencode: false,
    opencodeModel: '9router/qd/ultimate'
};

var chatHistory = [];
var SERVER = 'https://localhost:3000';
var MAX_RETRY = 2;
var MAX_REPORT_CYCLES = 2; // maksimum siklus "eksekusi -> lapor output ke AI" per pesan
var reportCycles = 0;
var MAX_NOCODE_RETRY = 2; // berapa kali paksa AI menulis kode bila ia hanya menjawab teks
var noCodeRetries = 0;
var expectingAnswer = false; // true saat AI diminta memberi jawaban akhir tanpa kode
var AGENT_TURN_TIMEOUT_MS = 210000;
// Turn streaming: token pertama harus tiba jauh lebih cepat dari batas ini;
// batas panjang hanya menjaga keseluruhan jawaban yang sangat panjang.
var STREAM_TIMEOUT_MS = 600000;
var AGENT_WATCHDOG_MS = STREAM_TIMEOUT_MS + 30000;
// Kolaborasi OpenCode: harus lebih panjang dari timeout server (420 detik).
var OPENCODE_CLIENT_TIMEOUT_MS = 440000;
var opencodeSessionId = null; // session OpenCode aktif untuk kolaborasi multi-giliran
var isProcessing = false; // Deduplication flag
var sheetSnapshotBefore = null; // Snapshot sheet state sebelum eksekusi
var workbookBinding = null;
var bindingBlocked = false;
var historySaveInFlight = false;
var historyDeleteInFlight = false;
var pendingHistorySnapshot = null;
var historySaveGeneration = 0;
var historyRevision = 0;
var MAX_PENDING_IMAGES = 5;
var MAX_IMAGE_BYTES = 5 * 1024 * 1024;
var ALLOWED_IMAGE_TYPES = { 'image/png': true, 'image/jpeg': true, 'image/webp': true };
var pendingImages = [];
var activeTurnImages = [];

// ---- Pembatalan proses (tombol Hentikan / tombol Esc) ----
// cancelRequested: menandai giliran saat ini sudah dibatalkan user, sehingga
// callback yang terlanjur berjalan tidak melanjutkan ke langkah berikutnya.
var cancelRequested = false;
// Daftar XHR yang sedang berjalan agar bisa di-abort seketika.
var activeXhrs = [];

function trackXhr(xhr) {
    if (xhr) activeXhrs.push(xhr);
    return xhr;
}

function untrackXhr(xhr) {
    for (var i = activeXhrs.length - 1; i >= 0; i--) {
        if (activeXhrs[i] === xhr) activeXhrs.splice(i, 1);
    }
}

function abortAllXhrs() {
    var list = activeXhrs.slice();
    activeXhrs = [];
    for (var i = 0; i < list.length; i++) {
        try { list[i].abort(); } catch (e) { }
    }
    return list.length;
}

// Hentikan giliran agent yang sedang berjalan: putuskan koneksi AI,
// minta server membunuh proses Python/VBA, lalu buka kunci UI.
function stopProcessing(reason) {
    if (!isProcessing) return false;
    cancelRequested = true;
    abortAllXhrs();

    // Minta server menghentikan proses eksekusi yang masih hidup.
    try {
        var body = {};
        if (workbookBinding && workbookBinding.workbookPath) {
            body.workbookPath = workbookBinding.workbookPath;
        }
        var xhr = new XMLHttpRequest();
        xhr.open('POST', SERVER + '/api/cancel', true);
        xhr.setRequestHeader('Content-Type', 'application/json');
        xhr.send(JSON.stringify(body));
    } catch (e) { /* server mungkin sudah tidak merespons */ }

    isProcessing = false;
    activeTurnImages = [];
    try {
        var loading = document.getElementById('loadingIndicator');
        if (loading) loading.className = 'loading hidden';
        var progress = document.getElementById('progressText');
        if (progress) progress.className = 'progress-text hidden';
    } catch (e) { }

    addMessage('system', reason || 'Proses dihentikan oleh pengguna.');
    saveHistory();
    updateComposerState();
    return true;
}

// Initialization
var isInitialized = false;

function initApp() {
    if (isInitialized) return;
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', initApp);
        return;
    }
    
    // Pastikan elemen DOM sudah ada
    if (!document.getElementById('userInput')) return;

    isInitialized = true;
    
    setupEventListeners();
    startHealthPolling();
    bindTaskPane(function (bound) {
        loadSettings(function () {
            if (bound) {
                loadHistory(function () { restoreHistory(); });
            }
        });
        // Profil identitas + memori: agar model mengenali pemilik sejak awal
        loadProfilMemori(function () {
            fillProfilForm();
            renderMemoriList();
        });
    });
}

// Gunakan Office.onReady sebagai metode utama, namun sediakan fallback window.onload
Office.onReady(function () {
    initApp();
});

window.onload = function () {
    initApp();
};

// Muat setting dari server (persisten di disk). localStorage dipakai
// sebagai cadangan bila server sedang mati.
function loadSettings(done) {
    loadSettingsLocal();
    var migrationApiKey = settings.migrationApiKey || '';
    var xhr = new XMLHttpRequest();
    xhr.open('GET', SERVER + '/api/settings', true);
    xhr.onreadystatechange = function () {
        if (xhr.readyState !== 4) return;

        var loaded = false;
        if (xhr.status === 200) {
            try {
                var cfg = JSON.parse(xhr.responseText);
                for (var key in cfg) {
                    if (cfg.hasOwnProperty(key) && key !== 'chatHistory') {
                        settings[key] = cfg[key];
                    }
                }
                loaded = true;
                if (migrationApiKey && !cfg.hasApiKey) {
                    saveSettings(done, migrationApiKey);
                    return;
                }
                delete settings.migrationApiKey;
            } catch (e) { }
        }
        if (!loaded) {
            // Migrasi setting lama dari cache Office ke penyimpanan disk.
            if (settings.migrationApiKey || settings.apiEndpoint || settings.model) {
                saveSettings(null, settings.migrationApiKey);
            }
        }
        if (done) done();
    };
    xhr.onerror = function () {
        if (done) done();
    };
    xhr.send();
}

function loadSettingsLocal() {
    try {
        var saved = localStorage.getItem('excelAgentSettings');
        if (!saved) return;
        var parsed = JSON.parse(saved);
        for (var key in parsed) {
            if (parsed.hasOwnProperty(key) && key !== 'apiKey') settings[key] = parsed[key];
        }
        if (typeof parsed.apiKey === 'string' && parsed.apiKey) settings.migrationApiKey = parsed.apiKey;
        delete parsed.apiKey;
        localStorage.setItem('excelAgentSettings', JSON.stringify(parsed));
    } catch (e) { }
}

// API key hanya disimpan server. Cache lokal tidak boleh memuat secret.
function saveSettings(done, apiKey) {
    try {
        localStorage.setItem('excelAgentSettings', JSON.stringify({
            apiEndpoint: settings.apiEndpoint,
            model: settings.model,
            autoRun: settings.autoRun,
            useOpencode: settings.useOpencode,
            opencodeModel: settings.opencodeModel,
            modelList: settings.modelList
        }));
    } catch (e) { }

    var payload = {
        apiEndpoint: settings.apiEndpoint,
        model: settings.model,
        autoRun: settings.autoRun,
        useOpencode: settings.useOpencode,
        opencodeModel: settings.opencodeModel
    };
    if (Array.isArray(settings.modelList)) payload.modelList = settings.modelList;
    if (typeof apiKey === 'string' && apiKey) payload.apiKey = apiKey;

    var xhr = new XMLHttpRequest();
    xhr.open('POST', SERVER + '/api/settings', true);
    xhr.setRequestHeader('Content-Type', 'application/json');
    xhr.onreadystatechange = function () {
        if (xhr.readyState !== 4) return;
        if (xhr.status === 200) {
            try { settings.hasApiKey = JSON.parse(xhr.responseText).config.hasApiKey === true; } catch (e) { }
            delete settings.migrationApiKey;
        }
        if (done) done(xhr.status === 200);
    };
    xhr.onerror = function () { if (done) done(false); };
    xhr.send(JSON.stringify(payload));
}

function createSessionId() {
    return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function(c) {
        var r = Math.random() * 16 | 0;
        return (c === 'x' ? r : (r & 3 | 8)).toString(16);
    });
}

// ---- Profil identitas & memori jangka panjang ----
// Dimuat sekali saat start, disuntik ke system prompt setiap chat
// agar model mengenali pemilik dan belajar dari sesi sebelumnya.
var userProfil = {};
var userMemori = [];

function loadProfilMemori(done) {
    var finished = 0;
    function oneDone() {
        finished += 1;
        if (finished >= 2 && done) done();
    }
    var xhr1 = new XMLHttpRequest();
    xhr1.open('GET', SERVER + '/api/profil', true);
    xhr1.onreadystatechange = function () {
        if (xhr1.readyState !== 4) return;
        if (xhr1.status === 200) {
            try { userProfil = JSON.parse(xhr1.responseText).profil || {}; } catch (e) { }
        }
        oneDone();
    };
    xhr1.onerror = oneDone;
    xhr1.send();
    var xhr2 = new XMLHttpRequest();
    xhr2.open('GET', SERVER + '/api/memori', true);
    xhr2.onreadystatechange = function () {
        if (xhr2.readyState !== 4) return;
        if (xhr2.status === 200) {
            try { userMemori = JSON.parse(xhr2.responseText).memori || []; } catch (e) { }
        }
        oneDone();
    };
    xhr2.onerror = oneDone;
    xhr2.send();
}

function buildProfilPrompt() {
    var p = userProfil || {};
    var lines = ['### PROFIL PENGGUNA ###'];
    lines.push('Kamu mengenali pengguna ini. Sapa dan layani sesuai identitasnya.');
    if (p.nama) lines.push('Nama: ' + p.nama);
    if (p.peran) lines.push('Peran: ' + p.peran);
    if (p.zonaWaktu) lines.push('Zona waktu: ' + p.zonaWaktu);
    if (p.bahasa) lines.push('Bahasa: ' + p.bahasa);
    if (p.gayaJawab) lines.push('Gaya jawab: ' + p.gayaJawab);
    lines.push('');
    return lines.join('\n');
}

function buildMemoriPrompt() {
    var lines = ['### MEMORI JANGKA PANJANG ###'];
    lines.push('Ini pelajaran dari sesi-sesi sebelumnya dengan pengguna ini. Patuhi semuanya.');
    if (!userMemori.length) {
        lines.push('(Belum ada memori tersimpan.)');
    } else {
        userMemori.forEach(function (m) {
            lines.push('- [' + m.kategori + '] ' + m.isi);
        });
    }
    lines.push('');
    return lines.join('\n');
}

// Protokol agar agent bisa MENULIS memori sendiri lewat blok ```memory.
var MEMORY_PROTOCOL =
    '### PROTOKOL MEMORI (CARA BELAJAR) ###\n' +
    'Kamu BISA dan HARUS belajar dari pengguna. Jika pengguna berkata "ingat ini", "catat",\n' +
    '"jangan lupa", atau mengoreksimu ("bukan begitu, yang benar...", "salah, harusnya..."),\n' +
    'simpan pelajarannya ke memori permanen dengan blok:\n' +
    '```memory\n' +
    'kategori: preferensi\n' +
    'isi: ...\n' +
    '```\n' +
    'Kategori: preferensi (kebiasaan user), koreksi (kesalahanmu yang diperbaiki user),\n' +
    'pola (cara kerja yang terbukti berhasil, bisa dipakai ulang), fakta (info tentang user/pekerjaannya).\n' +
    'Tulis SATU pelajaran per blok, singkat dan konkret (maks 2 kalimat).\n' +
    'Contoh: user bilang "ingat, daftar hadir selalu kolom tanggal format DD-MM-YYYY"\n' +
    '-> ```memory\\nkategori: preferensi\\nisi: Format tanggal di daftar hadir selalu DD-MM-YYYY.\\n```\n' +
    'DILARANG menyimpan API key, password, token, atau data rahasia apa pun ke memori.\n' +
    'JANGAN umumkan bahwa kamu menyimpan memori dengan kalimat panjang; cukup kerjakan.\n\n';

function extractMemoryBlocks(content) {
    var blocks = [];
    var regex = /```memory[ \t]*\r?\n([\s\S]*?)```/gi;
    var match;
    while ((match = regex.exec(content)) !== null) {
        var body = match[1];
        var kategori = 'fakta';
        var isi = '';
        var km = body.match(/kategori\s*:\s*(preferensi|koreksi|pola|fakta)/i);
        if (km) kategori = km[1].toLowerCase();
        var im = body.match(/isi\s*:\s*([\s\S]*)/i);
        if (im) isi = im[1].replace(/^\s+|\s+$/g, '');
        if (isi) blocks.push({ kategori: kategori, isi: isi });
    }
    return blocks;
}

function saveMemoryBlock(block, done) {
    var xhr = new XMLHttpRequest();
    xhr.open('POST', SERVER + '/api/memori', true);
    xhr.setRequestHeader('Content-Type', 'application/json');
    xhr.onreadystatechange = function () {
        if (xhr.readyState !== 4) return;
        if (xhr.status === 200) {
            try { userMemori = JSON.parse(xhr.responseText).memori || userMemori; } catch (e) { }
        }
        if (done) done(xhr.status === 200);
    };
    xhr.onerror = function () { if (done) done(false); };
    xhr.send(JSON.stringify(block));
}

function saveProfil(done) {
    var payload = {
        nama: document.getElementById('profilNama').value,
        peran: document.getElementById('profilPeran').value,
        zonaWaktu: document.getElementById('profilZona').value,
        bahasa: document.getElementById('profilBahasa').value,
        gayaJawab: document.getElementById('profilGaya').value
    };
    var xhr = new XMLHttpRequest();
    xhr.open('POST', SERVER + '/api/profil', true);
    xhr.setRequestHeader('Content-Type', 'application/json');
    xhr.onreadystatechange = function () {
        if (xhr.readyState !== 4) return;
        if (xhr.status === 200) {
            try { userProfil = JSON.parse(xhr.responseText).profil || {}; } catch (e) { }
        }
        if (done) done(xhr.status === 200);
    };
    xhr.onerror = function () { if (done) done(false); };
    xhr.send(JSON.stringify(payload));
}

function deleteMemoriEntry(id, done) {
    var xhr = new XMLHttpRequest();
    xhr.open('DELETE', SERVER + '/api/memori', true);
    xhr.setRequestHeader('Content-Type', 'application/json');
    xhr.onreadystatechange = function () {
        if (xhr.readyState !== 4) return;
        if (xhr.status === 200) {
            try { userMemori = JSON.parse(xhr.responseText).memori || []; } catch (e) { }
        }
        if (done) done(xhr.status === 200);
    };
    xhr.onerror = function () { if (done) done(false); };
    xhr.send(JSON.stringify({ id: id }));
}

// Isi form profil di modal Pengaturan dari data server.
function fillProfilForm() {
    var p = userProfil || {};
    var map = {
        profilNama: p.nama, profilPeran: p.peran, profilZona: p.zonaWaktu,
        profilBahasa: p.bahasa, profilGaya: p.gayaJawab
    };
    for (var id in map) {
        var el = document.getElementById(id);
        if (el && map[id]) el.value = map[id];
    }
}

// Tampilkan daftar memori di modal Pengaturan, bisa dihapus satu per satu.
function renderMemoriList() {
    var box = document.getElementById('memoriList');
    if (!box) return;
    box.innerHTML = '';
    if (!userMemori.length) {
        box.innerHTML = '<div class="memori-empty">Belum ada memori. Agent mencatat otomatis saat kamu bilang "ingat ini" atau mengoreksinya.</div>';
        return;
    }
    userMemori.forEach(function (m) {
        var row = document.createElement('div');
        row.className = 'memori-row';
        var tag = document.createElement('span');
        tag.className = 'memori-tag memori-' + m.kategori;
        tag.textContent = m.kategori;
        var txt = document.createElement('span');
        txt.className = 'memori-text';
        txt.textContent = m.isi;
        var del = document.createElement('button');
        del.className = 'inline-btn memori-del';
        del.type = 'button';
        del.textContent = 'Hapus';
        del.onclick = (function (id) {
            return function () {
                deleteMemoriEntry(id, function () { renderMemoriList(); });
            };
        })(m.id);
        row.appendChild(tag);
        row.appendChild(txt);
        row.appendChild(del);
        box.appendChild(row);
    });
}

function normalizeFileUrl(url) {
    if (typeof url !== 'string') return null;
    url = url.replace(/^\s+|\s+$/g, '');
    if (/^[A-Za-z]:[\\\/]/.test(url)) return url.replace(/\//g, '\\');
    if (/^\\\\[^\\]+\\[^\\]+/.test(url)) return url;
    if (/^https:\/\//i.test(url)) {
        var cloudMatch = url.match(/^https:\/\/([^\/]+)(\/.*)$/i);
        if (!cloudMatch) return null;
        var cloudHost = cloudMatch[1].toLowerCase();
        if (cloudHost === 'd.docs.live.net' || /\.sharepoint\.com$/.test(cloudHost)) {
            var cloudPath;
            try {
                cloudPath = decodeURIComponent(cloudMatch[2]);
            } catch (e) {
                return null;
            }
            if (cloudHost === 'd.docs.live.net') {
                var cloudParts = cloudPath.split('/');
                if (cloudParts.length > 1) cloudParts[1] = cloudParts[1].toLowerCase();
                cloudPath = cloudParts.join('/');
            }
            return 'https://' + cloudHost + cloudPath;
        }
        return null;
    }
    if (url.toLowerCase().indexOf('file:') !== 0) return null;
    var path;
    try {
        path = decodeURIComponent(url.substring(5));
    } catch (e) {
        return null;
    }
    if (path.indexOf('//') === 0) {
        path = path.substring(2);
        if (path.charAt(0) !== '/') {
            path = '\\\\' + path;
        }
    }
    if (/^\/[A-Za-z]:\//.test(path)) path = path.substring(1);
    path = path.replace(/\//g, '\\');
    return (/^[A-Za-z]:\\/.test(path) || /^\\\\[^\\]+\\[^\\]+/.test(path)) ? path : null;
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

// Dapatkan atau buat ID unik yang TERTANAM di dalam file Excel (Custom Property).
// Karena tersimpan di dalam file, saat file di-copy/paste atau di-Save As dengan
// nama lain, ID ini ikut terbawa sehingga riwayat chat otomatis muncul di file baru.
var EXCEL_AGENT_ID_PROP = 'ExcelAgentId';

function getOrCreateWorkbookId(done) {
    if (typeof Excel === 'undefined' || !Excel.run) {
        if (done) done(null);
        return;
    }
    Excel.run(function (context) {
        var custom = context.workbook.properties.custom;
        var existing = custom.getItemOrNullObject(EXCEL_AGENT_ID_PROP);
        existing.load('value, isNullObject');
        return context.sync().then(function () {
            if (existing && !existing.isNullObject && typeof existing.value === 'string' && existing.value.trim()) {
                return existing.value.trim();
            }
            // Belum ada: buat ID acak permanen untuk file ini.
            var newId = 'ea_' + createSessionId() + '_' + Date.now().toString(36);
            try {
                custom.add(EXCEL_AGENT_ID_PROP, newId);
                return context.sync().then(function () { return newId; });
            } catch (e) {
                return newId;
            }
        });
    }).then(function (id) {
        if (done) done(id);
    }).catch(function (e) {
        // Fallback aman: jika permission custom properties dibatasi, tetap jalan tanpa ID.
        if (done) done(null);
    });
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
            setBindingBlocked('Path workbook tidak dapat dibaca. Simpan ke drive lokal atau folder jaringan, lalu buka ulang task pane.');
            done(false);
            return;
        }
        var parts = workbookPath.split(/[\\\/]/);
        bindingBlocked = false;
        // Sesi OpenCode terikat pada workbook+percakapan; binding baru = sesi baru.
        opencodeSessionId = null;
        workbookBinding = { sessionId: createSessionId(), workbookName: parts[parts.length - 1], workbookPath: workbookPath };
        var status = document.getElementById('workbookStatus');
        status.className = 'workbook-status';
        status.textContent = 'Target: ' + workbookBinding.workbookName;
        document.getElementById('userInput').disabled = false;
        document.getElementById('sendBtn').disabled = !document.getElementById('userInput').value.trim();
        // Baca atau tanam ID unik file (Custom Property) agar saat file diduplikat
        // dengan nama berbeda, riwayat chat tetap bisa disambungkan.
        getOrCreateWorkbookId(function (id) {
            if (id && workbookBinding) {
                workbookBinding.workbookId = id;
            }
            done(true);
        });
    });
}

function loadHistory(done) {
    if (!workbookBinding) { if (done) done(); return; }
    var url = SERVER + '/api/history?workbookPath=' + encodeURIComponent(workbookBinding.workbookPath);
    if (workbookBinding.workbookId) {
        url += '&workbookId=' + encodeURIComponent(workbookBinding.workbookId);
    }
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
    pendingHistorySnapshot = {
        generation: historySaveGeneration,
        revision: ++historyRevision,
        chatHistory: chatHistory.slice(-50)
    };
    flushHistorySave();
}

function flushHistorySave(done) {
    if (historyDeleteInFlight || historySaveInFlight || !pendingHistorySnapshot || !workbookBinding) {
        if (done && !historySaveInFlight) done();
        return;
    }
    var snapshot = pendingHistorySnapshot;
    pendingHistorySnapshot = null;
    historySaveInFlight = true;
    xhrPost(SERVER + '/api/history', {
        workbookPath: workbookBinding.workbookPath,
        workbookId: workbookBinding.workbookId || null,
        sessionId: workbookBinding.sessionId,
        revision: snapshot.revision,
        chatHistory: snapshot.chatHistory
    }, function() {
        historySaveInFlight = false;
        if (pendingHistorySnapshot) flushHistorySave(done);
        else if (done) done();
    });
}

function deleteHistory(done) {
    if (!workbookBinding) { if (done) done(false); return; }
    if (historyDeleteInFlight) return;
    historyDeleteInFlight = true;
    historySaveGeneration += 1;
    pendingHistorySnapshot = null;
    function sendDelete() {
        if (historySaveInFlight) {
            setTimeout(sendDelete, 20);
            return;
        }
        var xhr = new XMLHttpRequest();
        xhr.open('DELETE', SERVER + '/api/history', true);
        xhr.setRequestHeader('Content-Type', 'application/json');
        var payload = { workbookPath: workbookBinding.workbookPath };
        if (workbookBinding.workbookId) payload.workbookId = workbookBinding.workbookId;
        xhr.onreadystatechange = function() {
            if (xhr.readyState !== 4) return;
            var ok = xhr.status >= 200 && xhr.status < 300;
            historyDeleteInFlight = false;
            if (ok) {
                pendingHistorySnapshot = null;
                chatHistory = [];
                opencodeSessionId = null; // riwayat baru = sesi OpenCode baru
            } else {
                flushHistorySave();
            }
            if (done) done(ok);
        };
        xhr.onerror = function() {
            historyDeleteInFlight = false;
            flushHistorySave();
            if (done) done(false);
        };
        xhr.send(JSON.stringify(payload));
    }
    sendDelete();
}

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

// Tampilkan kembali percakapan sesi sebelumnya.
function restoreHistory() {
    try {
        if (!chatHistory || !chatHistory.length) return;

        var welcomeEl = document.getElementById('welcomeScreen');
        var chatEl = document.getElementById('chatMessages');
        if (welcomeEl) welcomeEl.className = 'welcome-screen hidden';
        if (chatEl) chatEl.className = 'chat-messages';

        for (var i = 0; i < chatHistory.length; i++) {
            var msg = chatHistory[i];
            if (!msg || (msg.role !== 'user' && msg.role !== 'assistant')) continue;
            var code = msg.role === 'assistant' ? extractCode(msg.content) : null;
            // addMessage akan otomatis menyembunyikan kode jika autoRun true
            addMessage(msg.role, msg.content, code, true);
        }
        addMessage('system', 'Percakapan sebelumnya dipulihkan.', null, true);
    } catch (err) {
        console.error("Error restoring history", err);
    }
}

function setupEventListeners() {
    var userInput = document.getElementById('userInput');
    var sendBtn = document.getElementById('sendBtn');
    var settingsBtn = document.getElementById('settingsBtn');
    var closeSettingsBtn = document.getElementById('closeSettingsBtn');
    var saveSettingsBtn = document.getElementById('saveSettingsBtn');
    var testConnBtn = document.getElementById('testConnBtn');
    var providerSelect = document.getElementById('providerSelect');
    var refreshProvidersBtn = document.getElementById('refreshProvidersBtn');
    var toggleApiKeyBtn = document.getElementById('toggleApiKeyBtn');
    var attachImageBtn = document.getElementById('attachImageBtn');
    var imageInput = document.getElementById('imageInput');

    // Auto-resize textarea
    userInput.oninput = function () {
        userInput.style.height = 'auto';
        userInput.style.height = Math.min(userInput.scrollHeight, 120) + 'px';
        updateComposerState();
    };

    // Send message on Enter, hentikan proses dengan Esc
    userInput.onkeydown = function (e) {
        if (e.keyCode === 27) { // Esc
            if (isProcessing) {
                e.preventDefault();
                stopProcessing('Proses dihentikan oleh pengguna (Esc).');
            }
            return;
        }
        if (e.keyCode === 13 && !e.shiftKey) {
            e.preventDefault();
            if (userInput.value.trim() || pendingImages.length > 0) {
                sendMessage();
            }
        }
    };

    // Send button click
    sendBtn.onclick = function () { sendMessage(); };

    // Tombol Hentikan pada indikator loading
    var stopBtn = document.getElementById('stopBtn');
    if (stopBtn) {
        stopBtn.onclick = function () {
            stopProcessing('Proses dihentikan oleh pengguna.');
        };
    }

    // Esc di mana pun dalam task pane (bukan hanya saat fokus di textarea).
    // Modal pengaturan punya prioritas: Esc menutup modal dulu.
    document.addEventListener('keydown', function (e) {
        if (e.keyCode !== 27) return;
        var modal = document.getElementById('settingsModal');
        if (modal && modal.className.indexOf('hidden') === -1) {
            modal.className = 'modal hidden';
            e.preventDefault();
            return;
        }
        if (isProcessing) {
            e.preventDefault();
            stopProcessing('Proses dihentikan oleh pengguna (Esc).');
        }
    });

    // Connection badge & popover
    var connectionBadge = document.getElementById('connectionBadge');
    var connectionPopover = document.getElementById('connectionPopover');
    var refreshHealthBtn = document.getElementById('refreshHealthBtn');

    if (connectionBadge && connectionPopover) {
        connectionBadge.onclick = function (e) {
            e.stopPropagation();
            var isHidden = connectionPopover.classList.contains('hidden');
            if (isHidden) {
                connectionPopover.classList.remove('hidden');
                connectionBadge.setAttribute('aria-expanded', 'true');
                checkSystemHealth(false);
            } else {
                connectionPopover.classList.add('hidden');
                connectionBadge.setAttribute('aria-expanded', 'false');
            }
        };

        document.addEventListener('click', function (e) {
            if (!connectionPopover.contains(e.target) && e.target !== connectionBadge) {
                connectionPopover.classList.add('hidden');
                connectionBadge.setAttribute('aria-expanded', 'false');
            }
        });
    }

    if (refreshHealthBtn) {
        refreshHealthBtn.onclick = function (e) {
            e.stopPropagation();
            checkSystemHealth(true);
        };
    }

    if (attachImageBtn && imageInput) {
        attachImageBtn.onclick = function () { if (!bindingBlocked && workbookBinding) imageInput.click(); };
        imageInput.onchange = function () {
            addImageFiles(imageInput.files);
            imageInput.value = '';
        };
    }
    userInput.onpaste = function (event) {
        var clipboard = event.clipboardData || (event.originalEvent && event.originalEvent.clipboardData);
        if (!clipboard || !clipboard.items) return;
        var files = [];
        for (var i = 0; i < clipboard.items.length; i++) {
            if (clipboard.items[i].kind === 'file' && /^image\//i.test(clipboard.items[i].type || '')) {
                var file = clipboard.items[i].getAsFile();
                if (file) files.push(file);
            }
        }
        if (files.length) {
            event.preventDefault();
            addImageFiles(files);
        }
    };

    // Settings modal (disederhanakan): 1 dropdown provider (Hermes + Manual),
    // 1 kolom model dengan saran, API key, endpoint, dan OpenCode.
    settingsBtn.onclick = function () {
        document.getElementById('settingsModal').className = 'modal';
        document.getElementById('apiKey').value = '';
        setApiKeyVisible(false);
        document.getElementById('apiEndpoint').value = settings.apiEndpoint;
        document.getElementById('modelInput').value = settings.model;
        document.getElementById('useOpencode').checked = settings.useOpencode === true;
        document.getElementById('opencodeModel').value = settings.opencodeModel;
        populateProviderDropdown();
        populateModelDatalist(Array.isArray(settings.modelList) ? settings.modelList : []);
        hideSettingsStatus();
    };

    // ---- API Key show/hide toggle ----
    function setApiKeyVisible(visible) {
        var el = document.getElementById('apiKey');
        if (!el) return;
        el.type = visible ? 'text' : 'password';
        if (toggleApiKeyBtn) toggleApiKeyBtn.textContent = visible ? 'Sembunyi' : 'Lihat';
    }
    if (toggleApiKeyBtn) {
        toggleApiKeyBtn.onclick = function () {
            var el = document.getElementById('apiKey');
            setApiKeyVisible(el && el.type === 'password');
        };
    }

    // ---- Provider: satu dropdown berisi provider Hermes + opsi Manual ----
    var hermesCatalog = [];

    function populateProviderDropdown() {
        if (!providerSelect) return;
        providerSelect.innerHTML = '';
        var optManual = document.createElement('option');
        optManual.value = '';
        optManual.textContent = '\u2014 Manual \u2014';
        providerSelect.appendChild(optManual);
        var found = false;
        for (var i = 0; i < hermesCatalog.length; i++) {
            var p = hermesCatalog[i];
            var opt = document.createElement('option');
            opt.value = p.name;
            opt.textContent = p.name + ' (' + p.modelCount + ' model' + (p.hasKey ? '' : ', tanpa key') + ')';
            providerSelect.appendChild(opt);
            if (p.name === settings.activeProvider) found = true;
        }
        // Tandai provider yang sedang aktif bila masih ada di katalog Hermes.
        providerSelect.value = found ? settings.activeProvider : '';
    }

    // Saran model untuk kolom model (datalist). Di WebView lama yang tidak
    // mendukung datalist, kolom tetap berfungsi sebagai input teks biasa.
    function populateModelDatalist(list) {
        var dl = document.getElementById('modelDatalist');
        if (!dl) return;
        dl.innerHTML = '';
        var seen = {};
        var items = [];
        if (settings.model) { items.push(settings.model); seen[settings.model] = true; }
        for (var i = 0; i < list.length; i++) {
            if (list[i] && !seen[list[i]]) { seen[list[i]] = true; items.push(list[i]); }
        }
        for (var j = 0; j < items.length; j++) {
            var opt = document.createElement('option');
            opt.value = items[j];
            dl.appendChild(opt);
        }
    }

    function updateProviderHint(ok, hermesHome) {
        var hint = document.getElementById('providerHint');
        if (!hint) return;
        hint.textContent = ok
            ? 'Ditemukan ' + hermesCatalog.length + ' provider Hermes' + (hermesHome ? ' di ' + hermesHome : '') + '.'
            : 'Konfigurasi Hermes tidak ditemukan \u2014 pakai mode Manual.';
    }

    function loadHermesProviders(done) {
        var xhr = new XMLHttpRequest();
        xhr.open('GET', SERVER + '/api/hermes/providers', true);
        xhr.onreadystatechange = function () {
            if (xhr.readyState !== 4) return;
            var data = null;
            try { data = JSON.parse(xhr.responseText); } catch (e) { }
            if (xhr.status === 200 && data && data.ok) {
                hermesCatalog = data.providers || [];
                updateProviderHint(true, data.hermesHome);
                if (done) done(true);
            } else {
                hermesCatalog = [];
                updateProviderHint(false);
                if (done) done(false);
            }
        };
        xhr.onerror = function () { hermesCatalog = []; updateProviderHint(false); if (done) done(false); };
        xhr.send();
    }

    // Pilih provider -> impor config dari Hermes, isi form otomatis.
    if (providerSelect) {
        providerSelect.onchange = function () {
            var name = providerSelect.value;
            if (!name) {
                document.getElementById('apiKey').value = '';
                setApiKeyVisible(false);
                document.getElementById('apiEndpoint').value = '';
                document.getElementById('modelInput').value = '';
                populateModelDatalist([]);
                settings.activeProvider = '';
                settings.model = '';
                showSettingsStatus('Mode manual: isi endpoint, key, dan model, lalu Simpan.', 'loading');
                return;
            }
            var chosenModel = document.getElementById('modelInput').value.trim();
            providerSelect.disabled = true;
            showSettingsStatus('Memuat provider "' + name + '"...', 'loading');
            xhrPost(SERVER + '/api/hermes/import', { name: name, model: chosenModel }, function (status, data) {
                providerSelect.disabled = false;
                if (status >= 200 && status < 300 && data && data.ok) {
                    var cfg = data.config || {};
                    settings.apiEndpoint = cfg.apiEndpoint || '';
                    settings.model = cfg.model || data.model || '';
                    settings.activeProvider = data.providerName || '';
                    settings.hasApiKey = cfg.hasApiKey === true;
                    if (Array.isArray(cfg.modelList)) settings.modelList = cfg.modelList;
                    if (cfg.providers) settings.providers = cfg.providers;
                    document.getElementById('apiKey').value = '';
                    setApiKeyVisible(false);
                    document.getElementById('apiEndpoint').value = settings.apiEndpoint;
                    document.getElementById('modelInput').value = settings.model;
                    populateModelDatalist(data.models || cfg.modelList || []);
                    saveSettingsLocalOnly();
                    showSettingsStatus('Provider "' + name + '" dimuat (' + (data.models || []).length + ' model' + (settings.hasApiKey ? '' : ', tanpa API key') + '). Klik Simpan agar permanen.', settings.hasApiKey ? 'success' : 'error');
                    if (typeof checkSystemHealth === 'function') checkSystemHealth();
                } else {
                    showSettingsStatus('Gagal memuat provider: ' + ((data && data.error) || 'server error'), 'error');
                }
            });
        };
    }

    if (refreshProvidersBtn) {
        refreshProvidersBtn.onclick = function () {
            refreshProvidersBtn.disabled = true;
            showSettingsStatus('Membaca konfigurasi Hermes...', 'loading');
            loadHermesProviders(function (ok) {
                refreshProvidersBtn.disabled = false;
                populateProviderDropdown();
                showSettingsStatus(ok ? 'Daftar provider dimuat.' : 'Hermes tidak ditemukan.', ok ? 'success' : 'error');
            });
        };
    }

    // Muat katalog Hermes sekali saat task pane siap.
    loadHermesProviders();


    closeSettingsBtn.onclick = function () {
        document.getElementById('settingsModal').className = 'modal hidden';
    };

    saveSettingsBtn.onclick = function () {
        var newApiKey = document.getElementById('apiKey').value.trim();
        settings.apiEndpoint = document.getElementById('apiEndpoint').value.trim();
        settings.model = document.getElementById('modelInput').value.trim();
        settings.useOpencode = document.getElementById('useOpencode').checked === true;
        settings.opencodeModel = document.getElementById('opencodeModel').value.trim();

        saveSettingsBtn.disabled = true;
        showSettingsStatus('Menyimpan...', 'loading');

        saveSettings(function (ok) {
            saveSettingsBtn.disabled = false;
            if (ok) {
                // Simpan profil identitas juga (agar model mengenali pemilik)
                saveProfil(function () {
                    showSettingsStatus('Tersimpan permanen. Aman ditutup.', 'success');
                    checkSystemHealth(false);
                    setTimeout(function () {
                        document.getElementById('settingsModal').className = 'modal hidden';
                    }, 900);
                });
            } else {
                showSettingsStatus('Tersimpan lokal saja (server mati).', 'error');
            }
        }, newApiKey);
    };

    // Test connection button
    testConnBtn.onclick = function () {
        var apiKey = document.getElementById('apiKey').value.trim();
        var apiEndpoint = document.getElementById('apiEndpoint').value.trim();
        var model = document.getElementById('modelInput').value.trim();

        if (!apiEndpoint) {
            showSettingsStatus('Isi API Endpoint dulu.', 'error');
            return;
        }
        if (!model) {
            showSettingsStatus('Isi nama Model dulu.', 'error');
            return;
        }

        testConnBtn.disabled = true;
        saveSettingsBtn.disabled = true;
        showSettingsStatus('Menghubungkan...', 'loading');

        testConnection(apiEndpoint, apiKey, model, function (result) {
            testConnBtn.disabled = false;
            saveSettingsBtn.disabled = false;

            if (result.ok) {
                showSettingsStatus('Berhasil! Model "' + (result.model || model) + '" merespons.', 'success');
            } else {
                showSettingsStatus('Gagal: ' + result.error, 'error');
            }
        });
    };


    // Hapus riwayat percakapan
    var clearHistoryBtn = document.getElementById('clearHistoryBtn');
    if (clearHistoryBtn) {
        clearHistoryBtn.onclick = function () {
            assertBindingCurrent(function(error) {
                if (error) {
                    addMessage('system', error.error);
                    return;
                }
                deleteHistory(function(ok) {
                    if (!ok) {
                        addMessage('system', 'Gagal menghapus riwayat percakapan.');
                        return;
                    }
                    chatHistory = [];
                    document.getElementById('chatMessages').innerHTML = '';
                    document.getElementById('chatMessages').className = 'chat-messages hidden';
                    document.getElementById('welcomeScreen').className = 'welcome-screen';
                });
            });
        };
    }

    // Example cards
    var cards = document.querySelectorAll('.example-card');
    for (var i = 0; i < cards.length; i++) {
        (function (card) {
            card.onclick = function () {
                if (bindingBlocked || !workbookBinding) return;
                userInput.value = card.getAttribute('data-prompt');
                sendBtn.disabled = false;
                sendMessage();
            };
        })(cards[i]);
    }
}

function updateComposerState() {
    var input = document.getElementById('userInput');
    var send = document.getElementById('sendBtn');
    var attach = document.getElementById('attachImageBtn');
    if (send) send.disabled = bindingBlocked || !workbookBinding || isProcessing || (!input.value.trim() && pendingImages.length === 0);
    if (attach) attach.disabled = bindingBlocked || !workbookBinding || isProcessing || pendingImages.length >= MAX_PENDING_IMAGES;
}

function addImageFiles(files) {
    if (!files || !files.length) return;
    var available = MAX_PENDING_IMAGES - pendingImages.length;
    if (available <= 0) {
        addMessage('system', 'Maksimal 5 gambar per pesan.');
        return;
    }
    var count = Math.min(files.length, available);
    if (files.length > available) addMessage('system', 'Hanya ' + available + ' gambar tambahan yang dapat dilampirkan.');
    for (var i = 0; i < count; i++) readPendingImage(files[i]);
}

function readPendingImage(file) {
    if (!file || !ALLOWED_IMAGE_TYPES[String(file.type || '').toLowerCase()]) {
        addMessage('system', 'Format gambar tidak didukung. Gunakan PNG, JPG, atau WEBP.');
        return;
    }
    if (file.size > MAX_IMAGE_BYTES) {
        addMessage('system', 'Ukuran setiap gambar maksimal 5 MB.');
        return;
    }
    var reader = new FileReader();
    reader.onload = function () {
        if (pendingImages.length >= MAX_PENDING_IMAGES) return;
        pendingImages.push({
            name: file.name || ('screenshot-' + (pendingImages.length + 1) + '.png'),
            type: file.type.toLowerCase(),
            size: file.size,
            data: reader.result
        });
        renderImagePreviews();
        updateComposerState();
    };
    reader.onerror = function () { addMessage('system', 'Gagal membaca gambar.'); };
    reader.readAsDataURL(file);
}

function removePendingImage(index) {
    pendingImages.splice(index, 1);
    renderImagePreviews();
    updateComposerState();
}

function renderImagePreviews() {
    var list = document.getElementById('imagePreviewList');
    if (!list) return;
    if (!pendingImages.length) {
        list.innerHTML = '';
        list.className = 'image-preview-list hidden';
        return;
    }
    var html = '';
    for (var i = 0; i < pendingImages.length; i++) {
        html += '<div class="image-preview-item">' +
            '<img src="' + escapeHtml(pendingImages[i].data) + '" alt="' + escapeHtml(pendingImages[i].name) + '">' +
            '<button type="button" class="image-remove-btn" onclick="removePendingImage(' + i + ')" aria-label="Hapus gambar">&times;</button>' +
            '</div>';
    }
    list.innerHTML = html;
    list.className = 'image-preview-list';
}

// XMLHttpRequest helper (IE11 compatible, replaces fetch)
function xhrPost(url, data, callback, timeoutMs) {
    var xhr = new XMLHttpRequest();
    var completed = false;
    trackXhr(xhr);
    function finish(status, response) {
        if (completed) return;
        completed = true;
        untrackXhr(xhr);
        callback(status, response);
    }
    xhr.open('POST', url, true);
    xhr.setRequestHeader('Content-Type', 'application/json');
    if (timeoutMs) {
        xhr.timeout = timeoutMs;
        xhr.ontimeout = function () {
            finish(0, { error: { message: 'Request timeout setelah ' + Math.round(timeoutMs / 1000) + ' detik.' } });
        };
    }
    xhr.onabort = function () {
        finish(0, { cancelled: true, error: { message: 'Dibatalkan oleh pengguna.' } });
    };
    xhr.onreadystatechange = function () {
        if (xhr.readyState === 4 && !completed) {
            var response = null;
            try {
                response = JSON.parse(xhr.responseText);
            } catch (e) {
                response = { error: { message: xhr.statusText || 'Parse error' } };
            }
            finish(xhr.status, response);
        }
    };
    xhr.onerror = function () {
        finish(0, { error: { message: 'Network error' } });
    };
    xhr.send(JSON.stringify(data));
}

// Generic request dengan method bebas (mis. DELETE) + body JSON. Dipakai provider profiles.
function xhrRequest(method, url, data, callback, timeoutMs) {
    var xhr = new XMLHttpRequest();
    var completed = false;
    function finish(status, response) {
        if (completed) return;
        completed = true;
        callback(status, response);
    }
    xhr.open(method, url, true);
    xhr.setRequestHeader('Content-Type', 'application/json');
    if (timeoutMs) {
        xhr.timeout = timeoutMs;
        xhr.ontimeout = function () {
            finish(0, { error: { message: 'Request timeout setelah ' + Math.round(timeoutMs / 1000) + ' detik.' } });
        };
    }
    xhr.onreadystatechange = function () {
        if (xhr.readyState === 4 && !completed) {
            var response = null;
            try {
                response = JSON.parse(xhr.responseText);
            } catch (e) {
                response = { error: { message: xhr.statusText || 'Parse error' } };
            }
            finish(xhr.status, response);
        }
    };
    xhr.onerror = function () {
        finish(0, { error: { message: 'Network error' } });
    };
    xhr.send(data ? JSON.stringify(data) : null);
}

// Simpan cache setting non-secret ke localStorage saja (tanpa hit server).
// Dipakai saat mengaktifkan provider: server sudah punya sumber kebenaran.
function saveSettingsLocalOnly() {
    try {
        localStorage.setItem('excelAgentSettings', JSON.stringify({
            apiEndpoint: settings.apiEndpoint,
            model: settings.model,
            autoRun: settings.autoRun,
            useOpencode: settings.useOpencode,
            opencodeModel: settings.opencodeModel,
            activeProvider: settings.activeProvider
        }));
    } catch (e) { }
}

function formatApiError(status, data) {
    if (data && typeof data.error === 'string' && data.error) return data.error;
    if (data && data.error && typeof data.error.message === 'string' && data.error.message) return data.error.message;
    if (data && data.error && typeof data.error === 'object' && data.error.message) return String(data.error.message);
    if (data && typeof data.message === 'string' && data.message) return data.message;
    if (!status) return 'Koneksi ke server lokal gagal.';
    if (status >= 500) return 'Server error (HTTP ' + status + '). Coba lagi.';
    if (status === 429) return 'Rate limit tercapai. Tunggu sebentar lalu coba lagi.';
    if (status === 401 || status === 403) return 'API key tidak valid atau tidak diizinkan (HTTP ' + status + ').';
    if (status === 504) return 'Gateway timeout. AI tidak merespons tepat waktu.';
    return 'HTTP ' + status;
}

// --- Progress polling ---
var progressPollTimer = null;

function showProgress(text) {
    var el = document.getElementById('progressText');
    if (el) {
        el.textContent = text;
        el.className = 'progress-text';
    }
}

function hideProgress() {
    var el = document.getElementById('progressText');
    if (el) {
        el.textContent = '';
        el.className = 'progress-text hidden';
    }
}

function startProgressPolling() {
    stopProgressPolling();
    if (!workbookBinding) return;
    var url = SERVER + '/api/exec-progress?workbookPath=' + encodeURIComponent(workbookBinding.workbookPath);
    progressPollTimer = setInterval(function () {
        var xhr = new XMLHttpRequest();
        xhr.open('GET', url, true);
        xhr.onreadystatechange = function () {
            if (xhr.readyState === 4 && xhr.status === 200) {
                try {
                    var data = JSON.parse(xhr.responseText);
                    if (data.progress) showProgress(data.progress);
                } catch (e) { }
            }
        };
        xhr.send();
    }, 500);
}

function stopProgressPolling() {
    if (progressPollTimer) {
        clearInterval(progressPollTimer);
        progressPollTimer = null;
    }
    hideProgress();
}

// Ambil snippet JS (Office.js API) dari balasan AI.
function extractCode(content) {
    // Pertahankan tag JavaScript lama hanya untuk rendering riwayat terdahulu.
    var m = content.match(/```(?:javascript|js|)\r?\n([\s\S]*?)```/i);
    if (m && m[1].replace(/^\s+|\s+$/g, '')) {
        return m[1].replace(/^\s+|\s+$/g, '');
    }
    // Fallback: fence tidak tertutup
    m = content.match(/```(?:javascript|js|)\r?\n([\s\S]*)$/i);
    if (m && m[1].replace(/^\s+|\s+$/g, '')) {
        return m[1].replace(/^\s+|\s+$/g, '');
    }
    return null;
}

function normalizeEndpoint(ep) {
    if (!ep) return ep;
    ep = ep.replace(/\/+$/, '');
    if (!/\/chat\/completions$/i.test(ep)) {
        if (/\/v1$/i.test(ep)) {
            ep = ep + '/chat/completions';
        } else if (!/\/(v1\/chat\/completions|chat\/completions)$/i.test(ep)) {
            ep = ep + '/v1/chat/completions';
        }
    }
    return ep;
}

// Test connection via proxy
function testConnection(endpoint, apiKey, model, callback) {
    endpoint = normalizeEndpoint(endpoint);
    var request = {
        apiEndpoint: endpoint,
        payload: {
            model: model,
            messages: [{ role: 'user', content: 'Hi' }],
            stream: false
        }
    };
    if (apiKey) request.apiKey = apiKey;
    xhrPost('https://localhost:3000/api/proxy', request, function (status, data) {
        if (status >= 200 && status < 300 && data.choices && data.choices.length > 0) {
            callback({ ok: true, model: data.model || model });
        } else {
            var errMsg = formatApiError(status, data);
            callback({ ok: false, error: errMsg });
        }
    }, 30000);
}

// Show/hide settings status
function showSettingsStatus(msg, type) {
    var el = document.getElementById('settingsStatus');
    el.textContent = msg;
    el.className = 'settings-status ' + type;
}

function hideSettingsStatus() {
    var el = document.getElementById('settingsStatus');
    el.className = 'settings-status hidden';
}

// System Health & Connection Status
var healthPollTimer = null;

function checkSystemHealth(showLoadingState) {
    var badge = document.getElementById('connectionBadge');
    var badgeLabel = document.getElementById('badgeLabel');
    var serverItem = document.getElementById('healthServerItem');
    var serverDesc = document.getElementById('healthServerDesc');
    var routerItem = document.getElementById('healthRouterItem');
    var routerDesc = document.getElementById('healthRouterDesc');
    var pythonItem = document.getElementById('healthPythonItem');
    var pythonDesc = document.getElementById('healthPythonDesc');

    if (!badge) return;

    if (showLoadingState) {
        badge.className = 'connection-badge status-loading';
        if (badgeLabel) badgeLabel.textContent = 'Checking...';
    }

    var xhr = new XMLHttpRequest();
    xhr.open('GET', SERVER + '/api/health', true);
    xhr.timeout = 5000;

    xhr.onreadystatechange = function () {
        if (xhr.readyState !== 4) return;

        if (xhr.status === 200) {
            try {
                var res = JSON.parse(xhr.responseText);
                var isServerOk = res.server && res.server.status === 'online';
                var isRouterOk = res.router && res.router.ok;
                var isPythonOk = res.python && res.python.ok;

                // Server Status UI
                if (serverItem && serverDesc) {
                    var sDot = serverItem.querySelector('.item-dot');
                    if (sDot) sDot.className = 'item-dot dot-green';
                    serverDesc.textContent = 'Port ' + (res.server.port || 3000) + ' (Online)';
                }

                // Router Status UI
                if (routerItem && routerDesc) {
                    var rDot = routerItem.querySelector('.item-dot');
                    if (isRouterOk) {
                        if (rDot) rDot.className = 'item-dot dot-green';
                        var latencyStr = res.router.latencyMs ? ' (' + res.router.latencyMs + 'ms)' : '';
                        var modelStr = res.router.activeModel ? ' \u2022 ' + res.router.activeModel : '';
                        routerDesc.textContent = 'Online' + latencyStr + modelStr;
                    } else {
                        if (rDot) rDot.className = 'item-dot ' + (res.router.status === 'unconfigured' ? 'dot-yellow' : 'dot-red');
                        routerDesc.textContent = res.router.message || (res.router.status === 'unconfigured' ? 'Belum dikonfigurasi' : 'Offline');
                    }
                }

                // Python Status UI
                if (pythonItem && pythonDesc) {
                    var pDot = pythonItem.querySelector('.item-dot');
                    if (isPythonOk) {
                        if (pDot) pDot.className = 'item-dot ' + (res.python.hasWin32com ? 'dot-green' : 'dot-yellow');
                        pythonDesc.textContent = res.python.hasWin32com ? 'win32com Ready' : 'Python OK (No win32com)';
                    } else {
                        if (pDot) pDot.className = 'item-dot dot-red';
                        pythonDesc.textContent = res.python.message || 'Offline';
                    }
                }

                // Main Badge state
                if (isServerOk && isRouterOk && isPythonOk && (!res.python || res.python.hasWin32com)) {
                    badge.className = 'connection-badge status-online';
                    if (badgeLabel) badgeLabel.textContent = 'Ready';
                    badge.title = 'Sistem Siap (Server, Router, Python COM Online)';
                } else if (isServerOk && (isRouterOk || isPythonOk)) {
                    badge.className = 'connection-badge status-degraded';
                    if (badgeLabel) badgeLabel.textContent = !isRouterOk ? 'Router Offline' : 'COM Degraded';
                    badge.title = 'Koneksi Terbatas (Klik untuk detail)';
                } else {
                    badge.className = 'connection-badge status-offline';
                    if (badgeLabel) badgeLabel.textContent = 'Offline';
                    badge.title = 'Layanan bermasalah (Klik untuk detail)';
                }
            } catch (e) {
                badge.className = 'connection-badge status-offline';
                if (badgeLabel) badgeLabel.textContent = 'Error';
            }
        } else {
            // Server offline
            badge.className = 'connection-badge status-offline';
            if (badgeLabel) badgeLabel.textContent = 'Server Offline';
            badge.title = 'Server lokal tidak merespons di ' + SERVER;

            if (serverItem && serverDesc) {
                var sDot = serverItem.querySelector('.item-dot');
                if (sDot) sDot.className = 'item-dot dot-red';
                serverDesc.textContent = 'Offline (Server mati)';
            }
            if (routerItem && routerDesc) {
                var rDot = routerItem.querySelector('.item-dot');
                if (rDot) rDot.className = 'item-dot dot-gray';
                routerDesc.textContent = 'Tidak dapat diperiksa';
            }
            if (pythonItem && pythonDesc) {
                var pDot = pythonItem.querySelector('.item-dot');
                if (pDot) pDot.className = 'item-dot dot-gray';
                pythonDesc.textContent = 'Tidak dapat diperiksa';
            }
        }
    };

    xhr.onerror = function () {
        badge.className = 'connection-badge status-offline';
        if (badgeLabel) badgeLabel.textContent = 'Server Offline';
    };

    xhr.send();
}

function startHealthPolling() {
    if (healthPollTimer) clearInterval(healthPollTimer);
    checkSystemHealth(true);
    healthPollTimer = setInterval(function () {
        checkSystemHealth(false);
    }, 20000);
}

// Get cell address helper (row/col 0-based to Excel address)
function getCellAddress(row, col) {
    var colLetter = '';
    var c = col;
    while (c >= 0) {
        colLetter = String.fromCharCode(65 + (c % 26)) + colLetter;
        c = Math.floor(c / 26) - 1;
    }
    return colLetter + (row + 1);
}

// Build context message from workbook schema
function buildContextMessage(schema) {
    var msg = "=== KONTEKS WORKBOOK ===\n";
    
    // Multi-sheet: tampilkan semua sheet dengan dimensi + header baris pertama
    if (schema.workbook.sheets.length > 1) {
        msg += "Workbook memiliki " + schema.workbook.sheets.length + " sheet:\n";
        for (var i = 0; i < schema.workbook.sheets.length; i++) {
            var s = schema.workbook.sheets[i];
            var dim = (s.rowCount !== undefined) ? " (" + s.rowCount + " baris x " + s.colCount + " kolom)" : "";
            var hdr = s.headers ? " | baris pertama: [" + s.headers.join(", ") + "]" : "";
            msg += "  - " + s.name + dim + hdr + "\n";
        }
        msg += "Gunakan wb.Worksheets(\"NamaSheet\") untuk mengakses sheet tertentu.\n";
        msg += "Jika butuh data dari sheet lain, BACA dulu lewat kode sebelum menulis (jangan menebak nama kolom).\n";
    } else {
        msg += "Sheet: [" + schema.workbook.sheets.map(function(s) { return s.name; }).join(", ") + "]\n";
    }
    msg += "Sheet aktif: " + schema.activeSheet.name + "\n";
    
    if (schema.activeSheet.rowCount) {
        msg += "Ukuran: " + schema.activeSheet.rowCount + " baris x " + schema.activeSheet.colCount + " kolom\n";
    }
    
    if (schema.activeSheet.headers) {
        msg += "Header: " + JSON.stringify(schema.activeSheet.headers) + "\n";
    }
    
    if (schema.activeSheet.columnWidths) {
        msg += "Lebar kolom saat ini: " + JSON.stringify(schema.activeSheet.columnWidths) + "\n";
    }
    
    if (schema.activeSheet.sampleData) {
        msg += "Sample data:\n" + JSON.stringify(schema.activeSheet.sampleData) + "\n";
    }
    
    if (schema.activeSheet.formulas && schema.activeSheet.formulas.length > 0) {
        msg += "Formula: " + schema.activeSheet.formulas.map(function(f) {
            return f.cell + "=" + f.formula;
        }).join(", ") + "\n";
    }
    
    msg += "\nPENTING: Anda HARUS menulis kode Python (win32com.client) untuk mengeksekusi permintaan user. JANGAN hanya memberi analisis.\n";
    
    return msg;
}

// Get rich workbook schema for AI context
function getWorkbookSchema(callback) {
    if (typeof Excel === 'undefined') {
        callback({ error: "Excel API tidak tersedia" });
        return;
    }
    
    Excel.run(function(context) {
        var allSheets = context.workbook.worksheets;
        allSheets.load("name");
        
        var sheet = context.workbook.worksheets.getActiveWorksheet();
        sheet.load("name");
        
        var usedRange = sheet.getUsedRangeOrNullObject();
        // Hanya load dimensi dulu; values/formulas di-load terpisah setelah sync
        // agar bisa membatasi ukuran data dan menghindari overload Excel.
        usedRange.load("address, rowCount, columnCount");
        
        // Muat used range semua sheet untuk dimensi multi-sheet
        var sheetRanges = [];
        
        return context.sync().then(function() {
            for (var i = 0; i < allSheets.items.length; i++) {
                var sr = allSheets.items[i].getUsedRangeOrNullObject();
                sr.load("rowCount, columnCount");
                sheetRanges.push(sr);
            }
            return context.sync();
        }).then(function() {
            // Muat baris pertama (header) setiap sheet supaya AI tahu struktur
            // semua sheet, bukan hanya sheet aktif. Ini mengurangi halusinasi
            // nama kolom/sheet saat tugas melibatkan sheet lain.
            var headerRanges = [];
            var maxHeaderSheets = Math.min(sheetRanges.length, 15);
            for (var h = 0; h < sheetRanges.length; h++) {
                var srh = sheetRanges[h];
                if (h < maxHeaderSheets && srh && !srh.isNullObject && srh.rowCount > 0 && srh.columnCount > 0) {
                    var hdr = srh.getRow(0);
                    hdr.load("values");
                    headerRanges.push(hdr);
                } else {
                    headerRanges.push(null);
                }
            }
            // Load active sheet data: cap baris dan kolom supaya tidak overload
            var activeValues = null;
            var activeFormulas = null;
            var activeHeader = null;
            if (!usedRange.isNullObject && usedRange.rowCount > 0 && usedRange.columnCount > 0) {
                var capRows = Math.min(usedRange.rowCount, 50);
                var capCols = Math.min(usedRange.columnCount, 20);
                activeValues = usedRange.getCell(0, 0).getResizedRange(capRows - 1, capCols - 1);
                activeValues.load("values");
                // Formulas: hanya cap area yang sama
                activeFormulas = usedRange.getCell(0, 0).getResizedRange(capRows - 1, capCols - 1);
                activeFormulas.load("formulas");
                activeHeader = usedRange.getRow(0);
                activeHeader.load("values");
            }
            return context.sync().then(function() {
                var schema = {
                    workbook: { sheetCount: allSheets.items.length, sheets: [] },
                    activeSheet: { name: sheet.name }
                };

                for (var i = 0; i < allSheets.items.length; i++) {
                    var sheetInfo = { name: allSheets.items[i].name };
                    if (sheetRanges[i] && !sheetRanges[i].isNullObject) {
                        sheetInfo.rowCount = sheetRanges[i].rowCount;
                        sheetInfo.colCount = sheetRanges[i].columnCount;
                    }
                    var hdrLoaded = headerRanges[i];
                    if (hdrLoaded && hdrLoaded.values && hdrLoaded.values.length > 0) {
                        var hdrRow = hdrLoaded.values[0];
                        var hdrList = [];
                        for (var hc = 0; hc < Math.min(hdrRow.length, 15); hc++) {
                            var hv = hdrRow[hc];
                            if (hv === null || hv === undefined || hv === '') continue;
                            var hs = String(hv);
                            if (hs.length > 24) hs = hs.substring(0, 24) + '..';
                            hdrList.push(hs);
                        }
                        if (hdrList.length > 0) sheetInfo.headers = hdrList;
                    }
                    schema.workbook.sheets.push(sheetInfo);
                }

                if (!usedRange.isNullObject) {
                    schema.activeSheet.range = usedRange.address;
                    schema.activeSheet.rowCount = usedRange.rowCount;
                    schema.activeSheet.colCount = usedRange.columnCount;

                    if (activeHeader) {
                        schema.activeSheet.headers = activeHeader.values[0];
                    }

                    // Get column widths (approximate from header length)
                    var colWidths = {};
                    var headerRow = activeHeader ? activeHeader.values[0] : null;
                    for (var c = 0; c < Math.min(usedRange.columnCount, 20); c++) {
                        var colLetter = getCellAddress(0, c).replace(/[0-9]/g, '');
                        var headerVal = headerRow && headerRow[c] ? String(headerRow[c]) : '';
                        colWidths[colLetter] = headerVal.length > 0 ? Math.max(headerVal.length + 2, 8) : 8;
                    }
                    schema.activeSheet.columnWidths = colWidths;

                    if (activeValues) {
                        var sampleSize = Math.min(6, activeValues.values.length);
                        schema.activeSheet.sampleData = activeValues.values.slice(0, sampleSize);
                    }

                    if (activeFormulas) {
                        var formulas = [];
                        var fRows = activeFormulas.formulas;
                        for (var r = 0; r < fRows.length; r++) {
                            for (var c2 = 0; c2 < fRows[r].length; c2++) {
                                var f = fRows[r][c2];
                                if (typeof f === 'string' && f.charAt(0) === '=') {
                                    formulas.push({ cell: getCellAddress(r, c2), formula: f });
                                }
                            }
                        }
                        schema.activeSheet.formulas = formulas.slice(0, 20);
                    }
                }

                callback(null, schema);
            });
        });
    }).catch(function(err) {
        callback({ error: err.message });
    });
}

// Ambil metadata workbook via MCP server.
// Untuk URL OneDrive, server.js akan coba konversi ke path lokal sync folder.
function getMcpWorkbookInfo(callback) {
    if (!workbookBinding) { callback(null); return; }
    var wp = workbookBinding.workbookPath;
    if (!wp) { callback(null); return; }
    
    xhrPost(SERVER + '/api/mcp/workbook-info', { filepath: wp }, function(status, data) {
        if (status >= 200 && status < 300 && data.ok && data.data) {
            callback(data.data);
        } else {
            callback(null);
        }
    }, 15000);
}

// Send message
function sendMessage() {
    if (isProcessing) return; // Deduplication
    if (bindingBlocked || !workbookBinding) return;
    
    var userInput = document.getElementById('userInput');
    var message = userInput.value.trim();

    if (!message && pendingImages.length === 0) return;
    if (!message) message = 'Analisis gambar ini sebagai referensi workbook.';

    // Giliran baru: bersihkan status pembatalan sebelumnya.
    cancelRequested = false;
    activeXhrs = [];

    isProcessing = true;
    assertBindingCurrent(function(bindingError) {
        if (bindingError) {
            addMessage('system', bindingError.error);
            isProcessing = false;
            return;
        }
        sendBoundMessage(message);
    });
}

function sendBoundMessage(message) {
    var userInput = document.getElementById('userInput');
    activeTurnImages = pendingImages.slice(0);
    pendingImages = [];
    renderImagePreviews();
    addMessage('user', message, null, true, null, activeTurnImages);

    // Siklus laporan & ekspektasi jawaban berlaku per pesan dari user.
    reportCycles = 0;
    noCodeRetries = 0;
    expectingAnswer = false;

    userInput.value = '';
    userInput.style.height = 'auto';
    updateComposerState();

    document.getElementById('loadingIndicator').className = 'loading';
    document.getElementById('welcomeScreen').className = 'welcome-screen hidden';
    document.getElementById('chatMessages').className = 'chat-messages';

    chatHistory.push({
        role: 'user',
        content: message + (activeTurnImages.length ? '\n[Lampiran: ' + activeTurnImages.length + ' gambar]' : '')
    });
    saveHistory();

    if (!settings.hasApiKey) {
        addMessage('assistant', 'Silakan atur API Key terlebih dahulu. Klik ikon gear di pojok kanan atas.');
        document.getElementById('loadingIndicator').className = 'loading hidden';
        isProcessing = false;
        activeTurnImages = [];
        updateComposerState();
        return;
    }
    
    // Gambar bisa dianalisis baik dengan OpenCode maupun mode proxy biasa
    // (LLM vision-capable mendukung input image-base64 / multimodal)
    
    // System prompt with Python win32com cheat sheet.
    // Profil pengguna + memori jangka panjang disuntik di depan
    // agar model mengenali pemilik dan belajar dari sesi sebelumnya.
    var systemPrompt =
        buildProfilPrompt() +
        buildMemoriPrompt() +
        MEMORY_PROTOCOL +
        '### IDENTITAS ###\n' +
        'Anda adalah AI EXCEL AGENT yang TERPASANG DI MICROSOFT EXCEL.\n' +
        'Anda punya AKSES PENUH ke workbook target via Python win32com.client COM automation.\n' +
        'Kode yang Anda tulis akan DIEKSEKUSI OTOMATIS oleh Python di belakang layar.\n' +
        'Variabel `wb` adalah satu-satunya workbook target untuk sesi task pane ini.\n' +
        'JANGAN gunakan `xl.ActiveWorkbook`, `xl.Workbooks(...)`, atau membuka workbook lain untuk memilih target.\n' +
        'Gunakan `ws = wb.ActiveSheet` atau `ws = wb.Worksheets("Nama Sheet")`.\n' +
        'PENTING: Semua blok kode dieksekusi dalam SATU sesi Python. Variabel yang dibuat di blok sebelumnya ' +
        'tersedia di blok berikutnya. JANGAN baca ulang data yang sudah dibaca di blok sebelumnya.\n' +
        'PENTING: Pikirkan dengan matang sebelum menulis kode. Pastikan nama sheet, kolom, dan baris sesuai ' +
        'dengan data yang tersedia di konteks. JANGAN mengarang data.\n\n' +
        
        '### ATURAN MUTLAK (JIKA DILANGGAR = GAGAL) ###\n' +
        '1. Untuk permintaan AKSI / MANIPULASI WORKBOOK, Anda adalah AGENT EKSEKUSI: wajib menulis kode Python (win32com) di dalam tag ```python ... ```. JANGAN hanya memberi analisis.\n' +
        '2. Untuk PERTANYAAN PENJELASAN / KONSULTASI ("apa itu...", "bagaimana cara...", "kenapa...", "apakah bisa..."), berikan jawaban jelas dan ringkas langsung dalam teks Bahasa Indonesia TANPA perlu kode Python.\n' +
        '3. JANGAN PERNAH menolak atau bilang "tidak bisa" untuk tugas Excel yang didukung.\n' +
        '4. JANGAN suruh user: "Salin kode", "Jalankan manual", "Tekan Alt+F11". Bila butuh kode, tulis dan sistem kami yang mengeksekusinya otomatis.\n' +
        '5. Bila menulis kode, HANYA Python (win32com.client) - JANGAN VBA atau Office.js.\n' +
        '6. WAJIB bungkus kode dengan tag ```python ... ``` (huruf kecil).\n' +
        '7. Jika error terjadi pada kode, PERBAIKI kode lalu kirim ulang.\n' +
         '8a. Conditional Formatting COM pada Excel ini WAJIB memakai 4 argumen lengkap: import pythoncom lalu ' +
         'rng.FormatConditions.Add(2, 1, formula, pythoncom.Missing). Aktifkan sheet sebelum Add(). ' +
         'JANGAN gunakan named argument Formula1=, None, atau hanya 2/3 argumen. Formula harus diawali "=". ' +
         'Untuk status dengan output teks tetap, pakai perbandingan langsung seperti =$M$13="TEKS"; fungsi SEARCH/COUNTIF bisa ditolak COM lokal.\n' +
        '9. Jika tugas memerlukan beberapa langkah, tulis SEMUA langkah dalam SATU respons. ' +
        'JANGAN hanya menulis langkah pertama lalu menunggu. JANGAN bilang "silakan jalankan dulu".\n\n' +
        
        '### API YANG DILARANG (JANGAN GUNAKAN!) ###\n' +
        'Python win32com TIDAK menggunakan syntax Office.js. JANGAN gunakan:\n' +
        '- context.workbook (gunakan wb)\n' +
        '- context.sync() (tidak ada di COM)\n' +
        '- sheet.getRange("A1") (gunakan ws.Range("A1"))\n' +
        '- .values = [[val]] (gunakan .Value = val atau .Value2 = val)\n' +
        '- await, async (tidak ada di COM)\n' +
        '- Excel.run() (tidak ada di COM)\n\n' +
        
        '### PROGRESS REPORTING ###\n' +
        'Fungsi progress(msg) tersedia di lingkungan eksekusi.\n' +
        'Untuk tugas dengan lebih dari 3 langkah, panggil progress("deskripsi") sebelum setiap tahap besar.\n' +
        'Contoh:\n' +
        'progress("Membuat header tabel...")\n' +
        'progress("Menulis data 30 baris...")\n' +
        'progress("Memformat kolom...")\n' +
        'Untuk tugas sederhana 1-2 baris, tidak perlu memanggil progress().\n\n' +
        
        '### MULTI-SHEET ###\n' +
        'Workbook ini mungkin memiliki beberapa sheet. Daftar sheet tersedia di konteks.\n' +
        'Jika user menyebut nama sheet tertentu, gunakan wb.Worksheets("NamaSheet").\n' +
        'Jika user tidak menyebut sheet, gunakan wb.ActiveSheet.\n' +
        'JANGAN membuat sheet baru kecuali diminta.\n' +
        'JANGAN menghapus sheet kecuali diminta.\n' +
        'Untuk tugas MODIFIKASI (membuat/mengubah data), setiap langkah kode harus menghasilkan perubahan nyata ' +
        'pada workbook (menulis data, membuat sheet, memformat, dll).\n\n' +
        
        '### TUGAS ANALISIS / PERTANYAAN ###\n' +
        'Jika user BERTANYA tentang isi workbook (mis. "tombol ini menyimpan ke mana", "apa isi sheet X", ' +
        '"cek rumus Y", "data ini sumbernya dari mana"):\n' +
        '1. Tulis kode Python yang MEMBACA/memeriksa workbook TANPA mengubah apa pun.\n' +
        '2. WAJIB print() semua temuan penting (nama sheet, alamat sel, nama macro, isi kode, nilai, dll).\n' +
        '3. Sistem akan menjalankan kode itu lalu mengembalikan output print() kepada Anda secara otomatis.\n' +
        '4. Setelah menerima output, berikan JAWABAN AKHIR kepada user TANPA kode Python lagi.\n' +
        'ANDA BISA melihat isi file ini lewat kode. JANGAN PERNAH berkata "saya tidak bisa melihat file Anda" ' +
        'dan JANGAN menyuruh user mengecek manual (Alt+F11, klik kanan tombol, screenshot kode, dll).\n' +
        'Contoh inspeksi tombol dan macro:\n' +
        '```python\n' +
        'ws = wb.Worksheets("NamaSheet")\n' +
        'for i in range(1, ws.Shapes.Count + 1):\n' +
        '    shp = ws.Shapes(i)\n' +
        '    print("Tombol:", shp.Name, "-> Macro:", shp.OnAction)\n' +
        '# Baca isi macro (butuh Trust access to VBA project object model):\n' +
        'for comp in wb.VBProject.VBComponents:\n' +
        '    if comp.Type == 1:\n' +
        '        cm = comp.CodeModule\n' +
        '        if cm.CountOfLines > 0:\n' +
        '            kode = cm.Lines(1, cm.CountOfLines)\n' +
        '            if "NamaMacro" in kode:\n' +
        '                print("=== Modul", comp.Name, "===")\n' +
        '                print(kode)\n' +
        '```\n\n' +
        
        '### PENCOCOKAN NAMA SHEET ###\n' +
        'Nama sheet yang disebut user mungkin tidak persis sama dengan nama sheet di workbook.\n' +
        'Contoh: user bilang "Data Siswa" tapi sheet sebenarnya "Daftar Siswa".\n' +
        'WAJIB: Cocokkan nama sheet yang disebut user dengan daftar sheet di konteks.\n' +
        'Cari sheet yang paling mirip (abaikan spasi, huruf besar/kecil, kata tambahan).\n' +
        'Jika tidak ada yang cocok, tampilkan daftar sheet yang tersedia dan minta user memilih.\n' +
        'JANGAN mengarang nama sheet. JANGAN menggunakan nama sheet yang tidak ada di konteks.\n\n' +
        
        '### SUMBER DATA ###\n' +
        'WAJIB: Baca data dari sheet yang ada di workbook. JANGAN mengarang atau membuat data sendiri.\n' +
        'Jika user meminta data dari sheet tertentu, baca dari sheet tersebut menggunakan kode Python.\n' +
        'Gunakan wb.Worksheets("NamaSheet").Range(...) atau .Cells(...) untuk membaca data.\n' +
        'Jika sheet tidak ditemukan, tampilkan error yang jelas, JANGAN lanjut dengan data palsu.\n\n' +
        
        '### KONSTANTA EXCEL ###\n' +
        'Konstanta Excel sudah tersedia langsung di lingkungan eksekusi.\n' +
        'Gunakan langsung tanpa import: xlUp, xlDown, xlToLeft, xlToRight,\n' +
        'xlCenter, xlLeft, xlRight, xlTop, xlBottom, xlContinuous, xlThin,\n' +
        'xlMedium, xlThick, xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight,\n' +
        'xlInsideVertical, xlInsideHorizontal, xlSolid, xlAutomatic, xlNone,\n' +
        'xlValues, xlFormulas, xlLandscape, xlPortrait, xlPaperA4, xlAscending,\n' +
        'xlDescending, xlColorIndexNone, xlColorIndexAutomatic, dll.\n' +
        'JANGAN definisikan ulang konstanta ini. JANGAN gunakan angka mentah.\n\n' +
        
        '### CONTOH RESPONS YANG BENAR ###\n' +
        'User: "Atur lebar kolom A jadi 7, B jadi 16, C jadi 18"\n' +
        'AI: "Baik, saya atur lebar kolom sesuai permintaan."\n' +
        '```python\n' +
        'ws = wb.ActiveSheet\n' +
        'ws.Columns("A").ColumnWidth = 7\n' +
        'ws.Columns("B").ColumnWidth = 16\n' +
        'ws.Columns("C").ColumnWidth = 18\n' +
        '```\n\n' +
        
        '### CONTOH: BUAT DAFTAR HADIR ###\n' +
        'User: "Buat daftar hadir rapat orang tua kelas 4A"\n' +
        'AI: "Baik, saya buatkan daftar hadir rapat orang tua."\n' +
        '```python\n' +
        'ws = wb.ActiveSheet\n' +
        'ws.Cells.Clear()\n' +
        '# Judul\n' +
        'ws.Range("A1:E1").Merge()\n' +
        'ws.Range("A1").Value = "DAFTAR HADIR RAPAT ORANG TUA KELAS 4A"\n' +
        'ws.Range("A1").Font.Size = 16\n' +
        'ws.Range("A1").Font.Bold = True\n' +
        'ws.Range("A1").Font.Color = 0xFFFFFF\n' +
        'ws.Range("A1").Interior.Color = 0x1F4E79\n' +
        'ws.Range("A1").HorizontalAlignment = -4108\n' +
        'ws.Rows(1).RowHeight = 36\n' +
        '# Header\n' +
        'headers = ["No", "Nama Siswa", "Nama Orang Tua", "Tanda Tangan", "Keterangan"]\n' +
        'for i, h in enumerate(headers, 1):\n' +
        '    c = ws.Cells(2, i)\n' +
        '    c.Value = h\n' +
        '    c.Font.Bold = True\n' +
        '    c.Font.Color = 0xFFFFFF\n' +
        '    c.Interior.Color = 0x2E75B6\n' +
        '    c.HorizontalAlignment = -4108\n' +
        '# Nomor urut 1-20\n' +
        'for i in range(1, 21):\n' +
        '    c = ws.Cells(i + 2, 1)\n' +
        '    c.Value = i\n' +
        '    c.HorizontalAlignment = -4108\n' +
        '# Lebar kolom\n' +
        'ws.Columns("A").ColumnWidth = 6\n' +
        'ws.Columns("B").ColumnWidth = 22\n' +
        'ws.Columns("C").ColumnWidth = 22\n' +
        'ws.Columns("D").ColumnWidth = 16\n' +
        'ws.Columns("E").ColumnWidth = 18\n' +
        '```\n\n' +
        
        '### CONTOH RESPONS YANG SALAH (DILARANG) ###\n' +
        'User: "Atur lebar kolom A jadi 7"\n' +
        'AI: "Baik, berikut analisis lebar kolom: Kolom A sebaiknya 7 karena..." (TANPA KODE) = GAGAL\n\n';

    // Detect skill from user message
    var skillTriggers = {
        'format': ['format', 'style', 'font', 'color', 'fill', 'border', 'bold', 'italic', 'number format', 'currency', 'percent'],
        'column-width': ['column width', 'lebar kolom', 'width', 'auto fit', 'autofit'],
        'row-height': ['row height', 'tinggi baris', 'height'],
        'write': ['write', 'tulis', 'set value', 'input data', 'isi'],
        'read': ['read', 'baca', 'get value', 'ambil data'],
        'formula': ['formula', 'rumus', 'sum', 'average', 'count', 'function'],
        'table': ['table', 'tabel', 'create table', 'buat tabel'],
        'sort': ['sort', 'urutkan', 'ascending', 'descending'],
        'filter': ['filter', 'saring', 'auto filter'],
        'chart': ['chart', 'grafik', 'bar chart', 'pie chart', 'line chart', 'visualisasi'],
        'pivot': ['pivot', 'pivot table', 'tabel pivot', 'summary', 'ringkasan'],
        'conditional': ['conditional', 'highlight', 'warnai', 'color scale'],
        'named-range': ['named range', 'range name', 'definisikan range', 'nama range'],
        'sheet': ['sheet', 'worksheet', 'buat sheet', 'hapus sheet', 'rename sheet'],
        'protect': ['protect', 'unprotect', 'lock', 'unlock', 'proteksi'],
        'merge': ['merge', 'gabung', 'unmerge', 'pisah'],
        'delete': ['delete', 'hapus', 'remove row', 'remove column'],
        'insert': ['insert', 'sisipkan', 'add row', 'add column']
    };
    
    function detectSkill(message) {
        var msg = message.toLowerCase();
        for (var skill in skillTriggers) {
            for (var i = 0; i < skillTriggers[skill].length; i++) {
                if (msg.indexOf(skillTriggers[skill][i]) !== -1) {
                    return skill;
                }
            }
        }
        return null;
    }
    
    function getSkillContext(skill) {
        if (!skill) return '';
        
        var contexts = {
            'format': '\n\n### SKILL: FORMAT RANGE ###\nGunakan .NumberFormat untuk format angka:\n- Currency: ws.Range("C:C").NumberFormat = "#,##0"\n- Percent: ws.Range("D:D").NumberFormat = "0.00%"\n- Date: ws.Range("B:B").NumberFormat = "dd/mm/yyyy"\nGunakan .Interior.Color, .Font.Bold untuk styling.\n',
            'column-width': '\n\n### SKILL: COLUMN WIDTH ###\nGunakan .ColumnWidth untuk atur lebar kolom:\nws.Columns("A").ColumnWidth = 7\nAtau auto-fit: ws.Columns("A").AutoFit()\n',
            'row-height': '\n\n### SKILL: ROW HEIGHT ###\nGunakan .RowHeight untuk atur tinggi baris:\nws.Rows(1).RowHeight = 30\n',
            'formula': '\n\n### SKILL: FORMULA ###\nGunakan .Formula (bukan .Value) untuk formula:\nws.Range("E1").Formula = "=SUM(B1:D1)"\n',
            'table': '\n\n### SKILL: TABLE ###\nBuat tabel untuk data terstruktur:\nws.ListObjects.Add(1, ws.Range("A1:C10"), None, 1).Name = "NamaTabel"\n',
            'sort': '\n\n### SKILL: SORT ###\nws.Range("A1:C100").Sort(ws.Range("A1"), 1)\n',
            'chart': '\n\n### SKILL: CHART ###\nchart = ws.ChartObjects().Add(100, 100, 400, 300)\nchart.Chart.SetSourceData(ws.Range("A1:B10"))\nchart.Chart.ChartType = 51\n',
            'pivot': '\n\n### SKILL: PIVOT TABLE ###\n# Gunakan wb.Connections.Add2() atau data model untuk pivot table\n'
        };
        
        return contexts[skill] || '';
    }

    // Get rich workbook schema, then supplement with MCP metadata if available
    getWorkbookSchema(function(err, schema) {
        var ctxMsg;
        if (!err && schema) {
            ctxMsg = buildContextMessage(schema);
        } else {
            ctxMsg = "Gagal membaca konteks workbook: " + (err ? err.error || err.message : "unknown");
        }
        
        // Detect and add skill context
        var detectedSkill = detectSkill(message);
        if (detectedSkill) {
            ctxMsg += getSkillContext(detectedSkill);
        }
        
        // Add multi-step instruction
        ctxMsg += "\n\nJika tugas kompleks, pecah menjadi beberapa langkah kecil. " +
                  "Tulis SEMUA langkah dalam SATU respons, masing-masing sebagai blok ```python terpisah. " +
                  "Sistem akan mengeksekusi semuanya secara berurutan. " +
                  "JANGAN hanya menulis langkah pertama. JANGAN minta user menjalankan kode manual.";
        
        // Coba tambahkan metadata MCP jika tersedia (hanya file lokal)
        getMcpWorkbookInfo(function(mcpData) {
            if (mcpData) {
                try {
                    var mcpStr = typeof mcpData === 'string' ? mcpData : JSON.stringify(mcpData);
                    // Batasi ukuran agar tidak terlalu panjang
                    if (mcpStr.length > 3000) mcpStr = mcpStr.substring(0, 3000) + '...(terpotong)';
                    ctxMsg += "\n\n=== METADATA WORKBOOK (MCP) ===\n" + mcpStr;
                } catch (e) { /* abaikan jika gagal serialize */ }
            }
            runAgentTurn(systemPrompt, ctxMsg, 0);
        });
    });
    
    // Watchdog pengaman: hanya boleh menyala lebih lama dari seluruh pipeline
    // (streaming + eksekusi). Jangan memakai AGENT_TURN_TIMEOUT_MS yang lebih
    // pendek karena bisa melepas kunci isProcessing saat turn masih berjalan.
    setTimeout(function() {
        // Sudah dibatalkan user: jangan tampilkan pesan pengaman.
        if (cancelRequested) return;
        if (isProcessing) {
            document.getElementById('loadingIndicator').className = 'loading hidden';
            addMessage('system', 'Proses memakan waktu sangat lama dan dihentikan oleh pengaman. Silakan coba lagi.');
            isProcessing = false;
        }
    }, AGENT_WATCHDOG_MS);
}

var MAX_RETRY = 2;

// Extract all code blocks from AI response (multi-step support).
// HANYA blok ```python yang diambil: blok bahasa lain (mis. contoh VBA di
// dalam penjelasan) tidak boleh ikut dieksekusi sebagai Python.
function extractCodeBlocks(content) {
    var blocks = [];
    var regex = /```python[ \t]*\r?\n([\s\S]*?)```/gi;
    var match;

    while ((match = regex.exec(content)) !== null) {
        var code = match[1].replace(/^\s+|\s+$/g, '');
        if (code) blocks.push(code);
    }

    return blocks;
}

// Ambil snapshot SELURUH sheet untuk verifikasi perubahan.
// Versi lama hanya memotret sheet aktif, sehingga perubahan pada sheet lain
// (mis. agent menulis ke sheet "REKAP" saat sheet aktif "Daftar Hadir")
// dilaporkan sebagai "tidak ada perubahan". Sekarang setiap sheet dipotret:
// nama, alamat + dimensi used range, dan sampel nilainya.
function getSheetSnapshot(callback) {
    if (typeof Excel === 'undefined') { callback(null); return; }
    Excel.run(function(context) {
        var allSheets = context.workbook.worksheets;
        allSheets.load("name");
        var sheet = context.workbook.worksheets.getActiveWorksheet();
        sheet.load("name");
        var usedRange = sheet.getUsedRangeOrNullObject();
        usedRange.load("rowCount, columnCount, address");
        var sheetRanges = [];
        var sampleRanges = [];

        return context.sync().then(function() {
            // Cap jumlah sheet yang di-scan: workbook dengan puluhan sheet
            // bisa membuat Excel hang jika semua di-load sekaligus.
            var maxSheets = Math.min(allSheets.items.length, 15);
            for (var i = 0; i < maxSheets; i++) {
                var sr = allSheets.items[i].getUsedRangeOrNullObject();
                sr.load("rowCount, columnCount, address");
                sheetRanges.push(sr);
            }
            return context.sync();
        }).then(function() {
            // Ambil sampel nilai (pojok kiri-atas used range) tiap sheet.
            // Cap ukuran sampel: 10 baris x 10 kolom cukup untuk deteksi perubahan.
            for (var j = 0; j < sheetRanges.length; j++) {
                var srj = sheetRanges[j];
                if (srj && !srj.isNullObject && srj.rowCount > 0 && srj.columnCount > 0) {
                    var rows = Math.min(srj.rowCount, 10);
                    var cols = Math.min(srj.columnCount, 10);
                    var smp = srj.getCell(0, 0).getResizedRange(rows - 1, cols - 1);
                    smp.load("values");
                    sampleRanges.push(smp);
                } else {
                    sampleRanges.push(null);
                }
            }
            return context.sync();
        }).then(function() {
            var sheetCount = allSheets.items.length;
            var sheetNames = [];
            var perSheet = [];
            for (var k = 0; k < sheetCount; k++) {
                var name = allSheets.items[k].name;
                sheetNames.push(name);
                var srk = sheetRanges[k];
                if (srk && !srk.isNullObject) {
                    var sig = name + ':' + srk.address + ':' + srk.rowCount + 'x' + srk.columnCount;
                    if (sampleRanges[k] && sampleRanges[k].values) {
                        sig += ':' + JSON.stringify(sampleRanges[k].values);
                    }
                    perSheet.push(sig);
                } else {
                    perSheet.push(name + ':EMPTY');
                }
            }
            callback({
                empty: usedRange.isNullObject,
                sheetCount: sheetCount,
                sheetNames: sheetNames.join(","),
                allSheetsSig: perSheet.join("|"),
                rowCount: usedRange.isNullObject ? 0 : usedRange.rowCount,
                colCount: usedRange.isNullObject ? 0 : usedRange.columnCount,
                address: usedRange.isNullObject ? '' : usedRange.address
            });
            return null;
        });
    }).catch(function() { callback(null); });
}

// Verify sheet has data after execution
function verifySheetChanged(callback) {
    assertBindingCurrent(function(bindingError) {
        if (bindingError) {
            callback(false, bindingError.error);
            return;
        }
        getSheetSnapshot(function(currentSnapshot) {
            if (!currentSnapshot) {
                callback(false);
            } else if (!sheetSnapshotBefore) {
                callback(!currentSnapshot.empty);
            } else {
                // Perbandingan sidik jari seluruh sheet: jumlah sheet, nama-nama
                // sheet, dan isi (alamat+dimensi+sampel nilai) tiap sheet.
                callback(
                    currentSnapshot.sheetCount !== sheetSnapshotBefore.sheetCount ||
                    currentSnapshot.sheetNames !== sheetSnapshotBefore.sheetNames ||
                    currentSnapshot.allSheetsSig !== sheetSnapshotBefore.allSheetsSig
                );
            }
        });
    });
}

// Execute code with retry and context refresh
function executeCodeWithRetry(code, attempt, systemPrompt, contextMessage) {
    if (cancelRequested) { isProcessing = false; return; }
    assertBindingCurrent(function(bindingError) {
        if (bindingError) {
            addMessage('system', bindingError.error);
            document.getElementById('loadingIndicator').className = 'loading hidden';
            isProcessing = false;
            return;
        }
        executeBoundCodeWithRetry(code, attempt, systemPrompt, contextMessage);
    });
}

function executeBoundCodeWithRetry(code, attempt, systemPrompt, contextMessage) {
    if (cancelRequested) { isProcessing = false; return; }
    document.getElementById('loadingIndicator').className = 'loading';
    console.log("[Agent] Eksekusi kode (percobaan " + (attempt + 1) + "):", code);
    
    // Snapshot sheet state sebelum eksekusi
    getSheetSnapshot(function(snapshot) {
        sheetSnapshotBefore = snapshot;
        
        startProgressPolling();
        runPythonCode(code, function (result) {
            stopProgressPolling();
            document.getElementById('loadingIndicator').className = 'loading hidden';

            if (result.ok) {
                console.log("[Agent] Eksekusi berhasil! Verifikasi...");
                var execOutput = result.output || '';
                var hasOutput = execOutput.replace(/^\s+|\s+$/g, '') !== '';

                setTimeout(function() {
                    verifySheetChanged(function(changed, verifyError) {
                        if (verifyError) {
                            addMessage('system', verifyError);
                            isProcessing = false;
                            return;
                        }
                        // TUGAS ANALISIS: kode menghasilkan output print() tetapi tidak
                        // harus mengubah sheet. Umpankan output kembali ke AI supaya ia
                        // bisa menyusun jawaban nyata untuk user (bukan menyerah dan
                        // menyuruh user mengecek manual).
                        if (hasOutput && reportCycles < MAX_REPORT_CYCLES) {
                            reportCycles += 1;
                            expectingAnswer = true;
                            addMessage('system', 'Kode berhasil dijalankan. Meminta AI menyusun jawaban dari hasilnya...');
                            chatHistory.push({
                                role: 'user',
                                content: 'KODE Python yang Anda tulis SUDAH SELESAI dijalankan dan berhasil. ' +
                                    'Berikut output print() yang dihasilkan:\n```\n' + tailLines(execOutput, 60) + '\n```\n\n' +
                                    'Berdasarkan output di atas, tulis JAWABAN AKHIR untuk user dalam bahasa Indonesia yang jelas dan ringkas. ' +
                                    'JANGAN menulis kode Python lagi pada balasan ini. JANGAN menyuruh user mengecek manual.'
                            });
                            saveHistory();
                            runAgentTurn(systemPrompt, contextMessage, attempt);
                            return;
                        }
                        if (changed) {
                            addMessage('system', 'Berhasil dieksekusi dan data tersimpan di workbook.');
                        } else if (hasOutput) {
                            // Sudah melewati batas siklus laporan: tampilkan hasil langsung.
                            addMessage('system', 'Kode dieksekusi (hanya membaca data). Hasil:\n' + tailLines(execOutput, 30));
                        } else {
                            addMessage('system', 'Kode dieksekusi tapi tidak ada perubahan terdeteksi di sheet.');
                        }
                        saveHistory();
                        isProcessing = false;
                    });
                }, 500);
                return;
            }

            console.log("[Agent] Eksekusi gagal:", result.error);

            if (result.targetError) {
                addMessage('system', 'Target workbook tidak valid: ' + result.error);
                setBindingBlocked('Target workbook tidak valid. Buka ulang task pane untuk melanjutkan.');
                saveHistory();
                isProcessing = false;
                return;
            }

            if (attempt < MAX_RETRY) {
                addMessage('system', 'Eksekusi gagal, memperbaiki sendiri (percobaan ' +
                    (attempt + 2) + '/' + (MAX_RETRY + 1) + ')...');
                
                assertBindingCurrent(function(bindingError) {
                    if (bindingError) {
                        addMessage('system', bindingError.error);
                        isProcessing = false;
                        return;
                    }
                    getWorkbookSchema(function(err, schema) {
                        var newContext = contextMessage;
                        if (!err && schema) {
                            newContext = buildContextMessage(schema);
                        }

                        // Sertakan traceback ringkas dan output print() agar AI
                        // punya bahan diagnostik yang cukup untuk memperbaiki diri.
                        var fixRequest = 'Kode Python gagal dijalankan dengan error: ' + result.error + '\n';
                        if (result.traceback) {
                            fixRequest += 'Traceback (baris terakhir):\n' + tailLines(result.traceback, 6) + '\n';
                        }
                        if (result.output) {
                            fixRequest += 'Output print() sebelum error:\n' + tailLines(result.output, 10) + '\n';
                        }
                        fixRequest += 'Kode yang ditulis:\n```python\n' + code + '\n```\n\n' +
                            'PERBAIKI kode ini. Tulis ulang kode Python yang BENAR menggunakan win32com.client.';
                        chatHistory.push({
                            role: 'user',
                            content: fixRequest
                        });
                        saveHistory();
                        runAgentTurn(systemPrompt, newContext, attempt + 1);
                    });
                });
            } else {
                addMessage('system', 'Gagal setelah ' + (MAX_RETRY + 1) +
                    ' percobaan: ' + result.error);
                saveHistory();
                isProcessing = false;
            }
        });
    });
}

// Execute multiple code blocks sequentially (multi-step)
function executeMultiStep(codeBlocks, index, systemPrompt, contextMessage) {
    if (cancelRequested) { isProcessing = false; return; }
    if (index > codeBlocks.length) {
        assertBindingCurrent(function(bindingError) {
            if (bindingError) {
                addMessage('system', bindingError.error);
                isProcessing = false;
                return;
            }
            verifySheetChanged(function(changed, verifyError) {
                if (verifyError) {
                    addMessage('system', verifyError);
                    isProcessing = false;
                    return;
                }
                addMessage('system', changed ? 'Semua langkah selesai. Data berubah di workbook.' :
                    'Semua langkah selesai. Tidak ada perubahan terdeteksi di sheet.');
                saveHistory();
                isProcessing = false;
            });
        });
        return;
    }

    if (index === 0) {
        getSheetSnapshot(function(snapshot) {
            sheetSnapshotBefore = snapshot;
            // Gabungkan semua code blocks ke satu script agar variabel persist antar langkah
            var combinedCode = '';
            for (var i = 0; i < codeBlocks.length; i++) {
                combinedCode += '# === Langkah ' + (i + 1) + '/' + codeBlocks.length + ' ===\n';
                combinedCode += 'progress("Langkah ' + (i + 1) + '/' + codeBlocks.length + ' sedang dieksekusi...")\n';
                combinedCode += codeBlocks[i] + '\n\n';
            }
            addMessage('system', 'Menjalankan ' + codeBlocks.length + ' langkah dalam satu eksekusi...');
            executeCodeWithRetry(combinedCode, 0, systemPrompt, contextMessage);
        });
        return;
    }
    
    assertBindingCurrent(function(bindingError) {
        if (bindingError) {
            addMessage('system', bindingError.error);
            isProcessing = false;
            return;
        }
        addMessage('system', 'Langkah ' + index + '/' + codeBlocks.length + ' sedang dieksekusi...');

        startProgressPolling();
        runPythonCode(codeBlocks[index - 1], function (result) {
        stopProgressPolling();
        if (result.ok) {
            addMessage('system', 'Langkah ' + index + ' berhasil.');
            
            // Refresh context after each step
            getWorkbookSchema(function(err, schema) {
                executeMultiStep(codeBlocks, index + 1, systemPrompt, contextMessage);
            });
        } else {
            addMessage('system', 'Error di langkah ' + index + ': ' + result.error);

            if (result.targetError) {
                setBindingBlocked('Target workbook tidak valid. Buka ulang task pane untuk melanjutkan.');
                saveHistory();
                isProcessing = false;
                return;
            }
            
            assertBindingCurrent(function(bindingError) {
                if (bindingError) {
                    addMessage('system', bindingError.error);
                    isProcessing = false;
                    return;
                }

                // Send error to AI for fixing - minta regenerasi SEMUA langkah
                var fixMsg = 'Error di langkah ' + index + ' dari ' + codeBlocks.length + ': ' + result.error + '\n\n';
                fixMsg += 'Tulis ulang SEMUA ' + codeBlocks.length + ' langkah dengan perbaikan berdasarkan error di atas. ' +
                          'Gunakan nama sheet yang benar. Untuk tugas modifikasi, setiap langkah harus menghasilkan perubahan nyata pada workbook. ' +
                          'Konstanta Excel seperti xlUp, xlDown, xlToLeft, xlToRight, xlCenter sudah tersedia langsung. ' +
                          'Tulis semua langkah dalam satu respons, masing-masing sebagai blok ```python terpisah.';
                chatHistory.push({
                    role: 'user',
                    content: fixMsg
                });
                saveHistory();

                getWorkbookSchema(function(err, schema) {
                    var newContext = contextMessage;
                    if (!err && schema) {
                        newContext = buildContextMessage(schema);
                    }
                    runAgentTurn(systemPrompt, newContext, 0);
                });
            });
        }
        });
    });
}

// Ambil N baris terakhir dari sebuah teks (untuk meringkas traceback/output).
function tailLines(text, n) {
    if (typeof text !== 'string' || !text) return '';
    var lines = text.split('\n');
    var out = [];
    for (var i = Math.max(0, lines.length - n); i < lines.length; i++) out.push(lines[i]);
    return out.join('\n');
}

// Buat pesan assistant "hidup" yang teksnya diperbarui saat streaming berjalan.
function createStreamingMessage() {
    var chatMessages = document.getElementById('chatMessages');
    var messageDiv = document.createElement('div');
    messageDiv.className = 'message assistant';
    messageDiv.innerHTML =
        '<div class="message-avatar">\uD83E\uDD16</div>' +
        '<div class="message-content"></div>';
    chatMessages.appendChild(messageDiv);
    chatMessages.scrollTop = chatMessages.scrollHeight;
    return { div: messageDiv, contentEl: messageDiv.querySelector('.message-content') };
}

// Panggil API dengan mode streaming (SSE). Token demi token dirender langsung
// sehingga (1) user melihat progres, dan (2) koneksi tetap aktif selama model
// berpikir - ini menghilangkan "Gateway timeout" pada jawaban panjang.
function streamAgentTurn(apiEndpoint, payload, handlers) {
    var xhr = new XMLHttpRequest();
    trackXhr(xhr);
    xhr.open('POST', SERVER + '/api/proxy', true);
    xhr.setRequestHeader('Content-Type', 'application/json');
    xhr.timeout = STREAM_TIMEOUT_MS;

    var parsedIndex = 0;
    var sseBuffer = '';
    var fullContent = '';
    var finished = false;

    function finish(ok, status, errData) {
        if (finished) return;
        finished = true;
        untrackXhr(xhr);
        // Dibatalkan user: jangan panggil onError agar tidak memicu fallback
        // non-streaming yang akan menghidupkan kembali giliran.
        if (cancelRequested) return;
        if (ok) handlers.onDone(fullContent);
        else handlers.onError(status, errData);
    }

    xhr.onabort = function () {
        finished = true;
        untrackXhr(xhr);
    };

    function processEvents(events) {
        for (var i = 0; i < events.length; i++) {
            var lines = events[i].split('\n');
            for (var j = 0; j < lines.length; j++) {
                var line = lines[j].replace(/^\s+|\s+$/g, '');
                if (line.indexOf('data:') !== 0) continue;
                var data = line.substring(5).replace(/^\s+/, '');
                if (!data || data === '[DONE]') continue;
                try {
                    var obj = JSON.parse(data);
                    if (obj && obj.error) {
                        finish(false, 502, obj);
                        return;
                    }
                    var choice = obj && obj.choices && obj.choices[0];
                    var delta = choice && (choice.delta || choice.message);
                    var piece = delta && typeof delta.content === 'string' ? delta.content : '';
                    if (piece) {
                        fullContent += piece;
                        if (handlers.onDelta) handlers.onDelta(fullContent);
                    }
                } catch (e) { /* potongan JSON tidak utuh, abaikan */ }
            }
        }
    }

    function drainNewText() {
        var text = xhr.responseText || '';
        if (text.length <= parsedIndex) return;
        var chunk = text.substring(parsedIndex);
        parsedIndex = text.length;
        sseBuffer += chunk.replace(/\r\n/g, '\n');
        var parts = sseBuffer.split('\n\n');
        sseBuffer = parts.pop();
        if (parts.length) processEvents(parts);
    }

    xhr.onprogress = function() { drainNewText(); };

    xhr.onload = function() {
        try { drainNewText(); } catch (e) { }
        if (sseBuffer.replace(/^\s+|\s+$/g, '')) {
            processEvents([sseBuffer]);
            sseBuffer = '';
        }
        if (xhr.status >= 200 && xhr.status < 300) {
            if (fullContent) { finish(true); return; }
            // Gateway mungkin mengabaikan stream=true dan membalas JSON utuh.
            var parsedFull = null;
            try { parsedFull = JSON.parse(xhr.responseText); } catch (e) { }
            if (parsedFull && parsedFull.choices && parsedFull.choices.length > 0 &&
                parsedFull.choices[0].message && parsedFull.choices[0].message.content) {
                fullContent = parsedFull.choices[0].message.content;
                finish(true);
                return;
            }
            finish(false, xhr.status, parsedFull || { error: { message: 'Respons streaming kosong.' } });
            return;
        }
        var parsedErr = null;
        try { parsedErr = JSON.parse(xhr.responseText); } catch (e) { }
        finish(false, xhr.status, parsedErr);
    };

    xhr.onerror = function() { finish(false, 0, { error: { message: 'Network error' } }); };
    xhr.ontimeout = function() {
        finish(false, 0, { error: { message: 'Request timeout setelah ' + Math.round(STREAM_TIMEOUT_MS / 1000) + ' detik.' } });
    };

    var reqPayload = {};
    for (var k in payload) {
        if (payload.hasOwnProperty(k)) reqPayload[k] = payload[k];
    }
    reqPayload.stream = true;

    xhr.send(JSON.stringify({ apiEndpoint: apiEndpoint, payload: reqPayload }));
    return xhr;
}

// Proses balasan AI (dipakai jalur streaming maupun non-streaming).
// ---- Deteksi intensi: aksi (butuh kode) vs tanya-cukup-jelaskan ----
// Pertanyaan penjelasan ("apa bedanya", "kenapa", "jelaskan") sah dijawab
// tanpa kode; jangan dipaksa retry yang membuat agent terasa stuck.
var ACTION_HINTS = /(buat|ubah|ganti|isi|hapus|tulis|rapi|format|rapikan|masukkan|input|salin|pindahkan|urutkan|sortir|hitungkan|rekap|jumlahkan|gabungkan|rapikan|pasang|terapkan|jalankan|eksekusi|hapus duplikat|cari dan|tambah baris|tambah sheet|buat chart|buat grafik|buat rumus|isilah|ubahlah|buatkan|rapikan|koreksi|perbaiki|lengkapi|generate|create|write|modify|change|delete|insert|update|format|apply|run|execute|edit)\b/i;
var EXPLAIN_HINTS = /(apa itu|apa bedanya|apa fungsi|apa yang|mengapa|kenapa|jelaskan|apa maksud|bagaimana cara|gimana caranya|cara kerja|berikan penjelasan|tolong jelaskan|apa perbedaan|apakah bisa|bisakah|mungkinkah|saya ingin tahu|what is|what's the difference|explain|why|how does|what does)\b/i;

// Kembalikan true bila pesan user terakhir terlihat meminta AKSI pada workbook.
function looksLikeActionRequest() {
    var lastUser = '';
    for (var i = chatHistory.length - 1; i >= 0; i--) {
        if (chatHistory[i].role === 'user') { lastUser = chatHistory[i].content || ''; break; }
    }
    if (!lastUser) return true; // tanpa konteks: anggap butuh aksi (perilaku lama)
    // Tanya-cukup-jelaskan menang telak: "jelaskan cara membuat pivot" = tanya.
    if (EXPLAIN_HINTS.test(lastUser)) return false;
    return ACTION_HINTS.test(lastUser);
}
// Ekspos ke global scope untuk pengujian unit (node vm context).
try { if (typeof window !== 'undefined') window.looksLikeActionRequest = looksLikeActionRequest; } catch (e) { }

function handleAgentResponse(content, systemPrompt, contextMessage, liveMessage) {
    if (cancelRequested) { isProcessing = false; return; }
    document.getElementById('loadingIndicator').className = 'loading hidden';

    var wasExpectingAnswer = expectingAnswer;
    expectingAnswer = false;

    var codeBlocks = extractCodeBlocks(content);

    // Blok ```memory = agent menyimpan pelajaran ke memori jangka panjang.
    // Disimpan ke server, BUKAN dieksekusi sebagai kode.
    var memoryBlocks = extractMemoryBlocks(content);
    if (memoryBlocks.length > 0) {
        var savedCount = 0;
        memoryBlocks.forEach(function (block, idx) {
            saveMemoryBlock(block, function (ok) {
                savedCount += 1;
                if (ok) {
                    addMessage('system', 'Dicatat di memori [' + block.kategori + ']: ' + block.isi);
                }
                if (savedCount >= memoryBlocks.length) renderMemoriList();
            });
        });
    }

    chatHistory.push({ role: 'assistant', content: content });

    if (liveMessage) {
        addMessage('assistant', content, codeBlocks.length > 0 ? codeBlocks[0] : null, true, liveMessage.div);
    } else {
        addMessage('assistant', content, codeBlocks.length > 0 ? codeBlocks[0] : null);
    }
    saveHistory();

    if (codeBlocks.length === 0) {
        if (wasExpectingAnswer) {
            // Balasan tanpa kode adalah HASIL YANG BENAR untuk tugas analisis:
            // AI sedang memberikan jawaban akhir berdasarkan output eksekusi.
            console.log("[Agent] Jawaban akhir diterima (tugas analisis).");
            isProcessing = false;
            return;
        }
        // Pertanyaan penjelasan ("apa bedanya...", "kenapa...") sah dijawab
        // tanpa kode. Dulu kasus ini ikut dipaksa retry 2x sehingga agent
        // terasa stuck; sekarang jawaban teks langsung diterima.
        if (!looksLikeActionRequest()) {
            console.log("[Agent] Balasan tanpa kode untuk pertanyaan non-aksi - diterima.");
            isProcessing = false;
            return;
        }
        // AI menjawab tanpa kode untuk tugas yang butuh aksi. Model lemah sering
        // melakukan ini - paksa dengan pesan koreksi sebelum menyerah.
        if (noCodeRetries < MAX_NOCODE_RETRY) {
            noCodeRetries += 1;
            console.log("[Agent] Balasan tanpa kode, memaksa percobaan " + noCodeRetries);
            addMessage('system', 'AI menjawab tanpa kode - meminta ulang dengan penekanan...');
            chatHistory.push({
                role: 'user',
                content: 'ANDA BELUM MENULIS KODE PYTHON. Itu melanggar aturan utama Anda.\n' +
                    'Ingat: Anda adalah AGENT EKSEKUSI dengan AKSES PENUH ke workbook ini via Python win32com. ' +
                    'Anda TIDAK PERLU melihat file secara manual dan TIDAK BOLEH menebak atau menyuruh user mengecek sendiri.\n' +
                    'Tulis SEKARANG kode Python di dalam tag ```python untuk menangani permintaan user. ' +
                    'Sistem akan menjalankannya otomatis dan mengembalikan hasilnya (termasuk output print()) kepada Anda.\n' +
                    'Contoh untuk memeriksa tombol/macro:\n' +
                    '```python\n' +
                    'ws = wb.Worksheets("NamaSheet")\n' +
                    'for i in range(1, ws.Shapes.Count + 1):\n' +
                    '    shp = ws.Shapes(i)\n' +
                    '    print("Tombol:", shp.Name, "-> Macro:", shp.OnAction)\n' +
                    '```'
            });
            saveHistory();
            runAgentTurn(systemPrompt, contextMessage, attempt);
            return;
        }
        console.log("[Agent] Tidak ada kode yang dieksekusi. Response:", content);
        addMessage('system', 'Tidak ada kode yang dieksekusi. AI tidak menghasilkan kode Python.');
        isProcessing = false;
        return;
    }

    console.log("[Agent] Kode yang akan dieksekusi:", codeBlocks);
    addMessage('system', 'Menjalankan ' + codeBlocks.length + ' blok kode...');

    if (codeBlocks.length === 1) {
        executeCodeWithRetry(codeBlocks[0], 0, systemPrompt, contextMessage);
    } else {
        executeMultiStep(codeBlocks, 0, systemPrompt, contextMessage);
    }
}

// --- KOLABORASI OPENCODE ---
// OpenCode (CLI agent dengan penalaran mendalam) dipakai sebagai "otak":
// ia menganalisis dan menulis kode win32com; eksekusi tetap dilakukan oleh
// executor Python yang sudah ada (variabel `wb` disuntikkan di sana).
var OPENCODE_BRIEF =
    'Anda adalah OpenCode, AI agent dengan penalaran mendalam, yang sedang berkolaborasi dengan Excel Agent (add-in Excel).\n' +
    'Kode Python yang Anda tulis akan DIEKSEKUSI OTOMATIS terhadap workbook Excel yang SEDANG TERBUKA di komputer user.\n' +
    'Lingkungan eksekusi sudah menyediakan: variabel `wb` (workbook target), `xl` (aplikasi Excel), fungsi progress(msg), ' +
    'dan konstanta Excel (xlUp, xlDown, xlCenter, xlContinuous, xlThin, dll). JANGAN mendefinisikan ulang konstanta itu.\n\n' +
    'ATURAN MUTLAK:\n' +
    '1. JANGAN gunakan GetActiveObject, Dispatch, atau Workbooks.Open; gunakan `wb` yang sudah disediakan.\n' +
    '2. JANGAN memakai tool apa pun (jangan membaca/menulis/mengedit file, jangan menjalankan perintah, ' +
    'jangan menjelajah folder). Working directory Anda kosong dan memang tidak perlu dijelajahi. ' +
    'Semua informasi yang Anda butuhkan sudah ada di konteks workbook di bawah. Cukup hasilkan analisis + kode sebagai teks jawaban.\n' +
    '3. Semua kode Python WAJIB dibungkus tag ```python ... ```.\n' +
    '4. Untuk tugas MODIFIKASI: tulis kode lengkap yang menghasilkan perubahan nyata pada workbook.\n' +
    '5. Untuk tugas ANALISIS/PERTANYAAN: tulis kode yang membaca data lalu print() semua temuan. ' +
    'Sistem akan mengembalikan output print() kepada Anda; setelah itu berikan JAWABAN AKHIR tanpa kode.\n' +
    '6. Analisis secara mendalam terlebih dahulu sebelum menulis kode. JANGAN mengarang nama sheet/kolom; ' +
    'gunakan konteks workbook yang diberikan di bawah ini.\n';

// Folder tempat workbook berada (dipakai OpenCode sebagai working directory).
function opencodeWorkdir() {
    if (workbookBinding && workbookBinding.workbookPath) {
        var wp = workbookBinding.workbookPath;
        if (/^https?:\/\//i.test(wp)) return '';
        var idx = wp.lastIndexOf('\\');
        if (idx < 0) idx = wp.lastIndexOf('/');
        if (idx > 0) return wp.substring(0, idx);
    }
    return '';
}

// Satu putaran agent melalui OpenCode. Jika sudah ada session aktif, hanya
// pesan terakhir yang dikirim (riwayat sudah dimiliki OpenCode); jika belum,
// kirim brief + konteks + ringkasan riwayat sebagai session baru.
// Ambil setting terbaru dari server sebelum giliran OpenCode, supaya
// penggantian model OpenCode langsung berlaku tanpa buka-tutup task pane.
function refreshSettingsFromServer(done) {
    var xhr = new XMLHttpRequest();
    xhr.open('GET', SERVER + '/api/settings', true);
    xhr.onreadystatechange = function() {
        if (xhr.readyState !== 4) return;
        if (xhr.status === 200) {
            try {
                var cfg = JSON.parse(xhr.responseText);
                if (typeof cfg.opencodeModel === 'string' && cfg.opencodeModel) settings.opencodeModel = cfg.opencodeModel;
                if (typeof cfg.useOpencode === 'boolean') settings.useOpencode = cfg.useOpencode;
            } catch (e) { }
        }
        if (done) done();
    };
    xhr.onerror = function() { if (done) done(); };
    xhr.send();
}

// Kirim request ke /api/opencode. Bila server masih sibuk (429), coba ulang
// otomatis satu kali setelah 20 detik.
function sendOpencodeRequest(payload, onSuccess, onFail, retried429) {
    xhrPost(SERVER + '/api/opencode', payload, function(status, data) {
        if (status === 429 && !retried429) {
            addMessage('system', 'OpenCode masih memproses permintaan sebelumnya - mencoba ulang otomatis dalam 20 detik...');
            setTimeout(function() { sendOpencodeRequest(payload, onSuccess, onFail, true); }, 20000);
            return;
        }
        if (status >= 200 && status < 300 && data && data.ok) onSuccess(data);
        else onFail(status, data);
    }, OPENCODE_CLIENT_TIMEOUT_MS);
}

function runOpencodeAgentTurn(systemPrompt, contextMessage, attempt) {
    if (cancelRequested) { isProcessing = false; return; }
    refreshSettingsFromServer(function() {
        runOpencodeAgentTurnInner(systemPrompt, contextMessage, attempt);
    });
}

// Satu putaran agent melalui OpenCode. Jika sudah ada session aktif, hanya
// pesan terakhir yang dikirim (riwayat sudah dimiliki OpenCode); jika belum,
// kirim brief + konteks + ringkasan riwayat sebagai session baru.
function runOpencodeAgentTurnInner(systemPrompt, contextMessage, attempt) {
    document.getElementById('loadingIndicator').className = 'loading';

    var continuing = opencodeSessionId !== null;
    var prompt = '';
    if (continuing) {
        var lastMsg = chatHistory.length > 0 ? chatHistory[chatHistory.length - 1] : null;
        prompt = lastMsg ? lastMsg.content : '';
        if (!prompt || !prompt.replace(/^\s+|\s+$/g, '')) continuing = false;
    }
    if (!continuing) {
        var historySlice = chatHistory.slice(-12);
        var convo = '';
        for (var i = 0; i < historySlice.length; i++) {
            convo += (historySlice[i].role === 'user' ? 'USER: ' : 'ASSISTANT: ') + historySlice[i].content + '\n\n';
        }
        prompt = OPENCODE_BRIEF + '\n' + (contextMessage || '') +
            '\n\n=== RIWAYAT PERCAKAPAN ===\n' + convo +
            'Tanggapi permintaan terakhir dari USER pada riwayat di atas.';
    }

    addMessage('system', 'OpenCode sedang menganalisis secara mendalam (bisa beberapa menit)...');
    startProgressPolling();

    var payload = {
        prompt: prompt,
        model: settings.opencodeModel,
        session: continuing ? opencodeSessionId : '',
        workdir: opencodeWorkdir(),
        target: workbookBinding,
        images: activeTurnImages.map(function(image) {
            return { name: image.name, type: image.type, data: image.data };
        })
    };

    sendOpencodeRequest(payload, function(data) {
        stopProgressPolling();
        activeTurnImages = [];
        if (data.sessionReset === true) opencodeSessionId = null;
        if (data.sessionId) opencodeSessionId = data.sessionId;
        // Jawaban OpenCode diproses pipeline yang sama: ekstrak blok
        // ```python, eksekusi, umpan-balik output, dst.
        handleAgentResponse(data.text, systemPrompt, contextMessage, null);
    }, function(status, data) {
        stopProgressPolling();
        activeTurnImages = [];
        document.getElementById('loadingIndicator').className = 'loading hidden';
        var errMsg = formatApiError(status, data);
        console.log('[Agent] OpenCode error - status:', status, 'data:', JSON.stringify(data).substring(0, 500));
        addMessage('assistant', 'Error OpenCode: ' + errMsg);
        isProcessing = false;
        updateComposerState();
    }, false);
}

// Satu putaran agent: minta kode ke AI, jalankan, retry otomatis bila error.
function runAgentTurn(systemPrompt, contextMessage, attempt) {
    if (cancelRequested) { isProcessing = false; return; }
    if (settings.useOpencode === true) {
        runOpencodeAgentTurn(systemPrompt, contextMessage, attempt);
        return;
    }
    var historySlice = chatHistory.slice(-20); // Increased from 12 to 20
    var messages = [{ role: 'system', content: systemPrompt }];
    
    // Inject context invisibly just before sending history
    if (contextMessage) {
        messages.push({ role: 'system', content: contextMessage });
    }

    // Add reasoning instruction
    messages.push({
        role: 'system',
        content: 'Jika tugas memerlukan beberapa langkah, tulis SEMUA langkah dalam SATU respons dengan format:\n' +
                 '**Langkah 1: [Judul]**\n[Penjelasan]\n```python\n[Kode]\n```\n\n' +
                 '**Langkah 2: [Judul]**\n[Penjelasan]\n```python\n[Kode]\n```\n\n' +
                 'JANGAN hanya menulis langkah pertama. JANGAN minta user menjalankan kode manual. ' +
                 'Semua blok kode dieksekusi dalam SATU sesi Python, variabel dari blok sebelumnya tersedia di blok berikutnya. ' +
                 'Gunakan ini untuk menghindari membaca ulang data yang sudah dibaca.'
    });

    // Tambahkan lampiran gambar jika ada (vision support - multimodal)
    // Format OpenAI compatible: [{type: "image_url", image_url: {url: "data:image/png;base64,..."}}, ...]
    if (activeTurnImages.length > 0) {
        // Image blocks sebagai bagian terakhir dari prompt user
        var recentUserIdx = -1;
        for (var i = messages.length - 1; i >= 0 && messages[i].role === 'user'; i--) {
            recentUserIdx = i;
            break;
        }
        
        if (recentUserIdx >= 0 && messages[recentUserIdx].content) {
            // Append images ke pesan user terakhir
            var msgObj = messages[recentUserIdx];
            msgObj.content = msgObj.content || '';
            
            // Check jika content string atau array
            if (typeof msgObj.content === 'string') {
                // Convert ke array untuk multimodal format
                msgObj.content = [msgObj.content];
            }
            
            // Tambahkan semua images ke array
            activeTurnImages.forEach(function(img) {
                msgObj.content.push({
                    type: "image_url",
                    image_url: {
                        url: img.data  // base64 encoded image
                    }
                });
            });
            messages[recentUserIdx] = msgObj;
        } else {
            // Tidak ada user message last — skip gambar (shouldn't happen normally)
            console.warn('[Agent] No user message found to attach images to');
        }
    }

    for (var i = 0; i < historySlice.length; i++) {
        messages.push(historySlice[i]);
    }

    document.getElementById('loadingIndicator').className = 'loading';

    var resolvedEndpoint = normalizeEndpoint(settings.apiEndpoint);
    var payload = {
        model: settings.model,
        messages: messages,
        // Suhu rendah: generasi KODE harus deterministik dan konsisten,
        // bukan kreatif. Ini signifikan mengurangi kode yang error.
        temperature: 0.2,
        max_tokens: 8192
    };

    // Jalur utama: streaming. Saat token mengalir, proxy tidak pernah idle
    // sehingga tidak kena "Gateway timeout" meskipun jawaban panjang.
    var liveMessage = createStreamingMessage();
    streamAgentTurn(resolvedEndpoint, payload, {
        onDelta: function(textSoFar) {
            if (liveMessage && liveMessage.contentEl) {
                liveMessage.contentEl.innerHTML = formatMessage(textSoFar);
                var chatMessages = document.getElementById('chatMessages');
                chatMessages.scrollTop = chatMessages.scrollHeight;
            }
        },
        onDone: function(content) {
            handleAgentResponse(content, systemPrompt, contextMessage, liveMessage);
        },
        onError: function(status, data) {
            // Fallback satu kali ke mode non-streaming (untuk gateway yang
            // tidak mendukung SSE) sebelum menyatakan gagal.
            console.log("[Agent] Streaming gagal (status " + status + "), fallback non-streaming");
            var fallbackPayload = {};
            for (var k in payload) {
                if (payload.hasOwnProperty(k)) fallbackPayload[k] = payload[k];
            }
            fallbackPayload.stream = false;
            xhrPost('https://localhost:3000/api/proxy', {
                apiEndpoint: resolvedEndpoint,
                payload: fallbackPayload
            }, function (status2, data2) {
                if (status2 >= 200 && status2 < 300 && data2.choices && data2.choices.length > 0) {
                    handleAgentResponse(data2.choices[0].message.content, systemPrompt, contextMessage, liveMessage);
                    return;
                }
                document.getElementById('loadingIndicator').className = 'loading hidden';
                var errMsg = formatApiError(status2, data2);
                console.log("[Agent] AI error - status:", status2, "data:", JSON.stringify(data2).substring(0, 500));
                if (liveMessage && liveMessage.contentEl) {
                    liveMessage.contentEl.innerHTML = formatMessage('Error: ' + errMsg);
                } else {
                    addMessage('assistant', 'Error: ' + errMsg);
                }
                isProcessing = false;
            }, AGENT_TURN_TIMEOUT_MS);
        }
    });
}

// Add message to chat UI. Jika targetDiv diberikan (mode streaming), pesan
// dirender ke div yang sudah ada di layar, bukan membuat div baru.
function addMessage(role, content, code, skipSave, targetDiv, images) {
    var chatMessages = document.getElementById('chatMessages');

    var messageDiv = targetDiv || document.createElement('div');
    messageDiv.className = 'message ' + role;

    var avatar = role === 'user' ? '\uD83D\uDC64' : '\uD83E\uDD16';

    var codeBlock = '';
    var imageBlock = '';
    if (images && images.length) {
        imageBlock = '<div class="message-images">';
        for (var imageIndex = 0; imageIndex < images.length; imageIndex++) {
            imageBlock += '<img src="' + escapeHtml(images[imageIndex].data) + '" alt="' + escapeHtml(images[imageIndex].name) + '">';
        }
        imageBlock += '</div>';
    }

    // Sembunyikan kode jika autoRun aktif (agar AI terlihat seperti "manusia" yang langsung bekerja)
    if (code && settings.autoRun) {
        // Hapus blok kode JS dari text content agar tidak muncul ganda
        content = content.replace(/```(?:javascript|js)[\s\S]*?```/ig, '');
        // Jangan render code block UI
        codeBlock = '';
    } else if (code) {
        codeBlock =
            '<div class="code-block">' +
            '<div class="code-block-header">' +
            '<span>Python</span>' +
            '<div class="code-block-btns">' +
            '<button class="copy-btn" onclick="copyCode(this)">Salin</button>' +
            '<button class="run-btn" onclick="executeJS(this)">Jalankan</button>' +
            '</div>' +
            '</div>' +
            '<pre><code>' + escapeHtml(code) + '</code></pre>' +
            '</div>';
    }

    messageDiv.innerHTML =
        '<div class="message-avatar">' + avatar + '</div>' +
        '<div class="message-content">' +
        imageBlock +
        formatMessage(content) +
        codeBlock +
        '</div>';

    if (!targetDiv) {
        chatMessages.appendChild(messageDiv);
    }
    chatMessages.scrollTop = chatMessages.scrollHeight;
}

// Format message (basic markdown)
function formatMessage(text) {
    // Remove code blocks from formatting
    var parts = text.split(/```[\s\S]*?```/g);
    var result = '';
    for (var i = 0; i < parts.length; i++) {
        var part = parts[i];
        part = part.replace(/`([^`]+)`/g, '<code>$1</code>');
        part = part.replace(/\*\*([^*]+)\*\*/g, '<strong>$1</strong>');
        part = part.replace(/\n/g, '<br>');
        result += part;
    }
    return result;
}

function escapeHtml(text) {
    var div = document.createElement('div');
    if (div.textContent !== undefined) {
        div.textContent = text;
    } else {
        div.innerText = text;
    }
    return div.innerHTML;
}

function copyCode(btn) {
    var code = btn.parentNode.parentNode.querySelector('code');
    var text = code.textContent || code.innerText;

    if (window.clipboardData) {
        window.clipboardData.setData('Text', text);
        btn.textContent = 'Tersalin!';
    } else {
        var textarea = document.createElement('textarea');
        textarea.value = text;
        document.body.appendChild(textarea);
        textarea.select();
        try {
            document.execCommand('copy');
            btn.textContent = 'Tersalin!';
        } catch (e) {
            btn.textContent = 'Gagal';
        }
        document.body.removeChild(textarea);
    }
    setTimeout(function () { btn.textContent = 'Salin'; }, 2000);
}

// Eksekusi Python code via server (win32com COM)

// Eksekusi Python code via server (win32com COM)
function runPythonCode(code, callback) {
    // Auto-recovery jika binding hilang - jangan langsung fail!
    if (!workbookBinding || !Office || !Office.context || !Office.context.document) {
        console.log('[ExcelAgent] Auto-recovery initiated: waiting for Excel context...');
        
        var attempt = 0;
        function tryRecover() {
            attempt++;
            
            if (!Office || !Office.context || !Office.context.document) {
                if (attempt < 10) {
                    console.log('[ExcelAgent] Waiting for Excel context...', attempt);
                    setTimeout(tryRecover, 500);
                    return;
                }
                callback({ ok: false, targetError: true, error: 'Excel tidak responsif. Tunggu 5-10 detik lalu coba lagi.' });
                return;
            }
            
            // Create new binding from current Excel
            var newPath = null;
            try {
                Office.context.document.getFilePropertiesAsync(function(result) {
                    if (result.status === Office.AsyncResultStatus.Succeeded && result.value && result.value.url) {
                        newPath = normalizeFileUrl(result.value.url);
                    }
                });
            } catch(e) {
                console.warn('[ExcelAgent] Failed to get path:', e);
            }
            
            if (!newPath || newPath === 'Unknown') {
                newPath = '(Active)-' + Date.now();
            }
            
            workbookBinding = {
                sessionId: createSessionId(),
                workbookName: newPath.split(/[\\\/]/).pop() || newPath,
                workbookPath: newPath
            };
            
            console.log('[ExcelAgent] ✅ Auto-binding created:', workbookBinding);
            
            // Now execute
            xhrPost(SERVER + '/api/exec-python', { code: code, target: workbookBinding }, function (status, data) {
                if (status >= 200 && status < 300 && data && data.ok) {
                    callback({ ok: true, message: data.message || 'OK', output: data.output || '' });
                } else {
                    // Dibatalkan user (ESC): tandai agar pemanggil tidak retry / lapor error.
                    if (cancelRequested || (data && data.cancelled === true)) {
                        callback({ ok: false, cancelled: true, error: 'Eksekusi dihentikan oleh pengguna.' });
                        return;
                    }
                    var errMsg = (data && data.error) ? data.error : ('HTTP ' + status);
                    callback({
                        ok: false,
                        targetError: !!(data && data.targetError === true) || status === 409,
                        error: errMsg,
                        traceback: (data && data.traceback) || '',
                        output: (data && data.output) || ''
                    });
                }
            });
        }
        
        tryRecover();
        return;
    }
    
    if (bindingBlocked) {
        callback({ ok: false, targetError: true, error: 'Target workbook tidak valid. Refresh task pane.' });
        return;
    }
    
    xhrPost(SERVER + '/api/exec-python', { code: code, target: workbookBinding }, function (status, data) {
        if (status >= 200 && status < 300 && data && data.ok) {
            callback({ ok: true, message: data.message || 'OK', output: data.output || '' });
        } else {
            // Dibatalkan user (ESC): tandai agar pemanggil tidak retry / lapor error.
            if (cancelRequested || (data && data.cancelled === true)) {
                callback({ ok: false, cancelled: true, error: 'Eksekusi dihentikan oleh pengguna.' });
                return;
            }
            var errMsg = (data && data.error) ? data.error : ('HTTP ' + status);
            callback({
                ok: false,
                targetError: !!(data && data.targetError === true) || status === 409,
                error: errMsg,
                traceback: (data && data.traceback) || '',
                output: (data && data.output) || ''
            });
        }
    });
}

// Eksekusi JS API langsung dalam add-in (Office.js fallback)
function runJS(code, callback) {
    callback({ ok: false, error: 'Eksekusi JavaScript tidak didukung. Gunakan executor Python.' });
}

// Tombol "Jalankan" manual pada code block
function executeJS(btn) {
    var codeBlock = btn.parentNode.parentNode.parentNode;
    var codeEl = codeBlock.querySelector('code');
    var code = codeEl.textContent || codeEl.innerText;

    btn.disabled = true;
    btn.textContent = 'Menjalankan...';

    assertBindingCurrent(function(bindingError) {
        if (bindingError) {
            btn.disabled = false;
            btn.textContent = 'Gagal';
            addMessage('system', bindingError.error);
            setTimeout(function () { btn.textContent = 'Jalankan'; }, 3000);
            return;
        }
        runPythonCode(code, function (result) {
            if (result.ok) {
                assertBindingCurrent(function(postBindingError) {
                    if (postBindingError) {
                        btn.textContent = 'Gagal';
                        addMessage('system', 'Gagal: ' + postBindingError.error);
                        setTimeout(function () { btn.textContent = 'Jalankan'; }, 3000);
                        return;
                    }
                    btn.disabled = false;
                    btn.textContent = 'Berhasil';
                    addMessage('system', 'Berhasil dieksekusi.');
                    setTimeout(function () { btn.textContent = 'Jalankan'; }, 3000);
                });
                return;
            } else {
                btn.textContent = 'Gagal';
                addMessage('system', 'Gagal: ' + result.error);
                if (result.targetError) {
                    setBindingBlocked('Target workbook tidak valid. Buka ulang task pane untuk melanjutkan.');
                    return;
                }
                btn.disabled = false;
            }
            setTimeout(function () { btn.textContent = 'Jalankan'; }, 3000);
        });
    });
}
