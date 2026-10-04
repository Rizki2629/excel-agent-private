const https = require('https');
const http = require('http');
const fs = require('fs');
const os = require('os');
const path = require('path');
const crypto = require('crypto');
const devCerts = require('office-addin-dev-certs');
const { getExcelMcpClient } = require('./mcp_client');
const hermesProviders = require('./hermes_providers');
const anthropic = require('./anthropic_provider');

const PORT = 3000;
const SERVER_HOST = '127.0.0.1';
const ALLOWED_ORIGIN = `https://localhost:${PORT}`;
const MAX_REQUEST_BODY_BYTES = 1024 * 1024;
const MAX_PENDING_COM = 20;
const PROXY_TIMEOUT_MS = 180000;
const EXEC_TIMEOUT_MS = 210000;
const OPENCODE_TIMEOUT_MS = 420000;
const MAX_OPENCODE_IMAGES = 5;
const MAX_IMAGE_BYTES = 5 * 1024 * 1024;
const MAX_OPENCODE_BODY_BYTES = 36 * 1024 * 1024;
const IMAGE_MIME_EXT = { 'image/png': '.png', 'image/jpeg': '.jpg', 'image/webp': '.webp' };
let comQueue = Promise.resolve();
let pendingCom = 0;
let opencodeRunning = false;

// Registry proses eksekusi yang sedang berjalan, dikunci per workbook.
// Dipakai endpoint /api/cancel untuk menghentikan eksekusi Python/VBA
// yang sedang berlangsung saat user menekan ESC di task pane.
const activeExecutions = new Map();
// Referensi pembatalan proses OpenCode yang sedang berjalan (maksimum satu).
const activeOpencodeCancel = { cancel: null };

function registerExecution(key, child) {
    if (!key || !child) return;
    let set = activeExecutions.get(key);
    if (!set) {
        set = new Set();
        activeExecutions.set(key, set);
    }
    set.add(child);
}

function unregisterExecution(key, child) {
    const set = activeExecutions.get(key);
    if (!set) return;
    set.delete(child);
    if (!set.size) activeExecutions.delete(key);
}

// Hentikan seluruh proses eksekusi untuk satu workbook (atau semuanya bila
// key kosong). Mengembalikan jumlah proses yang dihentikan.
function cancelExecutions(key) {
    let stopped = 0;
    const keys = key ? [key] : Array.from(activeExecutions.keys());
    for (const k of keys) {
        const set = activeExecutions.get(k);
        if (!set) continue;
        for (const child of Array.from(set)) {
            try {
                child.cancelled = true;
                killProcessTree(child);
                stopped += 1;
            } catch (e) { /* proses sudah mati */ }
        }
        activeExecutions.delete(k);
    }
    return stopped;
}

// Validasi ketat argumen CLI yang diteruskan ke shell (cegah injeksi perintah).
const OPENCODE_MODEL_RE = /^[A-Za-z0-9][A-Za-z0-9._\/-]{0,100}$/;
const OPENCODE_SESSION_RE = /^[A-Za-z0-9_-]{1,80}$/;

// Bunuh seluruh pohon proses. Di Windows, spawn memakai shell (cmd.exe)
// sehingga child.kill() hanya membunuh cmd dan membiarkan opencode.exe
// hidup sebagai proses yatim; taskkill /T mematikan sampai ke cucu proses.
function killProcessTree(child) {
    if (!child || child.pid == null) return;
    try {
        if (process.platform === 'win32') {
            require('child_process').exec('taskkill /pid ' + child.pid + ' /T /F', () => { });
        } else {
            child.kill('SIGKILL');
        }
    } catch (e) { }
}

// Jalankan OpenCode secara one-shot (mode `run`). Prompt dikirim lewat stdin
// agar tidak terkena batas panjang argumen command-line Windows.
// Output --format json berupa JSONL; teks jawaban diambil dari message terakhir.
// cancelRef (opsional): objek yang akan diisi fungsi cancel() agar pemanggil
// bisa membatalkan eksekusi (mis. saat client memutuskan koneksi).
function runOpencode(prompt, model, sessionId, dir, progressFile, cancelRef, imagePaths) {
    const { spawn } = require('child_process');
    return new Promise((resolve, reject) => {
        // --auto: setujui permission tool otomatis. TANPA ini opencode HANG
        // selamanya menunggu approval di mode non-interaktif (inilah penyebab
        // timeout 440 detik sebelumnya).
        const args = ['run', '--format', 'json', '--auto'];
        if (model) {
            if (!OPENCODE_MODEL_RE.test(model)) { reject(new Error('Nama model OpenCode tidak valid.')); return; }
            args.push('--model', model);
        }
        if (sessionId) {
            if (!OPENCODE_SESSION_RE.test(sessionId)) { reject(new Error('Session ID OpenCode tidak valid.')); return; }
            args.push('--session', sessionId);
        }
        (imagePaths || []).forEach(imagePath => args.push('--file', imagePath));

        let child;
        try {
            // cwd = folder scratch khusus, BUKAN folder proyek: dengan --auto
            // opencode bebas memakai tool, jadi jangan arahkan ke kode proyek.
            const ocWorkdir = path.join(os.tmpdir(), 'excelagent-opencode');
            try { fs.mkdirSync(ocWorkdir, { recursive: true }); } catch (e) { /* sudah ada */ }
            child = spawn('opencode', args, {
                cwd: ocWorkdir,
                shell: true,
                windowsHide: true
            });
        } catch (e) {
            reject(new Error('Gagal menjalankan opencode: ' + e.message));
            return;
        }

        if (cancelRef) {
            cancelRef.cancel = () => killProcessTree(child);
        }

        let buffer = '';
        let stderrTail = '';
        let errorEventMsg = '';
        let outSessionId = null;
        let lastMessageId = null;
        const finalTexts = [];
        let finished = false;

        const writeProgress = (msg) => {
            if (!progressFile) return;
            try { fs.writeFileSync(progressFile, msg, 'utf8'); } catch (e) { }
        };

        const timer = setTimeout(() => {
            if (finished) return;
            finished = true;
            killProcessTree(child);
            reject(new Error('OpenCode timeout setelah ' + Math.round(OPENCODE_TIMEOUT_MS / 1000) + ' detik.'));
        }, OPENCODE_TIMEOUT_MS);

        writeProgress('OpenCode sedang menganalisis secara mendalam...');

        child.stdout.on('data', (d) => {
            buffer += d.toString();
            let idx;
            while ((idx = buffer.indexOf('\n')) >= 0) {
                const line = buffer.slice(0, idx).trim();
                buffer = buffer.slice(idx + 1);
                if (!line.startsWith('{')) continue;
                let ev;
                try { ev = JSON.parse(line); } catch (e) { continue; }
                if (ev.sessionID) outSessionId = ev.sessionID;
                const part = ev.part || {};
                if (ev.type === 'text' && typeof part.text === 'string') {
                    if (part.messageID !== lastMessageId) {
                        lastMessageId = part.messageID;
                        finalTexts.length = 0; // hanya message terakhir yang dipakai
                    }
                    finalTexts.push(part.text);
                    writeProgress('OpenCode sedang menulis jawaban...');
                } else if (ev.type === 'error') {
                    const e = ev.error || {};
                    errorEventMsg = (e.data && e.data.message) || e.message || 'error tanpa pesan';
                    writeProgress('OpenCode menemui error, menunggu pemulihan...');
                } else if (part.type === 'tool' && part.tool) {
                    writeProgress('OpenCode memakai tool: ' + part.tool);
                } else if (ev.type === 'step_start') {
                    writeProgress('OpenCode: menyusun langkah penalaran...');
                }
            }
        });
        child.stderr.on('data', (d) => {
            stderrTail = (stderrTail + d.toString()).slice(-800);
        });
        child.on('error', (err) => {
            if (finished) return;
            finished = true;
            clearTimeout(timer);
            reject(new Error('Gagal menjalankan opencode: ' + err.message));
        });
        child.on('close', (code) => {
            if (finished) return;
            finished = true;
            clearTimeout(timer);
            const text = finalTexts.join('');
            if (!text.trim()) {
                reject(new Error('OpenCode tidak menghasilkan jawaban (exit ' + code + '). ' + (errorEventMsg || stderrTail)));
            } else {
                resolve({ ok: true, text: text, sessionId: outSessionId });
            }
        });

        child.stdin.on('error', () => { });
        child.stdin.write(prompt);
        child.stdin.end();
    });
}

// Penyimpanan setting permanen di profil user (bukan di dalam repo,
// supaya API key tidak ikut ter-commit dan tidak hilang saat cache Office dibersihkan).
const CONFIG_DIR = path.join(process.env.APPDATA || os.homedir(), 'ExcelAgent');
const CONFIG_FILE = path.join(CONFIG_DIR, 'config.json');

const DEFAULT_CONFIG = {
    apiKey: '',
    apiEndpoint: '',
    model: '',
    autoRun: true,
    useOpencode: false,
    opencodeModel: '9router/qd/ultimate',
    providers: {},
    activeProvider: '',
    modelList: [],
    chatHistoryByWorkbook: {}
};

function isAllowedOrigin(origin) {
    return !origin || origin === ALLOWED_ORIGIN;
}

function normalizeWorkbookPath(value) {
    if (typeof value !== 'string') {
        throw new Error('Path target workbook tidak valid.');
    }
    const trimmed = value.trim();
    if (/^https?:\/\//i.test(trimmed)) {
        let url;
        try { url = new URL(trimmed); } catch (e) { throw new Error('URL target workbook tidak valid.'); }
        const host = url.hostname.toLowerCase();
        if (url.protocol !== 'https:' || !(host === 'd.docs.live.net' || host.endsWith('.sharepoint.com'))) {
            throw new Error('URL workbook cloud tidak didukung.');
        }
        const parts = decodeURIComponent(url.pathname).split('/');
        if (host === 'd.docs.live.net' && parts.length > 1) parts[1] = parts[1].toLowerCase();
        return 'https://' + host + parts.join('/');
    }
    if (!path.win32.isAbsolute(trimmed)) throw new Error('Path target workbook tidak valid.');
    return path.win32.normalize(trimmed).toLowerCase();
}

function workbookNameFromPath(value) {
    const normalized = normalizeWorkbookPath(value);
    if (/^https:\/\//i.test(normalized)) return normalized.substring(normalized.lastIndexOf('/') + 1);
    return path.win32.basename(normalized);
}

// Konversi URL OneDrive ke path lokal sync folder.
// Return path lokal jika file ditemukan di OneDrive sync folder, null jika tidak.
function resolveOneDriveLocalPath(workbookPath) {
    // Hanya proses URL cloud, bukan path lokal
    if (!/^https?:\/\//i.test(workbookPath)) return null;
    
    // Extract nama file dari URL
    const fileName = workbookNameFromPath(workbookPath);
    if (!fileName) return null;
    
    // Decode URL-encoded characters (misal %20 jadi spasi)
    let decodedName;
    try {
        decodedName = decodeURIComponent(fileName);
    } catch (e) {
        decodedName = fileName;
    }
    
    // Cari file di OneDrive sync folder menggunakan recursive search
    const onedrivePaths = [
        process.env.OneDrive,
        process.env.OneDriveConsumer,
        process.env.OneDriveCommercial
    ].filter(Boolean);
    
    for (const odPath of onedrivePaths) {
        try {
            const found = findFileRecursive(odPath, decodedName, 0);
            if (found) return found;
        } catch (e) {
            // Abaikan error, coba path berikutnya
        }
    }
    return null;
}

// Recursive file search dengan depth limit
const MAX_SEARCH_DEPTH = 10;
function findFileRecursive(dir, fileName, depth) {
    if (depth > MAX_SEARCH_DEPTH) return null;
    try {
        const entries = fs.readdirSync(dir, { withFileTypes: true });
        for (const entry of entries) {
            const fullPath = path.join(dir, entry.name);
            if (entry.isFile() && entry.name.toLowerCase() === fileName.toLowerCase()) {
                return fullPath;
            }
            if (entry.isDirectory()) {
                // Skip folder sistem dan hidden
                if (entry.name.startsWith('.') || entry.name === 'node_modules' || entry.name === '.git') continue;
                const found = findFileRecursive(fullPath, fileName, depth + 1);
                if (found) return found;
            }
        }
    } catch (e) {
        // Permission denied atau error lain, abaikan
    }
    return null;
}

function workbookHistoryKey(workbookPath) {
    return crypto.createHash('sha256').update(normalizeWorkbookPath(workbookPath)).digest('hex');
}

// Kunci berbasis ID permanen yang tertanam di dalam file (custom property
// ExcelAgentId). Duplikat/ganti nama file tetap membawa ID yang sama, sehingga
// riwayat chat ikut muncul.
function workbookIdKey(workbookId) {
    if (typeof workbookId !== 'string') return null;
    const trimmed = workbookId.trim().toLowerCase();
    if (!trimmed || trimmed.length > 128) return null;
    return crypto.createHash('sha256').update('wbid:' + trimmed).digest('hex');
}

class ExecutionValidationError extends Error {
    constructor(message, targetError = false) {
        super(message);
        this.name = 'ExecutionValidationError';
        this.targetError = targetError;
    }
}

function validateExecutionEnvelope(body) {
    if (!body || typeof body.code !== 'string' || !body.code.trim()) throw new ExecutionValidationError('Kode Python kosong.');
    if (!body.target || typeof body.target !== 'object') throw new ExecutionValidationError('Metadata target workbook tidak tersedia.', true);
    ['sessionId', 'workbookName', 'workbookPath'].forEach((key) => {
        if (typeof body.target[key] !== 'string' || !body.target[key].trim()) {
            throw new ExecutionValidationError('Target workbook tidak lengkap: ' + key, true);
        }
    });
    try {
        normalizeWorkbookPath(body.target.workbookPath);
    } catch (e) {
        throw new ExecutionValidationError(e.message, true);
    }
    if (workbookNameFromPath(body.target.workbookPath).toLowerCase() !== body.target.workbookName.trim().toLowerCase()) {
        throw new ExecutionValidationError('Nama dan path workbook target tidak cocok.', true);
    }
    return { code: body.code, target: body.target };
}

function publicSettings(config) {
    const { chatHistory, chatHistoryByWorkbook, providers, ...publicConfig } = config;
    delete publicConfig.apiKey;
    publicConfig.hasApiKey = typeof config.apiKey === 'string' && config.apiKey.length > 0;
    // Sertakan daftar provider TANPA membocorkan API key: key mentah diganti flag hasKey.
    if (providers && typeof providers === 'object') {
        const safe = {};
        for (const name of Object.keys(providers)) {
            const p = providers[name] || {};
            safe[name] = {
                apiEndpoint: p.apiEndpoint || '',
                model: p.model || '',
                // Daftar model milik provider ini saja (dipisah dari modelList global).
                models: Array.isArray(p.models) ? p.models.slice() : [],
                opencodeModel: p.opencodeModel || '',
                useOpencode: p.useOpencode === true,
                hasKey: typeof p.apiKey === 'string' && p.apiKey.length > 0
            };
        }
        publicConfig.providers = safe;
    }
    return publicConfig;
}

function enqueueCom(work) {
    if (pendingCom >= MAX_PENDING_COM) {
        const error = new Error('Antrean eksekusi penuh. Coba lagi nanti.');
        error.statusCode = 429;
        return Promise.reject(error);
    }
    pendingCom += 1;
    const queued = comQueue.then(work, work);
    comQueue = queued.catch(() => undefined);
    return queued.finally(() => { pendingCom -= 1; });
}

function readConfig() {
    try {
        const raw = fs.readFileSync(CONFIG_FILE, 'utf8');
        return Object.assign({}, DEFAULT_CONFIG, JSON.parse(raw));
    } catch (e) {
        return Object.assign({}, DEFAULT_CONFIG);
    }
}

function writeConfig(partial) {
    const merged = Object.assign(readConfig(), partial);
    fs.mkdirSync(CONFIG_DIR, { recursive: true });
    fs.writeFileSync(CONFIG_FILE, JSON.stringify(merged, null, 2), 'utf8');
    // Batasi akses ke pemilik saja (best-effort; diabaikan bila FS tidak mendukung)
    try { fs.chmodSync(CONFIG_FILE, 0o600); } catch (e) { }
    return merged;
}

// Migrasi non-destruktif ke skema "model per provider".
// Provider lama belum punya field `models`; diisi dari model aktifnya sendiri.
// `modelList` global TIDAK dihapus — tetap disimpan sebagai arsip/fallback
// sehingga daftar model yang pernah dikumpulkan user tidak hilang.
function migrateProviderModels(readFn = readConfig, writeFn = writeConfig) {
    const config = readFn();
    const providers = config.providers || {};
    const names = Object.keys(providers);
    if (!names.length) return { migrated: 0, config };
    const next = {};
    let migrated = 0;
    for (const name of names) {
        const p = Object.assign({}, providers[name]);
        if (!Array.isArray(p.models)) {
            p.models = p.model ? [p.model] : [];
            migrated += 1;
        }
        next[name] = p;
    }
    if (!migrated) return { migrated: 0, config };
    return { migrated, config: writeFn({ providers: next }) };
}

function cleanHistory(value) {
    if (!Array.isArray(value)) throw new Error('Riwayat chat harus berupa array.');
    return value.filter(item => item && (item.role === 'user' || item.role === 'assistant') && typeof item.content === 'string').slice(-50);
}

function readWorkbookHistory(workbookPath, workbookId) {
    const config = readConfig();
    const histories = config.chatHistoryByWorkbook || {};
    const idKey = workbookIdKey(workbookId);
    if (idKey && histories[idKey]) return histories[idKey];
    if (workbookPath) {
        try { return histories[workbookHistoryKey(workbookPath)] || []; } catch (e) { }
    }
    return [];
}

function writeWorkbookHistory(workbookPath, history, workbookId) {
    const config = readConfig();
    const histories = Object.assign({}, config.chatHistoryByWorkbook || {});
    const cleaned = cleanHistory(history);
    const idKey = workbookIdKey(workbookId);
    let pathKey = null;
    try { if (workbookPath) pathKey = workbookHistoryKey(workbookPath); } catch (e) { }

    // Simpan di bawah ID bila ada, dan tautkan juga ke path untuk backward-compatibility.
    if (idKey) histories[idKey] = cleaned;
    if (pathKey) histories[pathKey] = cleaned;
    writeConfig({ chatHistoryByWorkbook: histories });
    return cleaned;
}

function resolveProxyApiKey(requestApiKey, storedApiKey) {
    return typeof requestApiKey === 'string' && requestApiKey.length > 0 ? requestApiKey : storedApiKey;
}

function deleteWorkbookHistory(workbookPath) {
    const config = readConfig();
    const histories = Object.assign({}, config.chatHistoryByWorkbook || {});
    delete histories[workbookHistoryKey(workbookPath)];
    writeConfig({ chatHistoryByWorkbook: histories });
}

// ---- Provider profiles: simpan banyak set (endpoint+key+model), pilih untuk aktif ----

function sanitizeProviderName(name) {
    if (typeof name !== 'string') return '';
    const trimmed = name.trim();
    if (!trimmed || trimmed.length > 60) return '';
    return trimmed;
}

// Simpan/perbarui satu provider. Key hanya ditimpa bila incoming membawa key baru,
// sehingga menyimpan ulang tanpa mengetik key tidak menghapus key lama.
function saveProvider(name, data) {
    const cleanName = sanitizeProviderName(name);
    if (!cleanName) { const e = new Error('Nama provider tidak valid.'); e.statusCode = 400; throw e; }
    const config = readConfig();
    const providers = Object.assign({}, config.providers || {});
    const existing = providers[cleanName] || {};
    const entry = {
        apiEndpoint: typeof data.apiEndpoint === 'string' ? data.apiEndpoint : (existing.apiEndpoint || ''),
        model: typeof data.model === 'string' ? data.model : (existing.model || ''),
        // models: daftar milik provider ini. Bila tidak dikirim, daftar lama dipertahankan.
        models: Array.isArray(data.models) ? data.models.slice() : (Array.isArray(existing.models) ? existing.models.slice() : []),
        opencodeModel: typeof data.opencodeModel === 'string' ? data.opencodeModel : (existing.opencodeModel || ''),
        useOpencode: data.useOpencode !== undefined ? data.useOpencode === true : (existing.useOpencode === true),
        apiKey: (typeof data.apiKey === 'string' && data.apiKey.length > 0) ? data.apiKey : (existing.apiKey || '')
    };
    // Model aktif selalu ikut masuk daftar provider agar tidak hilang dari dropdown.
    if (entry.model && entry.models.indexOf(entry.model) === -1) entry.models.unshift(entry.model);
    providers[cleanName] = entry;
    writeConfig({ providers });
    return cleanName;
}

function deleteProvider(name) {
    const cleanName = sanitizeProviderName(name);
    const config = readConfig();
    const providers = Object.assign({}, config.providers || {});
    delete providers[cleanName];
    const patch = { providers };
    if (config.activeProvider === cleanName) patch.activeProvider = '';
    writeConfig(patch);
}

// Aktifkan provider: salin field-nya (termasuk key) ke setting aktif yang dipakai proxy.
function activateProvider(name) {
    const cleanName = sanitizeProviderName(name);
    const config = readConfig();
    const providers = config.providers || {};
    const p = providers[cleanName];
    if (!p) { const e = new Error('Provider tidak ditemukan.'); e.statusCode = 404; throw e; }
    const patch = {
        apiEndpoint: p.apiEndpoint || '',
        model: p.model || '',
        activeProvider: cleanName
    };
    if (p.opencodeModel) patch.opencodeModel = p.opencodeModel;
    if (p.useOpencode !== undefined) patch.useOpencode = p.useOpencode === true;
    if (typeof p.apiKey === 'string' && p.apiKey.length > 0) patch.apiKey = p.apiKey;
    return writeConfig(patch);
}

const MIME_TYPES = {
    '.html': 'text/html; charset=utf-8',
    '.js': 'application/javascript',
    '.css': 'text/css',
    '.json': 'application/json',
    '.png': 'image/png',
    '.svg': 'image/svg+xml',
    '.xml': 'application/xml'
};

const STATIC_ROOTS = {
    taskpane: path.resolve(__dirname, 'taskpane'),
    assets: path.resolve(__dirname, 'assets')
};

function resolveContainedFile(root, relativePath) {
    const candidate = path.resolve(root, relativePath);
    const relative = path.relative(root, candidate);
    if (relative === '' || relative.startsWith('..' + path.sep) || path.isAbsolute(relative)) return null;
    return candidate;
}

// Layani hanya file task pane dan aset yang memang dibutuhkan add-in.
function resolveFilePath(reqUrl) {
    let urlPath;
    try {
        urlPath = decodeURIComponent(reqUrl.split('?')[0]);
    } catch (e) {
        return null;
    }
    if (urlPath === '/') urlPath = '/taskpane.html';
    if (!urlPath.startsWith('/') || path.posix.isAbsolute(urlPath.slice(1)) || path.win32.isAbsolute(urlPath.slice(1))) return null;
    if (urlPath.includes('\\')) return null;

    const segments = urlPath.split('/');
    if (segments.some(segment => segment === '..')) return null;

    if (urlPath.startsWith('/assets/')) {
        return resolveContainedFile(STATIC_ROOTS.assets, urlPath.slice('/assets/'.length));
    }
    if (segments.length === 2) {
        return resolveContainedFile(STATIC_ROOTS.taskpane, segments[1]);
    }
    return null;
}

// Baca body dari request
function readBody(req, maxBytes = MAX_REQUEST_BODY_BYTES) {
    return new Promise((resolve, reject) => {
        const chunks = [];
        let size = 0;
        let settled = false;
        function fail(error) {
            if (settled) return;
            settled = true;
            reject(error);
        }
        req.on('data', c => {
            size += c.length;
            if (size > maxBytes) {
                const error = new Error('Body request terlalu besar.');
                error.statusCode = 413;
                fail(error);
                // Tolak buffering, tetapi teruskan drain request agar koneksi tetap sehat.
                if (req.resume) req.resume();
                return;
            }
            if (!settled) chunks.push(c);
        });
        req.on('end', () => { if (!settled) { settled = true; resolve(Buffer.concat(chunks).toString()); } });
        req.on('error', fail);
    });
}

function prepareOpencodeImages(images, root = os.tmpdir()) {
    if (images === undefined) return { paths: [], cleanup: () => {} };
    if (!Array.isArray(images) || images.length > MAX_OPENCODE_IMAGES) {
        throw new Error('Lampiran gambar maksimal ' + MAX_OPENCODE_IMAGES + ' gambar.');
    }
    const requestDir = fs.mkdtempSync(path.join(root, 'excelagent-images-'));
    const paths = [];
    try {
        images.forEach((image, index) => {
            if (!image || typeof image.data !== 'string' || typeof image.type !== 'string') {
                throw new Error('Data lampiran gambar tidak valid.');
            }
            const ext = IMAGE_MIME_EXT[image.type.toLowerCase()];
            if (!ext) throw new Error('Format gambar tidak didukung. Gunakan PNG, JPG, atau WEBP.');
            const match = image.data.match(/^data:([^;]+);base64,([A-Za-z0-9+/=\s]+)$/);
            if (!match || match[1].toLowerCase() !== image.type.toLowerCase()) {
                throw new Error('Data gambar harus berupa Base64 data URL yang valid.');
            }
            const data = Buffer.from(match[2].replace(/\s/g, ''), 'base64');
            if (!data.length || data.length > MAX_IMAGE_BYTES) {
                throw new Error('Ukuran setiap gambar maksimal 5 MB.');
            }
            const filePath = path.join(requestDir, 'image-' + index + ext);
            fs.writeFileSync(filePath, data, { flag: 'wx' });
            paths.push(filePath);
        });
    } catch (error) {
        fs.rmSync(requestDir, { recursive: true, force: true });
        throw error;
    }
    return {
        paths,
        cleanup: () => fs.rmSync(requestDir, { recursive: true, force: true })
    };
}

// Fungsi utilitas pengecekan kesehatan server, router/gateway AI, dan python runtime
function checkRouterHealth(apiEndpoint, apiKey, timeoutMs = 2500) {
    return new Promise((resolve) => {
        if (!apiEndpoint || typeof apiEndpoint !== 'string') {
            return resolve({ ok: false, status: 'unconfigured', message: 'API Endpoint belum disetting.' });
        }
        let parsedUrl;
        try {
            parsedUrl = new URL(apiEndpoint);
        } catch (e) {
            return resolve({ ok: false, status: 'invalid_url', message: 'Format URL API tidak valid.' });
        }

        const isHttps = parsedUrl.protocol === 'https:';
        const transport = isHttps ? https : http;
        
        // Buat path /models atau gunakan path asal jika sudah spesifik
        let probePath = parsedUrl.pathname.replace(/\/chat\/completions\/?$/, '').replace(/\/+$/, '') + '/models';
        if (!probePath.startsWith('/')) probePath = '/' + probePath;

        const options = {
            hostname: parsedUrl.hostname,
            port: parsedUrl.port || (isHttps ? 443 : 80),
            path: probePath,
            method: 'GET',
            headers: {
                'User-Agent': 'ExcelAgent-HealthCheck/1.0',
                ...(apiKey ? { 'Authorization': `Bearer ${apiKey}` } : {})
            },
            timeout: timeoutMs,
            rejectUnauthorized: false
        };

        const startTime = Date.now();
        const req = transport.request(options, (res) => {
            const latencyMs = Date.now() - startTime;
            // 200 OK, atau 401/403 (berarti router hidup tapi butuh auth), atau 404 dianggap router reachable
            const isReachable = res.statusCode >= 200 && res.statusCode < 500;
            res.resume(); // drain data
            resolve({
                ok: isReachable,
                status: isReachable ? 'online' : 'error',
                statusCode: res.statusCode,
                latencyMs
            });
        });

        req.on('timeout', () => {
            req.destroy();
            resolve({ ok: false, status: 'timeout', message: `Timeout setelah ${timeoutMs}ms` });
        });

        req.on('error', (err) => {
            resolve({ ok: false, status: 'offline', message: err.code || err.message });
        });

        req.end();
    });
}

function checkPythonHealth(pythonPath = (process.env.PYTHON || 'python'), timeoutMs = 2500) {
    return new Promise((resolve) => {
        const { execFile } = require('child_process');
        const script = 'import sys; has_com = False\ntry:\n import win32com.client; has_com=True\nexcept: pass\nprint(f"{sys.version_info.major}.{sys.version_info.minor}|{has_com}")';
        
        execFile(pythonPath, ['-c', script], { timeout: timeoutMs }, (err, stdout) => {
            if (err) {
                return resolve({ ok: false, status: 'offline', message: err.message });
            }
            try {
                const parts = (stdout || '').trim().split('|');
                const pyVer = parts[0];
                const hasWin32com = parts[1] === 'True';
                resolve({
                    ok: true,
                    status: 'online',
                    version: pyVer,
                    hasWin32com
                });
            } catch (e) {
                resolve({ ok: true, status: 'online', hasWin32com: false });
            }
        });
    });
}

// Proxy khusus Anthropic: terjemahkan format OpenAI <-> Messages API dan
// pakai login OAuth Claude Code (atau API key sk-ant-api...).
async function proxyToAnthropic(parsedBody, res, readConfigFn = readConfig) {
    const payload = parsedBody.payload || {};
    const storedApiKey = readConfigFn().apiKey;

    const safeEnd = (statusCode, jsonBody) => {
        if (res.writableEnded) return;
        try {
            if (!res.headersSent) {
                res.writeHead(statusCode, {
                    'Content-Type': 'application/json',
                    'Access-Control-Allow-Origin': ALLOWED_ORIGIN
                });
            }
            res.end(jsonBody);
        } catch (e) { /* koneksi sudah tertutup */ }
    };

    let auth;
    try {
        // API key yang dipakai request hanya diterima bila benar-benar key Anthropic.
        const candidate = typeof parsedBody.apiKey === 'string' && parsedBody.apiKey.indexOf('sk-ant-api') === 0
            ? parsedBody.apiKey
            : storedApiKey;
        auth = await anthropic.resolveAuth(candidate);
    } catch (e) {
        safeEnd(401, JSON.stringify({ error: { message: e.message, type: 'authentication_error' } }));
        return;
    }

    const wantStream = payload.stream !== false;
    const baseRequest = anthropic.toAnthropicRequest(
        Object.assign({}, payload, { stream: wantStream }),
        auth.mode
    );

    // Kirim request; bila ditolak karena parameter sampling (temperature/top_p)
    // tidak didukung model, ulangi sekali tanpa parameter itu.
    // Aman karena pada kasus error upstream belum ada byte yang dikirim ke client.
    function send(anthropicRequest, allowSamplingRetry) {
        const body = JSON.stringify(anthropicRequest);
        const headers = Object.assign(anthropic.authHeaders(auth), {
            'Content-Length': Buffer.byteLength(body)
        });

        const proxyReq = https.request({
            hostname: anthropic.ANTHROPIC_HOST,
            path: '/v1/messages',
            method: 'POST',
            headers
        }, (proxyRes) => {
            const upstreamOk = proxyRes.statusCode >= 200 && proxyRes.statusCode < 300;

            if (wantStream && upstreamOk) {
                if (res.headersSent) return;
                res.writeHead(200, {
                    'Content-Type': 'text/event-stream',
                    'Access-Control-Allow-Origin': ALLOWED_ORIGIN,
                    'Cache-Control': 'no-cache',
                    'Connection': 'keep-alive',
                    'X-Accel-Buffering': 'no'
                });
                const translator = anthropic.createSseTranslator(
                    anthropicRequest.model,
                    (text) => { if (!res.writableEnded) res.write(text); }
                );
                proxyRes.on('data', (chunk) => translator.push(chunk));
                proxyRes.on('end', () => { translator.end(); if (!res.writableEnded) res.end(); });
                proxyRes.on('error', () => { if (!res.writableEnded) res.end(); });
                res.on('error', () => proxyRes.destroy());
                return;
            }

            const chunks = [];
            proxyRes.on('data', c => chunks.push(c));
            proxyRes.on('error', () => safeEnd(502, JSON.stringify({ error: { message: 'Koneksi ke Anthropic terputus.' } })));
            proxyRes.on('end', () => {
                const text = Buffer.concat(chunks).toString();
                if (!upstreamOk) {
                    // Teruskan pesan error Anthropic apa adanya agar bisa dibaca user.
                    let message = text.slice(0, 500);
                    try {
                        const parsed = JSON.parse(text);
                        if (parsed && parsed.error && parsed.error.message) message = parsed.error.message;
                    } catch (e) { /* biarkan teks mentah */ }

                    // Model menolak temperature/top_p: catat lalu ulangi tanpa parameter itu.
                    if (allowSamplingRetry && anthropic.isSamplingUnsupportedError(message)) {
                        const stripped = anthropic.stripSamplingParams(anthropicRequest);
                        if (stripped.removed) {
                            anthropic.rememberRejectsSampling(anthropicRequest.model);
                            console.log(`[Anthropic] ${anthropicRequest.model} menolak parameter sampling; mengulang tanpa temperature/top_p.`);
                            send(stripped.request, false);
                            return;
                        }
                    }

                    safeEnd(proxyRes.statusCode, JSON.stringify({ error: { message: message, type: 'anthropic_error' } }));
                    return;
                }
                try {
                    safeEnd(200, JSON.stringify(anthropic.toOpenAiResponse(JSON.parse(text))));
                } catch (e) {
                    safeEnd(502, JSON.stringify({ error: { message: 'Respons Anthropic tidak dapat dibaca.' } }));
                }
            });
        });

        proxyReq.setTimeout(PROXY_TIMEOUT_MS, () => proxyReq.destroy(new Error('timeout')));
        proxyReq.on('error', (err) => {
            const isTimeout = err && err.message === 'timeout';
            safeEnd(isTimeout ? 504 : 502, JSON.stringify({
                error: { message: isTimeout ? 'Anthropic tidak merespons (timeout).' : ('Gagal menghubungi Anthropic: ' + err.message) }
            }));
        });
        proxyReq.end(body);
    }

    send(baseRequest, true);
}

// Proxy request ke API eksternal (menghindari CORS)
async function proxyToAPI(reqBody, res, readConfigFn = readConfig) {
    let proxyReq = null;
    try {
        const parsedRequest = JSON.parse(reqBody);
        const { apiEndpoint, apiKey, payload } = parsedRequest;

        // Endpoint Anthropic dialihkan ke adapter Messages API.
        if (anthropic.isAnthropicEndpoint(apiEndpoint)) {
            await proxyToAnthropic(parsedRequest, res, readConfigFn);
            return;
        }

        const storedApiKey = readConfigFn().apiKey;

        // Support streaming - default to true for better UX
        if (payload.stream === undefined) payload.stream = true;

        const url = new URL(apiEndpoint);
        const transport = url.protocol === 'https:' ? https : http;

        const headers = { 'Content-Type': 'application/json' };
        const selectedApiKey = resolveProxyApiKey(apiKey, storedApiKey);
        if (selectedApiKey) headers['Authorization'] = `Bearer ${selectedApiKey}`;

        const body = JSON.stringify(payload);
        headers['Content-Length'] = Buffer.byteLength(body);

        const options = {
            hostname: url.hostname,
            port: url.port || (url.protocol === 'https:' ? 443 : 80),
            path: url.pathname + url.search,
            method: 'POST',
            headers,
            rejectUnauthorized: false  // gateway self-signed OK
        };

        const safeEnd = (statusCode, jsonBody) => {
            if (res.writableEnded) return;
            try {
                if (!res.headersSent) {
                    res.writeHead(statusCode, {
                        'Content-Type': 'application/json',
                        'Access-Control-Allow-Origin': ALLOWED_ORIGIN
                    });
                }
                res.end(jsonBody);
            } catch (e) {
                // Koneksi sudah tertutup; abaikan.
            }
        };

        proxyReq = transport.request(options, (proxyRes) => {
            if (payload.stream && proxyRes.statusCode >= 200 && proxyRes.statusCode < 300) {
                // Streaming mode - pipe SSE langsung ke client.
                if (res.headersSent) return;
                res.writeHead(proxyRes.statusCode, {
                    'Content-Type': 'text/event-stream',
                    'Access-Control-Allow-Origin': ALLOWED_ORIGIN,
                    'Cache-Control': 'no-cache',
                    'Connection': 'keep-alive',
                    'X-Accel-Buffering': 'no'
                });
                proxyRes.on('error', () => { if (!res.writableEnded) res.end(); });
                res.on('error', () => { proxyRes.destroy(); });
                proxyRes.pipe(res);
            } else {
                // Non-streaming, ATAU upstream error saat mode streaming:
                // buffer body lalu teruskan apa adanya agar client bisa membaca pesan error.
                const chunks = [];
                proxyRes.on('data', c => chunks.push(c));
                proxyRes.on('error', () => { if (!res.headersSent) safeEnd(502, JSON.stringify({ error: { message: 'Proxy error: koneksi upstream terputus' } })); });
                proxyRes.on('end', () => {
                    const responseBody = Buffer.concat(chunks).toString();
                    const contentType = payload.stream
                        ? 'application/json'
                        : (proxyRes.headers['content-type'] || 'application/json');
                    if (res.headersSent) return;
                    try {
                        res.writeHead(proxyRes.statusCode, {
                            'Content-Type': contentType,
                            'Access-Control-Allow-Origin': ALLOWED_ORIGIN
                        });
                        res.end(responseBody);
                    } catch (e) { /* koneksi sudah tertutup */ }
                });
            }
        });

        // Timeout hanya untuk fase idle (menunggu token pertama / selambat-lambatnya
        // keseluruhan respons non-streaming). Saat SSE mengalir, socket aktif terus
        // sehingga timer otomatis ter-reset dan tidak memicu "Gateway timeout" palsu.
        proxyReq.setTimeout(PROXY_TIMEOUT_MS, () => {
            proxyReq.destroy(new Error('timeout'));
        });

        proxyReq.on('error', (err) => {
            const isTimeout = err && (err.message === 'timeout' || err.code === 'ECONNRESET');
            if (res.headersSent || res.writableEnded) return;
            if (isTimeout) {
                safeEnd(504, JSON.stringify({ error: { message: 'Gateway timeout' } }));
            } else {
                safeEnd(502, JSON.stringify({ error: { message: 'Proxy error: ' + err.message } }));
            }
        });

        // Jika client (task pane) menutup koneksi, hentikan request upstream
        // supaya model tidak terus menghitung untuk hasil yang tidak dipakai.
        res.on('close', () => {
            if (!res.writableEnded && proxyReq && !proxyReq.destroyed) {
                proxyReq.destroy();
            }
        });

        proxyReq.write(body);
        proxyReq.end();

    } catch (err) {
        if (res.headersSent || res.writableEnded) return;
        res.writeHead(400, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ error: { message: 'Invalid request: ' + err.message } }));
    }
}

function createRequestHandler(dependencies = {}) {
    const readConfigForRequest = dependencies.readConfig || readConfig;
    const writeConfigForRequest = dependencies.writeConfig || writeConfig;
    const execFileForRequest = dependencies.execFile || require('child_process').execFile;
    const proxyToAPIForRequest = dependencies.proxyToAPI || proxyToAPI;

    const readWorkbookHistoryForRequest = (workbookPath, workbookId) => {
        const config = readConfigForRequest();
        const histories = config.chatHistoryByWorkbook || {};
        const idKey = workbookIdKey(workbookId);
        if (idKey && histories[idKey]) return histories[idKey];
        if (workbookPath) {
            try { return histories[workbookHistoryKey(workbookPath)] || []; } catch (e) { }
        }
        return [];
    };
    const writeWorkbookHistoryForRequest = (workbookPath, history, workbookId) => {
        const config = readConfigForRequest();
        const histories = Object.assign({}, config.chatHistoryByWorkbook || {});
        const cleaned = cleanHistory(history);
        const idKey = workbookIdKey(workbookId);
        let pathKey = null;
        try { if (workbookPath) pathKey = workbookHistoryKey(workbookPath); } catch (e) { }
        if (idKey) histories[idKey] = cleaned;
        if (pathKey) histories[pathKey] = cleaned;
        writeConfigForRequest({ chatHistoryByWorkbook: histories });
        return cleaned;
    };
    const deleteWorkbookHistoryForRequest = (workbookPath, workbookId) => {
        const config = readConfigForRequest();
        const histories = Object.assign({}, config.chatHistoryByWorkbook || {});
        const idKey = workbookIdKey(workbookId);
        if (idKey) delete histories[idKey];
        if (workbookPath) {
            try { delete histories[workbookHistoryKey(workbookPath)]; } catch (e) { }
        }
        writeConfigForRequest({ chatHistoryByWorkbook: histories });
    };

    // Provider profile wrappers (pakai read/write yang di-inject supaya bisa diuji).
    const saveProviderForRequest = (name, data) => {
        const cleanName = sanitizeProviderName(name);
        if (!cleanName) { const e = new Error('Nama provider tidak valid.'); e.statusCode = 400; throw e; }
        const config = readConfigForRequest();
        const providers = Object.assign({}, config.providers || {});
        const existing = providers[cleanName] || {};
        const entry = {
            apiEndpoint: typeof data.apiEndpoint === 'string' ? data.apiEndpoint : (existing.apiEndpoint || ''),
            model: typeof data.model === 'string' ? data.model : (existing.model || ''),
            models: Array.isArray(data.models) ? data.models.slice() : (Array.isArray(existing.models) ? existing.models.slice() : []),
            opencodeModel: typeof data.opencodeModel === 'string' ? data.opencodeModel : (existing.opencodeModel || ''),
            useOpencode: data.useOpencode !== undefined ? data.useOpencode === true : (existing.useOpencode === true),
            apiKey: (typeof data.apiKey === 'string' && data.apiKey.length > 0) ? data.apiKey : (existing.apiKey || '')
        };
        if (entry.model && entry.models.indexOf(entry.model) === -1) entry.models.unshift(entry.model);
        providers[cleanName] = entry;
        writeConfigForRequest({ providers });
        return cleanName;
    };
    const deleteProviderForRequest = (name) => {
        const cleanName = sanitizeProviderName(name);
        const config = readConfigForRequest();
        const providers = Object.assign({}, config.providers || {});
        delete providers[cleanName];
        const patch = { providers };
        if (config.activeProvider === cleanName) patch.activeProvider = '';
        writeConfigForRequest(patch);
    };
    const activateProviderForRequest = (name) => {
        const cleanName = sanitizeProviderName(name);
        const config = readConfigForRequest();
        const p = (config.providers || {})[cleanName];
        if (!p) { const e = new Error('Provider tidak ditemukan.'); e.statusCode = 404; throw e; }
        const patch = { apiEndpoint: p.apiEndpoint || '', model: p.model || '', activeProvider: cleanName };
        if (p.opencodeModel) patch.opencodeModel = p.opencodeModel;
        if (p.useOpencode !== undefined) patch.useOpencode = p.useOpencode === true;
        if (typeof p.apiKey === 'string' && p.apiKey.length > 0) patch.apiKey = p.apiKey;
        return writeConfigForRequest(patch);
    };

    return async (req, res) => {
        const origin = req.headers.origin;
        if (req.url.startsWith('/api/') && !isAllowedOrigin(origin)) {
            res.writeHead(403, { 'Content-Type': 'application/json' });
            res.end(JSON.stringify({ ok: false, error: 'Origin tidak diizinkan.' }));
            return;
        }

        // CORS preflight
        if (req.method === 'OPTIONS') {
            res.writeHead(204, {
                'Access-Control-Allow-Origin': ALLOWED_ORIGIN,
                'Access-Control-Allow-Methods': 'POST, GET, DELETE, OPTIONS',
                'Access-Control-Allow-Headers': 'Content-Type, Authorization'
            });
            res.end();
            return;
        }

        // Health Check: cek server, AI router/gateway, dan Python win32com
        if (req.url === '/api/health' && req.method === 'GET') {
            const config = readConfigForRequest();
            // Endpoint Anthropic tidak bisa diprobe via /models biasa: pakai cek login.
            const routerProbe = anthropic.isAnthropicEndpoint(config.apiEndpoint)
                ? anthropic.authStatus(config.apiKey).then((s) => ({
                    ok: s.ok === true,
                    status: s.ok ? 'online' : 'offline',
                    message: s.ok ? null : (s.error || 'Belum login Anthropic.')
                }))
                : checkRouterHealth(config.apiEndpoint, config.apiKey);
            const [routerStatus, pythonStatus] = await Promise.all([
                routerProbe,
                checkPythonHealth()
            ]);

            const healthData = {
                ok: true,
                timestamp: new Date().toISOString(),
                server: {
                    status: 'online',
                    port: PORT,
                    uptimeSeconds: Math.floor(process.uptime()),
                    version: '1.6.0'
                },
                router: {
                    status: routerStatus.status,
                    ok: routerStatus.ok,
                    endpoint: config.apiEndpoint || '',
                    activeModel: config.model || config.opencodeModel || '',
                    activeProvider: config.activeProvider || 'default',
                    latencyMs: routerStatus.latencyMs || null,
                    message: routerStatus.message || null
                },
                python: {
                    status: pythonStatus.status,
                    ok: pythonStatus.ok,
                    hasWin32com: pythonStatus.hasWin32com || false,
                    version: pythonStatus.version || null,
                    message: pythonStatus.message || null
                }
            };

            res.writeHead(200, {
                'Content-Type': 'application/json',
                'Access-Control-Allow-Origin': ALLOWED_ORIGIN,
                'Cache-Control': 'no-store'
            });
            res.end(JSON.stringify(healthData));
            return;
        }

        // Settings: baca konfigurasi tersimpan
        if (req.url === '/api/settings' && req.method === 'GET') {
            const config = readConfigForRequest();
            res.writeHead(200, {
                'Content-Type': 'application/json',
                'Access-Control-Allow-Origin': ALLOWED_ORIGIN,
                'Cache-Control': 'no-store'
            });
            res.end(JSON.stringify(publicSettings(config)));
            return;
        }

        // Settings: simpan konfigurasi
        if (req.url === '/api/settings' && req.method === 'POST') {
            try {
                const incoming = JSON.parse(await readBody(req));
                const allowed = {};
                if (typeof incoming.apiKey === 'string' && incoming.apiKey.length > 0) allowed.apiKey = incoming.apiKey;
                if (incoming.clearApiKey === true) allowed.apiKey = '';
                ['apiEndpoint', 'model', 'autoRun', 'useOpencode', 'opencodeModel', 'modelList'].forEach((k) => { if (incoming[k] !== undefined) allowed[k] = incoming[k]; });
                const saved = writeConfigForRequest(allowed);
                res.writeHead(200, {
                    'Content-Type': 'application/json',
                    'Access-Control-Allow-Origin': ALLOWED_ORIGIN
                });
                res.end(JSON.stringify({ ok: true, config: publicSettings(saved) }));
            } catch (e) {
                res.writeHead(e.statusCode || 400, { 'Content-Type': 'application/json' });
                res.end(JSON.stringify({ ok: false, error: e.message }));
            }
            return;
        }

        // Provider profiles: simpan/perbarui satu provider (nama + endpoint + key + model)
        if (req.url === '/api/providers' && req.method === 'POST') {
            try {
                const incoming = JSON.parse(await readBody(req));
                saveProviderForRequest(incoming.name, incoming);
                const saved = readConfigForRequest();
                res.writeHead(200, { 'Content-Type': 'application/json', 'Access-Control-Allow-Origin': ALLOWED_ORIGIN });
                res.end(JSON.stringify({ ok: true, config: publicSettings(saved) }));
            } catch (e) {
                res.writeHead(e.statusCode || 400, { 'Content-Type': 'application/json' });
                res.end(JSON.stringify({ ok: false, error: e.message }));
            }
            return;
        }

        // Provider profiles: aktifkan (salin field provider ke setting aktif yang dipakai proxy)
        if (req.url === '/api/providers/activate' && req.method === 'POST') {
            try {
                const incoming = JSON.parse(await readBody(req));
                const saved = activateProviderForRequest(incoming.name);
                res.writeHead(200, { 'Content-Type': 'application/json', 'Access-Control-Allow-Origin': ALLOWED_ORIGIN });
                res.end(JSON.stringify({ ok: true, config: publicSettings(saved) }));
            } catch (e) {
                res.writeHead(e.statusCode || 400, { 'Content-Type': 'application/json' });
                res.end(JSON.stringify({ ok: false, error: e.message }));
            }
            return;
        }

        // Provider profiles: hapus satu provider
        if (req.url === '/api/providers' && req.method === 'DELETE') {
            try {
                const incoming = JSON.parse(await readBody(req));
                deleteProviderForRequest(incoming.name);
                const saved = readConfigForRequest();
                res.writeHead(200, { 'Content-Type': 'application/json', 'Access-Control-Allow-Origin': ALLOWED_ORIGIN });
                res.end(JSON.stringify({ ok: true, config: publicSettings(saved) }));
            } catch (e) {
                res.writeHead(e.statusCode || 400, { 'Content-Type': 'application/json' });
                res.end(JSON.stringify({ ok: false, error: e.message }));
            }
            return;
        }

        // Hermes: daftar provider + model dari konfigurasi Hermes (config.yaml + .env)
        if (req.url === '/api/hermes/providers' && req.method === 'GET') {
            const result = hermesProviders.readHermesProviders();
            res.writeHead(result.ok ? 200 : 404, {
                'Content-Type': 'application/json',
                'Access-Control-Allow-Origin': ALLOWED_ORIGIN,
                'Cache-Control': 'no-store'
            });
            res.end(JSON.stringify(hermesProviders.publicHermesProviders(result)));
            return;
        }

        // Hermes: impor satu provider Hermes menjadi provider Excel Agent lalu aktifkan
        if (req.url === '/api/hermes/import' && req.method === 'POST') {
            try {
                const incoming = JSON.parse(await readBody(req));
                const found = hermesProviders.findHermesProvider(incoming.name);
                if (!found.ok) { const e = new Error(found.error); e.statusCode = 404; throw e; }
                const p = found.provider;
                const wantedModel = typeof incoming.model === 'string' && incoming.model.trim()
                    ? incoming.model.trim()
                    : p.model;
                const providerLabel = 'Hermes: ' + p.name;
                saveProviderForRequest(providerLabel, {
                    apiEndpoint: p.apiEndpoint,
                    model: wantedModel,
                    // Daftar model disimpan milik provider ini saja, tidak dicampur global.
                    models: p.models,
                    apiKey: p.apiKey
                });
                const finalConfig = activateProviderForRequest(providerLabel);
                res.writeHead(200, { 'Content-Type': 'application/json', 'Access-Control-Allow-Origin': ALLOWED_ORIGIN });
                res.end(JSON.stringify({
                    ok: true,
                    providerName: providerLabel,
                    model: wantedModel,
                    models: p.models,
                    config: publicSettings(finalConfig)
                }));
            } catch (e) {
                res.writeHead(e.statusCode || 400, { 'Content-Type': 'application/json' });
                res.end(JSON.stringify({ ok: false, error: e.message }));
            }
            return;
        }

        // Anthropic: status login (OAuth Claude Code / API key) + daftar model
        if (req.url === '/api/anthropic/status' && req.method === 'GET') {
            const config = readConfigForRequest();
            const status = await anthropic.authStatus(config.apiKey);
            res.writeHead(200, {
                'Content-Type': 'application/json',
                'Access-Control-Allow-Origin': ALLOWED_ORIGIN,
                'Cache-Control': 'no-store'
            });
            res.end(JSON.stringify({
                ok: status.ok,
                mode: status.mode,
                expiresAt: status.expiresAt || null,
                models: status.models,
                modelsFromApi: status.modelsFromApi === true,
                credentialsPath: anthropic.credentialsPath(),
                error: status.error || null
            }));
            return;
        }

        // Anthropic: aktifkan sebagai provider aktif (pakai login yang sudah ada)
        if (req.url === '/api/anthropic/activate' && req.method === 'POST') {
            try {
                const incoming = JSON.parse(await readBody(req));
                const status = await anthropic.authStatus(readConfigForRequest().apiKey);
                if (!status.ok) { const e = new Error(status.error || 'Belum login Anthropic.'); e.statusCode = 401; throw e; }
                const model = typeof incoming.model === 'string' && incoming.model.trim()
                    ? incoming.model.trim()
                    : (status.models[0] || 'claude-sonnet-4-5-20250929');
                const providerLabel = 'Anthropic (login)';
                saveProviderForRequest(providerLabel, {
                    apiEndpoint: anthropic.ANTHROPIC_ENDPOINT,
                    model: model,
                    models: status.models
                });
                const finalConfig = activateProviderForRequest(providerLabel);
                res.writeHead(200, { 'Content-Type': 'application/json', 'Access-Control-Allow-Origin': ALLOWED_ORIGIN });
                res.end(JSON.stringify({
                    ok: true,
                    providerName: providerLabel,
                    mode: status.mode,
                    model: model,
                    models: status.models,
                    config: publicSettings(finalConfig)
                }));
            } catch (e) {
                res.writeHead(e.statusCode || 400, { 'Content-Type': 'application/json' });
                res.end(JSON.stringify({ ok: false, error: e.message }));
            }
            return;
        }

        // Model per provider: tambah / hapus model pada satu provider saja
        if (req.url === '/api/providers/models' && (req.method === 'POST' || req.method === 'DELETE')) {
            try {
                const incoming = JSON.parse(await readBody(req));
                const cleanName = sanitizeProviderName(incoming.name);
                if (!cleanName) { const e = new Error('Nama provider tidak valid.'); e.statusCode = 400; throw e; }
                const config = readConfigForRequest();
                const providers = Object.assign({}, config.providers || {});
                const p = providers[cleanName];
                if (!p) { const e = new Error('Provider tidak ditemukan.'); e.statusCode = 404; throw e; }
                const model = typeof incoming.model === 'string' ? incoming.model.trim() : '';
                if (!model) { const e = new Error('Nama model kosong.'); e.statusCode = 400; throw e; }

                const models = Array.isArray(p.models) ? p.models.slice() : [];
                if (req.method === 'POST') {
                    if (models.indexOf(model) === -1) models.push(model);
                } else {
                    const idx = models.indexOf(model);
                    if (idx !== -1) models.splice(idx, 1);
                }
                providers[cleanName] = Object.assign({}, p, { models });
                const saved = writeConfigForRequest({ providers });
                res.writeHead(200, { 'Content-Type': 'application/json', 'Access-Control-Allow-Origin': ALLOWED_ORIGIN });
                res.end(JSON.stringify({ ok: true, name: cleanName, models, config: publicSettings(saved) }));
            } catch (e) {
                res.writeHead(e.statusCode || 400, { 'Content-Type': 'application/json' });
                res.end(JSON.stringify({ ok: false, error: e.message }));
            }
            return;
        }

        // Batalkan eksekusi yang sedang berjalan (dipicu tombol Stop / tombol ESC)
        if (req.url === '/api/cancel' && req.method === 'POST') {
            try {
                let key = '';
                try {
                    const incoming = JSON.parse(await readBody(req));
                    if (incoming && typeof incoming.workbookPath === 'string' && incoming.workbookPath.trim()) {
                        key = workbookHistoryKey(incoming.workbookPath);
                    }
                } catch (e) { /* body kosong = batalkan semua */ }

                const stopped = cancelExecutions(key);
                // Hentikan juga proses OpenCode bila sedang berjalan.
                let opencodeStopped = false;
                if (activeOpencodeCancel.cancel) {
                    try { activeOpencodeCancel.cancel(); opencodeStopped = true; } catch (e) { }
                    activeOpencodeCancel.cancel = null;
                }

                res.writeHead(200, { 'Content-Type': 'application/json', 'Access-Control-Allow-Origin': ALLOWED_ORIGIN });
                res.end(JSON.stringify({
                    ok: true,
                    stopped: stopped,
                    opencodeStopped: opencodeStopped,
                    message: (stopped || opencodeStopped)
                        ? 'Eksekusi dihentikan.'
                        : 'Tidak ada eksekusi yang sedang berjalan.'
                }));
            } catch (e) {
                res.writeHead(500, { 'Content-Type': 'application/json' });
                res.end(JSON.stringify({ ok: false, error: e.message }));
            }
            return;
        }

        if (req.url.startsWith('/api/history?') && req.method === 'GET') {
            try {
                const requestUrl = new URL(req.url, ALLOWED_ORIGIN);
                const history = readWorkbookHistoryForRequest(
                    requestUrl.searchParams.get('workbookPath'),
                    requestUrl.searchParams.get('workbookId')
                );
                res.writeHead(200, { 'Content-Type': 'application/json', 'Access-Control-Allow-Origin': ALLOWED_ORIGIN });
                res.end(JSON.stringify({ ok: true, chatHistory: history }));
            } catch (e) {
                res.writeHead(e.statusCode || 400, { 'Content-Type': 'application/json' });
                res.end(JSON.stringify({ ok: false, error: e.message }));
            }
            return;
        }

        if (req.url === '/api/history' && req.method === 'POST') {
            try {
                const body = JSON.parse(await readBody(req));
                const history = writeWorkbookHistoryForRequest(body.workbookPath, body.chatHistory, body.workbookId);
                res.writeHead(200, { 'Content-Type': 'application/json', 'Access-Control-Allow-Origin': ALLOWED_ORIGIN });
                res.end(JSON.stringify({ ok: true, chatHistory: history }));
            } catch (e) {
                res.writeHead(e.statusCode || 400, { 'Content-Type': 'application/json' });
                res.end(JSON.stringify({ ok: false, error: e.message }));
            }
            return;
        }

        if (req.url === '/api/history' && req.method === 'DELETE') {
            try {
                const body = JSON.parse(await readBody(req));
                deleteWorkbookHistoryForRequest(body.workbookPath, body.workbookId);
                res.writeHead(200, { 'Content-Type': 'application/json', 'Access-Control-Allow-Origin': ALLOWED_ORIGIN });
                res.end(JSON.stringify({ ok: true }));
            } catch (e) {
                res.writeHead(e.statusCode || 400, { 'Content-Type': 'application/json' });
                res.end(JSON.stringify({ ok: false, error: e.message }));
            }
            return;
        }

        // API proxy endpoint
        if (req.url === '/api/proxy' && req.method === 'POST') {
            try {
                const body = await readBody(req);
                await proxyToAPIForRequest(body, res, readConfigForRequest);
            } catch (e) {
                if (!res.headersSent) {
                    res.writeHead(e.statusCode || 400, { 'Content-Type': 'application/json' });
                    res.end(JSON.stringify({ ok: false, error: e.message }));
                }
            }
            return;
        }

        // VBA execution endpoint
        if (req.url === '/api/run-vba' && req.method === 'POST') {
            let body;
            try {
                body = JSON.parse(await readBody(req));
            } catch (e) {
                res.writeHead(e.statusCode || 400, { 'Content-Type': 'application/json' });
                res.end(JSON.stringify({ ok: false, error: e.message }));
                return;
            }
            const vbaCode = body.vbaCode || '';
            if (!vbaCode.trim()) {
                res.writeHead(400, { 'Content-Type': 'application/json' });
                res.end(JSON.stringify({ ok: false, error: 'Kode VBA kosong.' }));
                return;
            }
            try {
                const { execFile } = require('child_process');
                const pythonPath = process.env.PYTHON || 'python';
                const scriptPath = path.join(__dirname, 'execute_vba.py');
                const child = execFile(pythonPath, [scriptPath], { timeout: 30000 }, (err, stdout, stderr) => {
                    unregisterExecution('vba', child);
                    if (child && child.cancelled === true) {
                        res.writeHead(200, { 'Content-Type': 'application/json', 'Access-Control-Allow-Origin': ALLOWED_ORIGIN });
                        res.end(JSON.stringify({ ok: false, cancelled: true, error: 'Eksekusi dihentikan oleh pengguna.' }));
                        return;
                    }
                    if (err) {
                        res.writeHead(500, { 'Content-Type': 'application/json' });
                        res.end(JSON.stringify({ ok: false, error: err.message }));
                        return;
                    }
                    try {
                        const result = JSON.parse(stdout.trim());
                        res.writeHead(200, { 'Content-Type': 'application/json', 'Access-Control-Allow-Origin': ALLOWED_ORIGIN });
                        res.end(JSON.stringify(result));
                    } catch (e) {
                        res.writeHead(500, { 'Content-Type': 'application/json' });
                        res.end(JSON.stringify({ ok: false, error: 'Gagal parse hasil: ' + stdout }));
                    }
                });
                registerExecution('vba', child);
                child.stdin.write(vbaCode);
                child.stdin.end();
            } catch (e) {
                res.writeHead(500, { 'Content-Type': 'application/json' });
                res.end(JSON.stringify({ ok: false, error: e.message }));
            }
            return;
        }

        // OpenCode collaboration endpoint: prompt -> `opencode run` -> jawaban mendalam
        if (req.url === '/api/opencode' && req.method === 'POST') {
            if (opencodeRunning) {
                res.writeHead(429, { 'Content-Type': 'application/json' });
                res.end(JSON.stringify({ ok: false, error: 'OpenCode sedang memproses permintaan lain. Tunggu sebentar.' }));
                return;
            }
            let body;
            try {
                body = JSON.parse(await readBody(req, MAX_OPENCODE_BODY_BYTES));
            } catch (e) {
                res.writeHead(e.statusCode || 400, { 'Content-Type': 'application/json' });
                res.end(JSON.stringify({ ok: false, error: e.message }));
                return;
            }
            const prompt = typeof body.prompt === 'string' ? body.prompt : '';
            if (!prompt.trim()) {
                res.writeHead(400, { 'Content-Type': 'application/json' });
                res.end(JSON.stringify({ ok: false, error: 'Prompt kosong.' }));
                return;
            }
            const model = typeof body.model === 'string' && body.model.trim() ? body.model.trim() : '';
            const sessionId = typeof body.session === 'string' && body.session.trim() ? body.session.trim() : '';
            let preparedImages;
            try {
                preparedImages = prepareOpencodeImages(body.images);
            } catch (e) {
                res.writeHead(400, { 'Content-Type': 'application/json' });
                res.end(JSON.stringify({ ok: false, error: e.message }));
                return;
            }
            // Default working directory netral (tanpa .mcp.json/proyek) agar opencode
            // tidak menghabiskan waktu memulai MCP server yang tidak dibutuhkan.
            const dir = typeof body.workdir === 'string' && body.workdir.trim() ? body.workdir.trim() : CONFIG_DIR;

            let progressKey;
            try {
                progressKey = body.target && typeof body.target.workbookPath === 'string' && body.target.workbookPath.trim()
                    ? workbookHistoryKey(body.target.workbookPath)
                    : crypto.createHash('sha256').update('opencode-general').digest('hex');
            } catch (e) {
                progressKey = crypto.createHash('sha256').update('opencode-general').digest('hex');
            }
            const progressFile = path.join(CONFIG_DIR, 'progress_' + progressKey + '.txt');

            opencodeRunning = true;
            const cancelRef = {};
            // Ekspos ke /api/cancel supaya ESC juga menghentikan OpenCode.
            activeOpencodeCancel.cancel = () => { if (cancelRef.cancel) cancelRef.cancel(); };
            // Bila client (task pane) menutup koneksi, hentikan proses opencode
            // segera agar tidak terus berjalan dan mengunci slot (429).
            res.on('close', () => {
                if (cancelRef.cancel) cancelRef.cancel();
            });
            try {
                let result;
                try {
                    result = await runOpencode(prompt, model, sessionId, dir, progressFile, cancelRef, preparedImages.paths);
                } catch (firstError) {
                    // Gateway sesekali me-reset route model ("reset after 30s") atau
                    // gagal sementara. Tunggu sebentar lalu coba sekali lagi; bila
                    // sedang memakai session, ulang sebagai session baru.
                    // KECUALI bila gagal karena timeout: percobaan kedua hanya akan
                    // menambah 7 menit sia-sia dan melewati watchdog client (440 detik).
                    if (/timeout/i.test(firstError.message)) throw firstError;
                    if (res.writableEnded) throw firstError; // client sudah pergi
                    await new Promise(r => setTimeout(r, 5000));
                    if (res.writableEnded) throw firstError;
                    try {
                        result = await runOpencode(prompt, model, '', dir, progressFile, cancelRef, preparedImages.paths);
                        if (sessionId) result.sessionReset = true;
                    } catch (secondError) {
                        throw new Error(firstError.message + ' | Percobaan kedua: ' + secondError.message);
                    }
                }
                try { fs.unlinkSync(progressFile); } catch (e) { }
                res.writeHead(200, {
                    'Content-Type': 'application/json',
                    'Access-Control-Allow-Origin': ALLOWED_ORIGIN
                });
                res.end(JSON.stringify(result));
            } catch (e) {
                try { fs.unlinkSync(progressFile); } catch (e2) { }
                res.writeHead(500, { 'Content-Type': 'application/json' });
                res.end(JSON.stringify({ ok: false, error: e.message }));
            } finally {
                preparedImages.cleanup();
                opencodeRunning = false;
                activeOpencodeCancel.cancel = null;
            }
            return;
        }

        // Progress polling endpoint
        if (req.url.startsWith('/api/exec-progress') && req.method === 'GET') {
            const params = new URL(req.url, `http://localhost:${PORT}`).searchParams;
            const wp = params.get('workbookPath') || '';
            if (!wp) {
                res.writeHead(400, { 'Content-Type': 'application/json' });
                res.end(JSON.stringify({ progress: '' }));
                return;
            }
            try {
                const key = workbookHistoryKey(wp);
                const progressFile = path.join(CONFIG_DIR, 'progress_' + key + '.txt');
                let progress = '';
                try { progress = fs.readFileSync(progressFile, 'utf8').trim(); } catch (e) { /* file belum ada */ }
                res.writeHead(200, { 'Content-Type': 'application/json', 'Access-Control-Allow-Origin': ALLOWED_ORIGIN });
                res.end(JSON.stringify({ progress }));
            } catch (e) {
                res.writeHead(400, { 'Content-Type': 'application/json' });
                res.end(JSON.stringify({ progress: '' }));
            }
            return;
        }

        // MCP: Get workbook metadata
        if (req.url === '/api/mcp/workbook-info' && req.method === 'POST') {
            try {
                const body = JSON.parse(await readBody(req));
                let filepath = body.filepath;
                if (!filepath) {
                    res.writeHead(400, { 'Content-Type': 'application/json' });
                    res.end(JSON.stringify({ ok: false, error: 'Filepath diperlukan.' }));
                    return;
                }
                // Jika URL OneDrive, coba konversi ke path lokal sync folder
                if (/^https?:\/\//i.test(filepath)) {
                    const localPath = resolveOneDriveLocalPath(filepath);
                    if (localPath) {
                        filepath = localPath;
                    } else {
                        res.writeHead(400, { 'Content-Type': 'application/json' });
                        res.end(JSON.stringify({ ok: false, error: 'File OneDrive tidak ditemukan di sync folder lokal. Pastikan file sudah di-sync ke komputer.' }));
                        return;
                    }
                } else if (!path.win32.isAbsolute(filepath)) {
                    res.writeHead(400, { 'Content-Type': 'application/json' });
                    res.end(JSON.stringify({ ok: false, error: 'Filepath tidak valid atau bukan path lokal.' }));
                    return;
                }
                const mcp = getExcelMcpClient();
                const result = await mcp.callTool('get_workbook_metadata', { filepath, include_ranges: true });
                res.writeHead(200, { 'Content-Type': 'application/json', 'Access-Control-Allow-Origin': ALLOWED_ORIGIN });
                res.end(JSON.stringify({ ok: true, data: result }));
            } catch (e) {
                res.writeHead(500, { 'Content-Type': 'application/json' });
                res.end(JSON.stringify({ ok: false, error: e.message }));
            }
            return;
        }

        // MCP: Read data from sheet
        if (req.url === '/api/mcp/read-data' && req.method === 'POST') {
            try {
                const body = JSON.parse(await readBody(req));
                const { sheet_name, start_cell, end_cell, preview_only } = body;
                let filepath = body.filepath;
                if (!filepath) {
                    res.writeHead(400, { 'Content-Type': 'application/json' });
                    res.end(JSON.stringify({ ok: false, error: 'Filepath diperlukan.' }));
                    return;
                }
                // Jika URL OneDrive, coba konversi ke path lokal sync folder
                if (/^https?:\/\//i.test(filepath)) {
                    const localPath = resolveOneDriveLocalPath(filepath);
                    if (localPath) {
                        filepath = localPath;
                    } else {
                        res.writeHead(400, { 'Content-Type': 'application/json' });
                        res.end(JSON.stringify({ ok: false, error: 'File OneDrive tidak ditemukan di sync folder lokal. Pastikan file sudah di-sync ke komputer.' }));
                        return;
                    }
                } else if (!path.win32.isAbsolute(filepath)) {
                    res.writeHead(400, { 'Content-Type': 'application/json' });
                    res.end(JSON.stringify({ ok: false, error: 'Filepath tidak valid atau bukan path lokal.' }));
                    return;
                }
                if (!sheet_name) {
                    res.writeHead(400, { 'Content-Type': 'application/json' });
                    res.end(JSON.stringify({ ok: false, error: 'Nama sheet diperlukan.' }));
                    return;
                }
                const mcp = getExcelMcpClient();
                const args = { filepath, sheet_name };
                if (start_cell) args.start_cell = start_cell;
                if (end_cell) args.end_cell = end_cell;
                if (preview_only !== undefined) args.preview_only = preview_only;
                const result = await mcp.callTool('read_data_from_excel', args);
                res.writeHead(200, { 'Content-Type': 'application/json', 'Access-Control-Allow-Origin': ALLOWED_ORIGIN });
                res.end(JSON.stringify({ ok: true, data: result }));
            } catch (e) {
                res.writeHead(500, { 'Content-Type': 'application/json' });
                res.end(JSON.stringify({ ok: false, error: e.message }));
            }
            return;
        }

        // Python execution endpoint (win32com direct)
        if (req.url === '/api/exec-python' && req.method === 'POST') {
            let envelope;
            try {
                envelope = validateExecutionEnvelope(JSON.parse(await readBody(req)));
            } catch (e) {
                const targetError = e instanceof ExecutionValidationError && e.targetError === true;
                res.writeHead(e.statusCode || 400, { 'Content-Type': 'application/json' });
                res.end(JSON.stringify(targetError
                    ? { ok: false, targetError: true, error: e.message }
                    : { ok: false, error: e.message }));
                return;
            }
            // Hitung path file progress dan sertakan di envelope
            const progressKey = workbookHistoryKey(envelope.target.workbookPath);
            const progressFile = path.join(CONFIG_DIR, 'progress_' + progressKey + '.txt');
            envelope.progressFile = progressFile;
            try {
                const result = await enqueueCom(() => new Promise((resolve, reject) => {
                    const pythonPath = process.env.PYTHON || 'python';
                    const scriptPath = path.join(__dirname, 'exec_python.py');
                    const child = execFileForRequest(pythonPath, [scriptPath], { timeout: EXEC_TIMEOUT_MS }, (err, stdout) => {
                        unregisterExecution(progressKey, child);
                        // Dihentikan user lewat ESC: bukan error, tapi status batal.
                        if (child && child.cancelled === true) {
                            return resolve({ ok: false, cancelled: true, error: 'Eksekusi dihentikan oleh pengguna.' });
                        }
                        if (err) return reject(err);
                        try {
                            resolve(JSON.parse(stdout.trim()));
                        } catch (e) {
                            reject(new Error('Gagal parse hasil: ' + stdout));
                        }
                    });
                    // Catat agar bisa dihentikan lewat /api/cancel.
                    registerExecution(progressKey, child);
                    child.stdin.on('error', reject);
                    child.stdin.write(JSON.stringify(envelope));
                    child.stdin.end();
                }));
                // Bersihkan file progress setelah eksekusi selesai
                try { fs.unlinkSync(progressFile); } catch (e) { /* sudah tidak ada */ }
                res.writeHead(result && result.targetError === true ? 409 : 200, {
                    'Content-Type': 'application/json',
                    'Access-Control-Allow-Origin': ALLOWED_ORIGIN
                });
                res.end(JSON.stringify(result));
            } catch (e) {
                try { fs.unlinkSync(progressFile); } catch (e2) { /* sudah tidak ada */ }
                res.writeHead(e.statusCode || 500, { 'Content-Type': 'application/json' });
                res.end(JSON.stringify({ ok: false, error: e.message }));
            }
            return;
        }

        // Static files
        const filePath = resolveFilePath(req.url);
        if (!filePath) {
            res.writeHead(404);
            res.end('File not found');
            return;
        }
        const ext = path.extname(filePath);
        const contentType = MIME_TYPES[ext] || 'application/octet-stream';

        fs.readFile(filePath, (err, content) => {
            if (err) {
                res.writeHead(err.code === 'ENOENT' ? 404 : 500);
                res.end(err.code === 'ENOENT' ? 'File not found' : 'Server error');
                return;
            }
            res.writeHead(200, {
                'Content-Type': contentType,
                'Cache-Control': 'no-store, no-cache, must-revalidate, max-age=0',
                'Pragma': 'no-cache',
                'Expires': '0'
            });
            res.end(content);
        });
    };
}

async function start() {
    // Migrasi skema model-per-provider (aman diulang, tidak menghapus data lama).
    try {
        const result = migrateProviderModels();
        if (result.migrated) {
            console.log(`Migrasi: ${result.migrated} provider kini punya daftar model sendiri (modelList global tetap disimpan).`);
        }
    } catch (e) {
        console.error('Migrasi daftar model per provider dilewati:', e.message);
    }

    const httpsOptions = await devCerts.getHttpsServerOptions();
    const server = https.createServer(httpsOptions, createRequestHandler());

    server.on('error', (err) => {
        if (err.code === 'EADDRINUSE') {
            // Instance lain sudah melayani port ini - keluar dengan tenang,
            // bukan crash, supaya launcher ganda tidak menumpuk error.
            console.log(`Port ${PORT} sudah dipakai instance Excel Agent lain. Instance ini berhenti.`);
            process.exit(0);
        }
        console.error('Gagal start server:', err.message);
        console.error('Jalankan dulu: npx office-addin-dev-certs install');
        process.exit(1);
    });

    server.listen(PORT, SERVER_HOST, () => {
        console.log(`Excel Agent server running at https://localhost:${PORT}`);
        console.log(`API proxy aktif di https://localhost:${PORT}/api/proxy`);
        console.log(`Config tersimpan di ${CONFIG_FILE}`);
        console.log('');
        console.log('Sideload add-in:');
        console.log('  npm run sideload');
    });
}

module.exports = {
  cleanHistory,
  createRequestHandler,
  enqueueCom,
    isAllowedOrigin,
    normalizeWorkbookPath,
    SERVER_HOST,
  validateExecutionEnvelope,
  workbookHistoryKey,
  workbookIdKey,
  publicSettings,
  resolveProxyApiKey,
  saveProvider,
  deleteProvider,
  activateProvider,
  sanitizeProviderName,
  runOpencode,
  MAX_REQUEST_BODY_BYTES,
  MAX_PENDING_COM,
  PROXY_TIMEOUT_MS,
  EXEC_TIMEOUT_MS,
  OPENCODE_TIMEOUT_MS,
  MAX_OPENCODE_IMAGES,
  MAX_IMAGE_BYTES,
  MAX_OPENCODE_BODY_BYTES,
  prepareOpencodeImages,
  OPENCODE_MODEL_RE,
  OPENCODE_SESSION_RE,
  readBody,
  checkRouterHealth,
  checkPythonHealth,
  hermesProviders,
  migrateProviderModels,
  registerExecution,
  unregisterExecution,
  cancelExecutions,
  anthropic,
  proxyToAnthropic
};

if (require.main === module) {
    start().catch((err) => {
        console.error('Gagal start server:', err.message);
        console.error('Jalankan dulu: npx office-addin-dev-certs install');
        process.exit(1);
    });
}

process.on('uncaughtException', (err) => {
    console.error('Uncaught exception (server masih jalan):', err.message);
});
