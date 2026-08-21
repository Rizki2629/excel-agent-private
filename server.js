const https = require('https');
const http = require('http');
const fs = require('fs');
const os = require('os');
const path = require('path');
const crypto = require('crypto');
const devCerts = require('office-addin-dev-certs');
const { getExcelMcpClient } = require('./mcp_client');

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
    const { chatHistory, chatHistoryByWorkbook, ...publicConfig } = config;
    delete publicConfig.apiKey;
    publicConfig.hasApiKey = typeof config.apiKey === 'string' && config.apiKey.length > 0;
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

function resolveProxyApiKey(requestApiKey, storedApiKey) {
    return typeof requestApiKey === 'string' && requestApiKey.length > 0 ? requestApiKey : storedApiKey;
}

function deleteWorkbookHistory(workbookPath) {
    const config = readConfig();
    const histories = Object.assign({}, config.chatHistoryByWorkbook || {});
    delete histories[workbookHistoryKey(workbookPath)];
    writeConfig({ chatHistoryByWorkbook: histories });
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

// Proxy request ke API eksternal (menghindari CORS)
async function proxyToAPI(reqBody, res, readConfigFn = readConfig) {
    let proxyReq = null;
    try {
        const { apiEndpoint, apiKey, payload } = JSON.parse(reqBody);
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

    const readWorkbookHistoryForRequest = (workbookPath) => {
        const config = readConfigForRequest();
        return (config.chatHistoryByWorkbook || {})[workbookHistoryKey(workbookPath)] || [];
    };
    const writeWorkbookHistoryForRequest = (workbookPath, history) => {
        const config = readConfigForRequest();
        const histories = Object.assign({}, config.chatHistoryByWorkbook || {});
        histories[workbookHistoryKey(workbookPath)] = cleanHistory(history);
        writeConfigForRequest({ chatHistoryByWorkbook: histories });
        return histories[workbookHistoryKey(workbookPath)];
    };
    const deleteWorkbookHistoryForRequest = (workbookPath) => {
        const config = readConfigForRequest();
        const histories = Object.assign({}, config.chatHistoryByWorkbook || {});
        delete histories[workbookHistoryKey(workbookPath)];
        writeConfigForRequest({ chatHistoryByWorkbook: histories });
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
                ['apiEndpoint', 'model', 'autoRun', 'useOpencode', 'opencodeModel'].forEach((k) => { if (incoming[k] !== undefined) allowed[k] = incoming[k]; });
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

        if (req.url.startsWith('/api/history?') && req.method === 'GET') {
            try {
                const requestUrl = new URL(req.url, ALLOWED_ORIGIN);
                const history = readWorkbookHistoryForRequest(requestUrl.searchParams.get('workbookPath'));
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
                const history = writeWorkbookHistoryForRequest(body.workbookPath, body.chatHistory);
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
                deleteWorkbookHistoryForRequest(body.workbookPath);
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
                        if (err) return reject(err);
                        try {
                            resolve(JSON.parse(stdout.trim()));
                        } catch (e) {
                            reject(new Error('Gagal parse hasil: ' + stdout));
                        }
                    });
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
  publicSettings,
  resolveProxyApiKey,
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
  readBody
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
