'use strict';

// Adapter Anthropic untuk Excel Agent.
//
// Tujuan: memakai login Anthropic (OAuth Claude Code atau API key) tanpa mengubah
// task pane. Task pane bicara format OpenAI chat/completions, sedangkan
// api.anthropic.com memakai format Messages API, jadi modul ini menerjemahkan
// permintaan dan respons (termasuk streaming SSE) di kedua arah.

const fs = require('fs');
const path = require('path');
const os = require('os');
const https = require('https');

const ANTHROPIC_HOST = 'api.anthropic.com';
const ANTHROPIC_VERSION = '2023-06-01';
// Beta header wajib saat memakai token OAuth Claude Code.
const OAUTH_BETA = 'oauth-2025-04-20';
// client_id publik Claude Code (dipakai untuk refresh token).
const CLAUDE_CODE_CLIENT_ID = '9d1c250a-e61b-44d9-88ed-5944d1962f5e';
const TOKEN_URL = 'https://console.anthropic.com/v1/oauth/token';
// System prompt identitas Claude Code wajib ada saat auth OAuth.
const CLAUDE_CODE_IDENTITY = "You are Claude Code, Anthropic's official CLI for Claude.";
// Lantai max_tokens: model thinking bisa memakai ribuan token hanya untuk
// penalaran internal, sehingga nilai kecil menghasilkan jawaban kosong.
const MIN_MAX_TOKENS = 2048;

// Endpoint sintetis: dipakai task pane sebagai apiEndpoint agar server tahu
// permintaan harus diarahkan ke adapter ini.
const ANTHROPIC_ENDPOINT = 'https://api.anthropic.com/v1/messages';

function isAnthropicEndpoint(endpoint) {
    if (typeof endpoint !== 'string' || !endpoint) return false;
    try {
        return new URL(endpoint).hostname.toLowerCase() === ANTHROPIC_HOST;
    } catch (e) {
        return false;
    }
}

function credentialsPath() {
    return path.join(os.homedir(), '.claude', '.credentials.json');
}

// Baca kredensial OAuth Claude Code dari disk.
function readOauthCredentials(file = credentialsPath()) {
    let raw;
    try {
        raw = fs.readFileSync(file, 'utf8');
    } catch (e) {
        return null;
    }
    let parsed;
    try {
        parsed = JSON.parse(raw);
    } catch (e) {
        return null;
    }
    const o = parsed && parsed.claudeAiOauth;
    if (!o || typeof o.accessToken !== 'string' || !o.accessToken) return null;
    return {
        accessToken: o.accessToken,
        refreshToken: typeof o.refreshToken === 'string' ? o.refreshToken : '',
        expiresAt: typeof o.expiresAt === 'number' ? o.expiresAt : 0,
        raw: parsed
    };
}

function isExpired(creds, skewMs = 120000) {
    if (!creds || !creds.expiresAt) return false;
    return Date.now() + skewMs >= creds.expiresAt;
}

function writeOauthCredentials(creds, tokenResponse, file = credentialsPath()) {
    const next = Object.assign({}, creds.raw || {});
    const block = Object.assign({}, next.claudeAiOauth || {});
    block.accessToken = tokenResponse.access_token;
    if (tokenResponse.refresh_token) block.refreshToken = tokenResponse.refresh_token;
    if (tokenResponse.expires_in) block.expiresAt = Date.now() + (tokenResponse.expires_in * 1000);
    next.claudeAiOauth = block;
    fs.writeFileSync(file, JSON.stringify(next, null, 2), 'utf8');
    try { fs.chmodSync(file, 0o600); } catch (e) { /* best effort */ }
    return block;
}

// Tukar refresh token menjadi access token baru.
function refreshAccessToken(refreshToken) {
    return new Promise((resolve, reject) => {
        if (!refreshToken) return reject(new Error('Refresh token Anthropic tidak tersedia.'));
        const body = JSON.stringify({
            grant_type: 'refresh_token',
            refresh_token: refreshToken,
            client_id: CLAUDE_CODE_CLIENT_ID
        });
        const url = new URL(TOKEN_URL);
        const req = https.request({
            hostname: url.hostname,
            path: url.pathname,
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'Content-Length': Buffer.byteLength(body)
            },
            timeout: 30000
        }, (res) => {
            const chunks = [];
            res.on('data', c => chunks.push(c));
            res.on('end', () => {
                const text = Buffer.concat(chunks).toString();
                if (res.statusCode < 200 || res.statusCode >= 300) {
                    return reject(new Error('Refresh token Anthropic gagal (' + res.statusCode + '): ' + text.slice(0, 200)));
                }
                try {
                    resolve(JSON.parse(text));
                } catch (e) {
                    reject(new Error('Respons refresh token Anthropic tidak valid.'));
                }
            });
        });
        req.on('timeout', () => req.destroy(new Error('Timeout saat refresh token Anthropic.')));
        req.on('error', reject);
        req.end(body);
    });
}

// Tentukan kredensial yang dipakai: API key eksplisit, atau OAuth Claude Code.
// storedApiKey dipakai bila user memang menyimpan sk-ant-api... sendiri.
async function resolveAuth(storedApiKey) {
    const key = typeof storedApiKey === 'string' ? storedApiKey.trim() : '';
    // Token OAuth (sk-ant-oat...) tidak boleh dikirim sebagai x-api-key.
    if (key && key.indexOf('sk-ant-api') === 0) {
        return { mode: 'api_key', apiKey: key };
    }
    let creds = readOauthCredentials();
    if (!creds) {
        if (key) return { mode: 'api_key', apiKey: key };
        throw new Error('Belum login Anthropic. Jalankan "claude login" (Claude Code) atau isi API key sk-ant-api... di Pengaturan.');
    }
    if (isExpired(creds) && creds.refreshToken) {
        const tokenResponse = await refreshAccessToken(creds.refreshToken);
        const block = writeOauthCredentials(creds, tokenResponse);
        creds = {
            accessToken: block.accessToken,
            refreshToken: block.refreshToken,
            expiresAt: block.expiresAt,
            raw: creds.raw
        };
    }
    return { mode: 'oauth', accessToken: creds.accessToken, expiresAt: creds.expiresAt };
}

function authHeaders(auth) {
    const headers = {
        'Content-Type': 'application/json',
        'anthropic-version': ANTHROPIC_VERSION
    };
    if (auth.mode === 'oauth') {
        headers['Authorization'] = 'Bearer ' + auth.accessToken;
        headers['anthropic-beta'] = OAUTH_BETA;
    } else {
        headers['x-api-key'] = auth.apiKey;
    }
    return headers;
}

function textFromContent(content) {
    if (typeof content === 'string') return content;
    if (!Array.isArray(content)) return '';
    const parts = [];
    for (const item of content) {
        if (!item) continue;
        if (typeof item === 'string') parts.push(item);
        else if (item.type === 'text' && typeof item.text === 'string') parts.push(item.text);
    }
    return parts.join('\n');
}

// Ubah blok konten OpenAI (teks + image_url) menjadi blok konten Anthropic.
function toAnthropicContent(content) {
    if (typeof content === 'string') return content;
    if (!Array.isArray(content)) return '';
    const blocks = [];
    for (const item of content) {
        if (!item) continue;
        if (typeof item === 'string') {
            blocks.push({ type: 'text', text: item });
        } else if (item.type === 'text' && typeof item.text === 'string') {
            blocks.push({ type: 'text', text: item.text });
        } else if (item.type === 'image_url' && item.image_url && typeof item.image_url.url === 'string') {
            const m = item.image_url.url.match(/^data:([^;]+);base64,(.+)$/);
            if (m) {
                blocks.push({ type: 'image', source: { type: 'base64', media_type: m[1], data: m[2] } });
            }
        }
    }
    return blocks.length ? blocks : '';
}

// Konversi payload OpenAI chat/completions -> Anthropic Messages API.
function toAnthropicRequest(payload, authMode) {
    const messages = Array.isArray(payload.messages) ? payload.messages : [];
    const systemParts = [];
    const converted = [];

    for (const msg of messages) {
        if (!msg || !msg.role) continue;
        if (msg.role === 'system') {
            const text = textFromContent(msg.content);
            if (text) systemParts.push(text);
            continue;
        }
        const role = msg.role === 'assistant' ? 'assistant' : 'user';
        const content = toAnthropicContent(msg.content);
        if (!content || (Array.isArray(content) && !content.length)) continue;
        // Anthropic menolak dua pesan berurutan dengan role sama: gabungkan.
        const last = converted[converted.length - 1];
        if (last && last.role === role) {
            const lastBlocks = Array.isArray(last.content) ? last.content : [{ type: 'text', text: String(last.content) }];
            const nextBlocks = Array.isArray(content) ? content : [{ type: 'text', text: String(content) }];
            last.content = lastBlocks.concat(nextBlocks);
        } else {
            converted.push({ role, content });
        }
    }

    // Anthropic mewajibkan minimal satu pesan, dan harus dimulai dari user.
    if (!converted.length) converted.push({ role: 'user', content: 'Lanjutkan.' });
    if (converted[0].role !== 'user') converted.unshift({ role: 'user', content: 'Lanjutkan.' });

    // Auth OAuth mewajibkan identitas Claude Code sebagai system prompt pertama.
    if (authMode === 'oauth') {
        const hasIdentity = systemParts.some(s => s.indexOf('You are Claude Code') !== -1);
        if (!hasIdentity) systemParts.unshift(CLAUDE_CODE_IDENTITY);
    }

    const request = {
        model: payload.model,
        // Model thinking memakai sebagian kuota untuk penalaran internal, jadi
        // beri lantai yang cukup agar masih ada ruang untuk jawaban teks.
        max_tokens: typeof payload.max_tokens === 'number' && payload.max_tokens > 0
            ? Math.max(payload.max_tokens, MIN_MAX_TOKENS)
            : 8192,
        messages: converted,
        stream: payload.stream === true
    };
    if (systemParts.length) {
        request.system = systemParts.map(text => ({ type: 'text', text }));
    }
    // Lewati parameter sampling bila model ini sudah diketahui menolaknya.
    if (!rejectsSampling(payload.model)) {
        if (typeof payload.temperature === 'number') request.temperature = payload.temperature;
        if (typeof payload.top_p === 'number') request.top_p = payload.top_p;
    }

    if (payload.thinking && typeof payload.thinking === 'object') {
        request.thinking = payload.thinking;
    }
    if (Array.isArray(payload.stop) && payload.stop.length) request.stop_sequences = payload.stop;
    return request;
}

function mapStopReason(reason) {
    if (reason === 'end_turn' || reason === 'stop_sequence') return 'stop';
    if (reason === 'max_tokens') return 'length';
    return reason || 'stop';
}

// Sebagian model Claude generasi baru (mis. claude-opus-5) menolak parameter
// sampling: "`temperature` is deprecated for this model." Daftar model yang
// menolak tidak seragam dan bisa berubah, jadi kita catat hasil temuan di
// memori supaya request berikutnya langsung dikirim tanpa parameter tersebut.
const SAMPLING_FIELDS = ['temperature', 'top_p', 'top_k'];
const modelsRejectingSampling = Object.create(null);

function rejectsSampling(model) {
    return modelsRejectingSampling[model] === true;
}

function rememberRejectsSampling(model) {
    if (model) modelsRejectingSampling[model] = true;
}

// Deteksi pesan error upstream yang menandakan parameter sampling ditolak.
function isSamplingUnsupportedError(message) {
    if (typeof message !== 'string' || !message) return false;
    const lower = message.toLowerCase();
    const mentionsField = SAMPLING_FIELDS.some(f => lower.indexOf(f) !== -1);
    if (!mentionsField) return false;
    return lower.indexOf('deprecated') !== -1
        || lower.indexOf('not supported') !== -1
        || lower.indexOf('unsupported') !== -1
        || lower.indexOf('cannot be used') !== -1
        || lower.indexOf('not permitted') !== -1;
}

// Salinan request tanpa parameter sampling.
function stripSamplingParams(request) {
    const next = Object.assign({}, request);
    let removed = false;
    for (const field of SAMPLING_FIELDS) {
        if (next[field] !== undefined) {
            delete next[field];
            removed = true;
        }
    }
    return { request: next, removed };
}

// Konversi respons non-streaming Anthropic -> bentuk OpenAI.
function toOpenAiResponse(anthropicBody) {
    const text = textFromContent(anthropicBody && anthropicBody.content);
    const usage = (anthropicBody && anthropicBody.usage) || {};
    const stopReason = anthropicBody && anthropicBody.stop_reason;
    // Model "thinking" (mis. claude-opus-5) bisa menghabiskan seluruh max_tokens
    // untuk penalaran internal sehingga tidak ada blok teks sama sekali.
    // Beri pesan jelas daripada mengembalikan jawaban kosong yang membingungkan.
    let content = text;
    if (!content.trim() && stopReason === 'max_tokens') {
        content = '[Jawaban terpotong: seluruh kuota token terpakai untuk penalaran internal model. Naikkan max_tokens atau pilih model non-thinking.]';
    }
    return {
        id: (anthropicBody && anthropicBody.id) || 'msg',
        object: 'chat.completion',
        created: Math.floor(Date.now() / 1000),
        model: (anthropicBody && anthropicBody.model) || '',
        choices: [{
            index: 0,
            message: { role: 'assistant', content: content },
            finish_reason: mapStopReason(stopReason)
        }],
        usage: {
            prompt_tokens: usage.input_tokens || 0,
            completion_tokens: usage.output_tokens || 0,
            total_tokens: (usage.input_tokens || 0) + (usage.output_tokens || 0)
        }
    };
}

// Penerjemah SSE Anthropic -> SSE OpenAI, dipakai streaming ke task pane.
function createSseTranslator(model, write) {
    let buffer = '';
    let currentEvent = '';
    let sentDone = false;
    // index blok -> jenis blok ('text' | 'thinking' | ...)
    const blockTypes = Object.create(null);

    function emitDelta(text) {
        write('data: ' + JSON.stringify({
            id: 'chatcmpl-anthropic',
            object: 'chat.completion.chunk',
            created: Math.floor(Date.now() / 1000),
            model: model,
            choices: [{ index: 0, delta: { content: text }, finish_reason: null }]
        }) + '\n\n');
    }

    function emitFinish(reason) {
        write('data: ' + JSON.stringify({
            id: 'chatcmpl-anthropic',
            object: 'chat.completion.chunk',
            created: Math.floor(Date.now() / 1000),
            model: model,
            choices: [{ index: 0, delta: {}, finish_reason: mapStopReason(reason) }]
        }) + '\n\n');
    }

    function emitDone() {
        if (sentDone) return;
        sentDone = true;
        write('data: [DONE]\n\n');
    }

    function handleData(dataText) {
        let obj;
        try {
            obj = JSON.parse(dataText);
        } catch (e) {
            return;
        }
        const type = obj.type || currentEvent;
        // Catat jenis tiap blok konten. Blok "thinking" adalah penalaran
        // internal model dan TIDAK boleh ikut ditampilkan ke pengguna.
        if (type === 'content_block_start') {
            const bt = (obj.content_block && obj.content_block.type) || 'text';
            blockTypes[obj.index] = bt;
            return;
        }
        if (type === 'content_block_delta') {
            const delta = obj.delta || {};
            const bt = blockTypes[obj.index] || 'text';
            const isThinking = bt === 'thinking' || bt === 'redacted_thinking' ||
                delta.type === 'thinking_delta' || delta.type === 'signature_delta';
            if (isThinking) return; // buang, jangan tampilkan
            if (typeof delta.text === 'string' && delta.text) emitDelta(delta.text);
            return;
        }
        if (type === 'message_delta') {
            const stop = obj.delta && obj.delta.stop_reason;
            if (stop) emitFinish(stop);
            return;
        }
        if (type === 'message_stop') {
            emitDone();
            return;
        }
        if (type === 'error') {
            const message = (obj.error && obj.error.message) || 'Anthropic stream error';
            write('data: ' + JSON.stringify({ error: { message: message, type: (obj.error && obj.error.type) || 'api_error' } }) + '\n\n');
            emitDone();
        }
    }

    return {
        push(chunk) {
            buffer += chunk.toString().replace(/\r\n/g, '\n');
            const lines = buffer.split('\n');
            buffer = lines.pop();
            for (const line of lines) {
                const trimmed = line.trim();
                if (!trimmed) { currentEvent = ''; continue; }
                if (trimmed.indexOf('event:') === 0) {
                    currentEvent = trimmed.substring(6).trim();
                } else if (trimmed.indexOf('data:') === 0) {
                    handleData(trimmed.substring(5).trim());
                }
            }
        },
        end() {
            if (buffer.trim().indexOf('data:') === 0) handleData(buffer.trim().substring(5).trim());
            buffer = '';
            emitDone();
        }
    };
}

// Daftar model Anthropic dari API (butuh auth), dengan fallback daftar statis.
const FALLBACK_MODELS = [
    'claude-opus-4-6-20260219',
    'claude-sonnet-4-6-20260219',
    'claude-opus-4-1-20250805',
    'claude-sonnet-4-5-20250929',
    'claude-haiku-4-5-20251001',
    'claude-3-7-sonnet-20250219',
    'claude-3-5-haiku-20241022'
];

function listModels(auth, timeoutMs = 15000) {
    return new Promise((resolve) => {
        const headers = authHeaders(auth);
        delete headers['Content-Type'];
        const req = https.request({
            hostname: ANTHROPIC_HOST,
            path: '/v1/models?limit=100',
            method: 'GET',
            headers,
            timeout: timeoutMs
        }, (res) => {
            const chunks = [];
            res.on('data', c => chunks.push(c));
            res.on('end', () => {
                if (res.statusCode < 200 || res.statusCode >= 300) {
                    return resolve({ ok: false, models: FALLBACK_MODELS.slice(), fallback: true, statusCode: res.statusCode });
                }
                try {
                    const parsed = JSON.parse(Buffer.concat(chunks).toString());
                    const ids = (parsed.data || []).map(m => m && m.id).filter(Boolean);
                    resolve(ids.length
                        ? { ok: true, models: ids, fallback: false }
                        : { ok: false, models: FALLBACK_MODELS.slice(), fallback: true });
                } catch (e) {
                    resolve({ ok: false, models: FALLBACK_MODELS.slice(), fallback: true });
                }
            });
        });
        req.on('timeout', () => { req.destroy(); resolve({ ok: false, models: FALLBACK_MODELS.slice(), fallback: true }); });
        req.on('error', () => resolve({ ok: false, models: FALLBACK_MODELS.slice(), fallback: true }));
        req.end();
    });
}

// Status login untuk ditampilkan di task pane.
async function authStatus(storedApiKey) {
    try {
        const auth = await resolveAuth(storedApiKey);
        const models = await listModels(auth);
        return {
            ok: true,
            mode: auth.mode,
            expiresAt: auth.expiresAt || null,
            models: models.models,
            modelsFromApi: models.ok === true
        };
    } catch (e) {
        return { ok: false, mode: null, error: e.message, models: FALLBACK_MODELS.slice(), modelsFromApi: false };
    }
}

module.exports = {
    ANTHROPIC_HOST,
    ANTHROPIC_ENDPOINT,
    CLAUDE_CODE_IDENTITY,
    FALLBACK_MODELS,
    isAnthropicEndpoint,
    credentialsPath,
    readOauthCredentials,
    isExpired,
    refreshAccessToken,
    writeOauthCredentials,
    resolveAuth,
    authHeaders,
    toAnthropicContent,
    toAnthropicRequest,
    toOpenAiResponse,
    mapStopReason,
    SAMPLING_FIELDS,
    rejectsSampling,
    rememberRejectsSampling,
    isSamplingUnsupportedError,
    stripSamplingParams,
    createSseTranslator,
    listModels,
    authStatus
};
