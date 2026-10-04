'use strict';

// Integrasi Hermes: baca provider & daftar model dari konfigurasi Hermes
// (config.yaml + .env) supaya Excel Agent bisa memakai model yang sama
// tanpa mengetik ulang endpoint/API key.

const fs = require('fs');
const path = require('path');
const os = require('os');
const yaml = require('js-yaml');

// Lokasi kandidat direktori Hermes, berurut prioritas.
function hermesHomeCandidates() {
    const list = [];
    if (process.env.HERMES_HOME) list.push(process.env.HERMES_HOME);
    if (process.env.LOCALAPPDATA) list.push(path.join(process.env.LOCALAPPDATA, 'hermes'));
    if (process.env.APPDATA) list.push(path.join(process.env.APPDATA, 'hermes'));
    list.push(path.join(os.homedir(), '.hermes'));
    list.push(path.join(os.homedir(), 'AppData', 'Local', 'hermes'));
    return list;
}

function findHermesHome() {
    for (const dir of hermesHomeCandidates()) {
        try {
            if (fs.existsSync(path.join(dir, 'config.yaml'))) return dir;
        } catch (e) { /* abaikan */ }
    }
    return null;
}

// Parser .env sederhana: KEY=VALUE per baris, abaikan komentar dan baris kosong.
function parseEnvFile(filePath) {
    const out = {};
    let raw;
    try {
        raw = fs.readFileSync(filePath, 'utf8');
    } catch (e) {
        return out;
    }
    for (const line of raw.split(/\r?\n/)) {
        const trimmed = line.trim();
        if (!trimmed || trimmed.charAt(0) === '#') continue;
        const eq = trimmed.indexOf('=');
        if (eq <= 0) continue;
        const key = trimmed.slice(0, eq).trim();
        let value = trimmed.slice(eq + 1).trim();
        if (value.length >= 2 && ((value[0] === '"' && value.endsWith('"')) || (value[0] === "'" && value.endsWith("'")))) {
            value = value.slice(1, -1);
        }
        if (key) out[key] = value;
    }
    return out;
}

// Resolusi nilai key: dukung literal, "${ENV_VAR}", dan nama env langsung.
function resolveSecret(value, env) {
    if (typeof value !== 'string' || !value.trim()) return '';
    const trimmed = value.trim();
    const placeholder = trimmed.match(/^\$\{([A-Za-z0-9_]+)\}$/);
    if (placeholder) return env[placeholder[1]] || process.env[placeholder[1]] || '';
    if (/^\$[A-Za-z0-9_]+$/.test(trimmed)) {
        const name = trimmed.slice(1);
        return env[name] || process.env[name] || '';
    }
    return trimmed;
}

function resolveKeyEnv(keyEnv, env) {
    if (typeof keyEnv !== 'string' || !keyEnv.trim()) return '';
    const name = keyEnv.trim();
    return env[name] || process.env[name] || '';
}

// Endpoint Excel Agent butuh URL chat completions penuh.
function toChatCompletionsEndpoint(baseUrl) {
    if (typeof baseUrl !== 'string' || !baseUrl.trim()) return '';
    let url = baseUrl.trim().replace(/\/+$/, '');
    if (/\/chat\/completions$/i.test(url)) return url;
    return url + '/chat/completions';
}

function modelNamesOf(entry) {
    const names = [];
    if (entry && typeof entry.model === 'string' && entry.model.trim()) names.push(entry.model.trim());
    const models = entry && entry.models;
    if (models && typeof models === 'object' && !Array.isArray(models)) {
        for (const name of Object.keys(models)) {
            if (name && names.indexOf(name) === -1) names.push(name);
        }
    } else if (Array.isArray(models)) {
        for (const m of models) {
            const name = typeof m === 'string' ? m : (m && m.name);
            if (name && names.indexOf(name) === -1) names.push(name);
        }
    }
    return names;
}

// Baca semua provider Hermes: model utama, custom_providers, fallback_providers.
// Mengembalikan array { name, baseUrl, apiEndpoint, apiKey, model, models[], source }.
function readHermesProviders(hermesHome = findHermesHome()) {
    if (!hermesHome) return { ok: false, error: 'Konfigurasi Hermes tidak ditemukan di komputer ini.', hermesHome: null, providers: [] };

    let config;
    try {
        config = yaml.load(fs.readFileSync(path.join(hermesHome, 'config.yaml'), 'utf8')) || {};
    } catch (e) {
        return { ok: false, error: 'Gagal membaca config.yaml Hermes: ' + e.message, hermesHome, providers: [] };
    }

    const env = parseEnvFile(path.join(hermesHome, '.env'));
    const providers = [];
    const seen = new Set();

    const push = (name, baseUrl, apiKey, model, models, source) => {
        const cleanName = (name || '').toString().trim();
        const endpoint = toChatCompletionsEndpoint(baseUrl);
        if (!cleanName || !endpoint) return;
        const key = cleanName.toLowerCase();
        if (seen.has(key)) {
            // Gabungkan daftar model bila provider sama muncul di beberapa bagian.
            const existing = providers.find(p => p.name.toLowerCase() === key);
            if (existing) {
                for (const m of models || []) if (existing.models.indexOf(m) === -1) existing.models.push(m);
                if (!existing.apiKey && apiKey) existing.apiKey = apiKey;
            }
            return;
        }
        seen.add(key);
        providers.push({
            name: cleanName,
            baseUrl: (baseUrl || '').trim(),
            apiEndpoint: endpoint,
            apiKey: apiKey || '',
            model: model || (models && models[0]) || '',
            models: (models || []).slice(),
            source
        });
    };

    // 1. custom_providers (sumber utama daftar model Hermes)
    const custom = Array.isArray(config.custom_providers) ? config.custom_providers : [];
    for (const entry of custom) {
        if (!entry || typeof entry !== 'object') continue;
        const apiKey = resolveKeyEnv(entry.key_env, env) || resolveSecret(entry.api_key, env);
        push(entry.name, entry.base_url, apiKey, entry.model, modelNamesOf(entry), 'custom_providers');
    }

    // 2. model aktif Hermes (default) — bila belum tercakup custom_providers
    const active = config.model && typeof config.model === 'object' ? config.model : null;
    if (active && active.base_url) {
        const apiKey = resolveSecret(active.api_key, env) || resolveKeyEnv(active.key_env, env);
        const activeName = (active.provider && active.provider !== 'custom')
            ? String(active.provider)
            : 'Hermes (aktif)';
        const models = [];
        if (typeof active.default === 'string' && active.default.trim()) models.push(active.default.trim());
        // Jika base_url-nya sama dengan salah satu custom provider, cukup pastikan modelnya ada.
        const twin = providers.find(p => p.baseUrl.replace(/\/+$/, '') === String(active.base_url).replace(/\/+$/, ''));
        if (twin) {
            for (const m of models) if (twin.models.indexOf(m) === -1) twin.models.unshift(m);
            if (models[0]) twin.model = models[0];
            if (!twin.apiKey && apiKey) twin.apiKey = apiKey;
        } else {
            push(activeName, active.base_url, apiKey, models[0] || '', models, 'model');
        }
    }

    // 3. fallback_providers
    const fallbacks = Array.isArray(config.fallback_providers) ? config.fallback_providers : [];
    for (const entry of fallbacks) {
        if (!entry || typeof entry !== 'object') continue;
        const apiKey = resolveSecret(entry.api_key, env) || resolveKeyEnv(entry.key_env, env);
        push(entry.name || entry.provider, entry.base_url, apiKey, entry.model, modelNamesOf(entry), 'fallback_providers');
    }

    return { ok: true, hermesHome, providers };
}

// Bentuk aman untuk dikirim ke task pane: tanpa API key mentah.
function publicHermesProviders(result) {
    return {
        ok: result.ok === true,
        error: result.error || null,
        hermesHome: result.hermesHome || null,
        providers: (result.providers || []).map(p => ({
            name: p.name,
            apiEndpoint: p.apiEndpoint,
            model: p.model,
            models: p.models,
            modelCount: p.models.length,
            hasKey: typeof p.apiKey === 'string' && p.apiKey.length > 0,
            source: p.source
        }))
    };
}

function findHermesProvider(name, hermesHome) {
    const result = readHermesProviders(hermesHome);
    if (!result.ok) return { ok: false, error: result.error };
    const wanted = (name || '').toString().trim().toLowerCase();
    const found = result.providers.find(p => p.name.toLowerCase() === wanted);
    if (!found) return { ok: false, error: 'Provider Hermes "' + name + '" tidak ditemukan.' };
    return { ok: true, provider: found };
}

module.exports = {
    findHermesHome,
    parseEnvFile,
    resolveSecret,
    resolveKeyEnv,
    toChatCompletionsEndpoint,
    modelNamesOf,
    readHermesProviders,
    publicHermesProviders,
    findHermesProvider
};
