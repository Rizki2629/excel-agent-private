const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { Readable } = require('node:stream');
const { createRequestHandler } = require('../server');
const a = require('../anthropic_provider');

function request(method, url, body) {
  const req = Readable.from(body === undefined ? [] : [Buffer.from(JSON.stringify(body))]);
  req.method = method;
  req.url = url;
  req.headers = {};
  return req;
}

function response() {
  return {
    statusCode: 0,
    headers: null,
    body: '',
    headersSent: false,
    writableEnded: false,
    writeHead(code, headers) { this.statusCode = code; this.headers = headers; this.headersSent = true; },
    write(chunk) { this.body += chunk; },
    end(chunk) { if (chunk) this.body += chunk; this.writableEnded = true; },
    on() { return this; }
  };
}

test('anthropic: deteksi endpoint hanya untuk host api.anthropic.com', () => {
  assert.equal(a.isAnthropicEndpoint('https://api.anthropic.com/v1/messages'), true);
  assert.equal(a.isAnthropicEndpoint('https://api.anthropic.com/v1/chat/completions'), true);
  assert.equal(a.isAnthropicEndpoint('https://raiis4h.abc-tunnel.us/v1/chat/completions'), false);
  assert.equal(a.isAnthropicEndpoint('https://evil.test/api.anthropic.com'), false);
  assert.equal(a.isAnthropicEndpoint(''), false);
  assert.equal(a.isAnthropicEndpoint(null), false);
});

test('anthropic: payload OpenAI diterjemahkan ke Messages API', () => {
  const req = a.toAnthropicRequest({
    model: 'claude-sonnet-4-5-20250929',
    max_tokens: 100,
    stream: true,
    messages: [
      { role: 'system', content: 'Kamu asisten Excel.' },
      { role: 'user', content: 'Halo' },
      { role: 'assistant', content: 'Hai' },
      { role: 'user', content: 'Buat rumus' }
    ]
  }, 'api_key');
  assert.equal(req.model, 'claude-sonnet-4-5-20250929');
  // max_tokens kecil dinaikkan ke lantai minimum (lihat test MIN_MAX_TOKENS)
  assert.ok(req.max_tokens >= 100);
  assert.equal(req.stream, true);
  assert.ok(Array.isArray(req.system));
  assert.equal(req.system[0].text, 'Kamu asisten Excel.');
  assert.equal(req.messages.length, 3);
  assert.equal(req.messages[0].role, 'user');
});

test('anthropic: max_tokens default terpasang bila tidak dikirim', () => {
  const req = a.toAnthropicRequest({ model: 'm', messages: [{ role: 'user', content: 'x' }] }, 'api_key');
  assert.ok(req.max_tokens > 0);
});

test('anthropic: pesan role sama berurutan digabung dan harus mulai dari user', () => {
  const req = a.toAnthropicRequest({
    model: 'm',
    messages: [
      { role: 'assistant', content: 'duluan assistant' },
      { role: 'user', content: 'satu' },
      { role: 'user', content: 'dua' }
    ]
  }, 'api_key');
  assert.equal(req.messages[0].role, 'user');
  for (let i = 1; i < req.messages.length; i++) {
    assert.notEqual(req.messages[i].role, req.messages[i - 1].role);
  }
});

test('anthropic: mode oauth menyuntikkan identitas Claude Code di system prompt', () => {
  const req = a.toAnthropicRequest({ model: 'm', messages: [{ role: 'user', content: 'hai' }] }, 'oauth');
  assert.ok(Array.isArray(req.system));
  assert.match(req.system[0].text, /You are Claude Code/);

  const noOauth = a.toAnthropicRequest({ model: 'm', messages: [{ role: 'user', content: 'hai' }] }, 'api_key');
  assert.equal(noOauth.system, undefined);
});

test('anthropic: gambar data URL dikonversi jadi blok image base64', () => {
  const blocks = a.toAnthropicContent([
    { type: 'text', text: 'lihat ini' },
    { type: 'image_url', image_url: { url: 'data:image/png;base64,AAAB' } }
  ]);
  assert.equal(blocks.length, 2);
  assert.equal(blocks[1].type, 'image');
  assert.equal(blocks[1].source.media_type, 'image/png');
  assert.equal(blocks[1].source.data, 'AAAB');
});

test('anthropic: respons Messages API dikonversi ke bentuk OpenAI', () => {
  const out = a.toOpenAiResponse({
    id: 'msg_1',
    model: 'claude-sonnet-4-5-20250929',
    content: [{ type: 'text', text: 'Jawaban' }],
    stop_reason: 'end_turn',
    usage: { input_tokens: 10, output_tokens: 5 }
  });
  assert.equal(out.object, 'chat.completion');
  assert.equal(out.choices[0].message.content, 'Jawaban');
  assert.equal(out.choices[0].finish_reason, 'stop');
  assert.equal(out.usage.total_tokens, 15);
});

test('anthropic: stop_reason dipetakan ke finish_reason OpenAI', () => {
  assert.equal(a.mapStopReason('end_turn'), 'stop');
  assert.equal(a.mapStopReason('stop_sequence'), 'stop');
  assert.equal(a.mapStopReason('max_tokens'), 'length');
});

test('anthropic: SSE Anthropic diterjemahkan ke SSE OpenAI dengan DONE', () => {
  const written = [];
  const t = a.createSseTranslator('claude-test', (s) => written.push(s));
  t.push('event: content_block_delta\ndata: {"type":"content_block_delta","delta":{"type":"text_delta","text":"Ha"}}\n\n');
  t.push('event: content_block_delta\ndata: {"type":"content_block_delta","delta":{"type":"text_delta","text":"lo"}}\n\n');
  t.push('event: message_delta\ndata: {"type":"message_delta","delta":{"stop_reason":"end_turn"}}\n\n');
  t.push('event: message_stop\ndata: {"type":"message_stop"}\n\n');

  const all = written.join('');
  assert.match(all, /data: \[DONE\]/);

  // Rakit ulang persis seperti parser task pane.
  let text = '';
  let finish = null;
  all.split('\n\n').forEach(block => {
    const line = block.trim();
    if (line.indexOf('data:') !== 0) return;
    const d = line.substring(5).trim();
    if (!d || d === '[DONE]') return;
    const obj = JSON.parse(d);
    const delta = obj.choices && obj.choices[0] && obj.choices[0].delta;
    if (delta && delta.content) text += delta.content;
    if (obj.choices && obj.choices[0].finish_reason) finish = obj.choices[0].finish_reason;
  });
  assert.equal(text, 'Halo');
  assert.equal(finish, 'stop');
});

test('anthropic: SSE terpotong di tengah chunk tetap dirakit utuh', () => {
  const written = [];
  const t = a.createSseTranslator('m', (s) => written.push(s));
  t.push('event: content_block_delta\ndata: {"type":"content_block_delta","delta":{"typ');
  t.push('e":"text_delta","text":"OK"}}\n\n');
  t.push('event: message_stop\ndata: {"type":"message_stop"}\n\n');
  const all = written.join('');
  assert.match(all, /OK/);
  assert.match(all, /\[DONE\]/);
});

test('anthropic: error stream diteruskan sebagai objek error lalu DONE', () => {
  const written = [];
  const t = a.createSseTranslator('m', (s) => written.push(s));
  t.push('event: error\ndata: {"type":"error","error":{"type":"overloaded_error","message":"kepenuhan"}}\n\n');
  const all = written.join('');
  assert.match(all, /"error"/);
  assert.match(all, /kepenuhan/);
  assert.match(all, /\[DONE\]/);
});

test('anthropic: token kedaluwarsa terdeteksi lewat isExpired', () => {
  assert.equal(a.isExpired({ expiresAt: Date.now() - 1000 }), true);
  assert.equal(a.isExpired({ expiresAt: Date.now() + 3600000 }), false);
  assert.equal(a.isExpired(null), false);
});

test('anthropic: kredensial OAuth dibaca dari file bentuk claudeAiOauth', () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'anthro-'));
  const file = path.join(dir, '.credentials.json');
  fs.writeFileSync(file, JSON.stringify({
    claudeAiOauth: { accessToken: 'sk-ant-oat01-xxx', refreshToken: 'sk-ant-ort01-yyy', expiresAt: Date.now() + 60000 }
  }), 'utf8');
  const creds = a.readOauthCredentials(file);
  assert.equal(creds.accessToken, 'sk-ant-oat01-xxx');
  assert.equal(creds.refreshToken, 'sk-ant-ort01-yyy');
  assert.equal(a.readOauthCredentials(path.join(dir, 'tidak-ada.json')), null);
  fs.rmSync(dir, { recursive: true, force: true });
});

test('anthropic: header auth berbeda antara oauth dan api key', () => {
  const oauth = a.authHeaders({ mode: 'oauth', accessToken: 'tok' });
  assert.equal(oauth['Authorization'], 'Bearer tok');
  assert.ok(oauth['anthropic-beta']);
  assert.equal(oauth['x-api-key'], undefined);

  const key = a.authHeaders({ mode: 'api_key', apiKey: 'sk-ant-api-1' });
  assert.equal(key['x-api-key'], 'sk-ant-api-1');
  assert.equal(key['Authorization'], undefined);
});

test('anthropic: endpoint /api/anthropic/status tidak membocorkan token', async () => {
  const handler = createRequestHandler();
  const res = response();
  await handler(request('GET', '/api/anthropic/status'), res);
  assert.equal(res.statusCode, 200);
  const data = JSON.parse(res.body);
  assert.ok(Array.isArray(data.models));
  assert.equal(/sk-ant-(oat|ort|api)/.test(res.body), false);
});

test('taskpane menyediakan kontrol login Anthropic', () => {
  const source = fs.readFileSync(path.join(__dirname, '..', 'taskpane', 'taskpane.js'), 'utf8');
  const html = fs.readFileSync(path.join(__dirname, '..', 'taskpane', 'taskpane.html'), 'utf8');
  assert.match(html, /id="anthropicModelSelect"/);
  assert.match(html, /id="useAnthropicBtn"/);
  assert.match(source, /\/api\/anthropic\/status/);
  assert.match(source, /\/api\/anthropic\/activate/);
});

test('server mengarahkan endpoint Anthropic ke adapter Messages API', () => {
  const source = fs.readFileSync(path.join(__dirname, '..', 'server.js'), 'utf8');
  assert.match(source, /anthropic\.isAnthropicEndpoint\(apiEndpoint\)/);
  assert.match(source, /proxyToAnthropic\(parsedRequest, res, readConfigFn\)/);
});

test('anthropic: deteksi error parameter sampling yang tidak didukung', () => {
  assert.equal(a.isSamplingUnsupportedError('`temperature` is deprecated for this model.'), true);
  assert.equal(a.isSamplingUnsupportedError('temperature is not supported'), true);
  assert.equal(a.isSamplingUnsupportedError('`top_p` cannot be used with this model'), true);
  // Error lain tidak boleh ikut terpicu
  assert.equal(a.isSamplingUnsupportedError('rate_limit_error'), false);
  assert.equal(a.isSamplingUnsupportedError('overloaded'), false);
  assert.equal(a.isSamplingUnsupportedError('model not found'), false);
  assert.equal(a.isSamplingUnsupportedError(''), false);
  assert.equal(a.isSamplingUnsupportedError(null), false);
});

test('anthropic: stripSamplingParams membuang temperature/top_p tanpa merusak field lain', () => {
  const out = a.stripSamplingParams({
    model: 'm', max_tokens: 50, temperature: 0.2, top_p: 0.9,
    messages: [{ role: 'user', content: 'x' }], system: [{ type: 'text', text: 's' }], stream: true
  });
  assert.equal(out.removed, true);
  assert.equal(out.request.temperature, undefined);
  assert.equal(out.request.top_p, undefined);
  assert.equal(out.request.model, 'm');
  assert.equal(out.request.max_tokens, 50);
  assert.equal(out.request.stream, true);
  assert.deepEqual(out.request.messages, [{ role: 'user', content: 'x' }]);
  assert.deepEqual(out.request.system, [{ type: 'text', text: 's' }]);

  const none = a.stripSamplingParams({ model: 'm', messages: [] });
  assert.equal(none.removed, false);
});

test('anthropic: model yang tercatat menolak sampling tidak lagi dikirimi temperature', () => {
  const model = 'model-uji-tolak-sampling-' + Date.now();
  const before = a.toAnthropicRequest({ model, temperature: 0.2, top_p: 0.9, messages: [{ role: 'user', content: 'x' }] }, 'api_key');
  assert.equal(before.temperature, 0.2);
  assert.equal(before.top_p, 0.9);

  a.rememberRejectsSampling(model);
  assert.equal(a.rejectsSampling(model), true);

  const after = a.toAnthropicRequest({ model, temperature: 0.2, top_p: 0.9, messages: [{ role: 'user', content: 'x' }] }, 'api_key');
  assert.equal(after.temperature, undefined);
  assert.equal(after.top_p, undefined);
  // model lain tidak terpengaruh
  const lain = a.toAnthropicRequest({ model: 'model-lain-' + Date.now(), temperature: 0.3, messages: [{ role: 'user', content: 'x' }] }, 'api_key');
  assert.equal(lain.temperature, 0.3);
});

test('anthropic: blok thinking diabaikan, hanya teks yang diambil', () => {
  const out = a.toOpenAiResponse({
    id: 'm', model: 'claude-opus-5',
    content: [
      { type: 'thinking', thinking: 'rahasia penalaran', signature: 'sig' },
      { type: 'text', text: 'Jawaban akhir' }
    ],
    stop_reason: 'end_turn', usage: {}
  });
  assert.equal(out.choices[0].message.content, 'Jawaban akhir');
  assert.equal(out.choices[0].message.content.includes('rahasia penalaran'), false);
});

test('anthropic: streaming thinking_delta tidak ikut jadi teks jawaban', () => {
  const written = [];
  const t = a.createSseTranslator('claude-opus-5', (s) => written.push(s));
  t.push('event: content_block_delta\ndata: {"type":"content_block_delta","delta":{"type":"thinking_delta","thinking":"nalar internal"}}\n\n');
  t.push('event: content_block_delta\ndata: {"type":"content_block_delta","delta":{"type":"text_delta","text":"Hasil"}}\n\n');
  t.push('event: message_stop\ndata: {"type":"message_stop"}\n\n');
  const all = written.join('');
  assert.equal(all.includes('nalar internal'), false);
  assert.match(all, /Hasil/);
});

test('server mengulang request Anthropic tanpa sampling saat ditolak model', () => {
  const source = fs.readFileSync(path.join(__dirname, '..', 'server.js'), 'utf8');
  assert.match(source, /isSamplingUnsupportedError\(message\)/);
  assert.match(source, /stripSamplingParams\(anthropicRequest\)/);
  assert.match(source, /rememberRejectsSampling\(anthropicRequest\.model\)/);
  // retry hanya sekali (flag dimatikan pada percobaan kedua)
  assert.match(source, /send\(stripped\.request, false\)/);
});

test('anthropic: max_tokens kecil dinaikkan ke lantai minimum untuk model thinking', () => {
  const req = a.toAnthropicRequest({ model: 'm', max_tokens: 60, messages: [{ role: 'user', content: 'x' }] }, 'api_key');
  assert.ok(req.max_tokens >= 2048, 'max_tokens harus dinaikkan, dapat ' + req.max_tokens);
  // nilai besar tidak diturunkan
  const besar = a.toAnthropicRequest({ model: 'm', max_tokens: 8192, messages: [{ role: 'user', content: 'x' }] }, 'api_key');
  assert.equal(besar.max_tokens, 8192);
});

test('anthropic: jawaban kosong karena thinking habiskan token diberi pesan jelas', () => {
  const out = a.toOpenAiResponse({
    id: 'm', model: 'claude-opus-5',
    content: [{ type: 'thinking', thinking: '' }],
    stop_reason: 'max_tokens',
    usage: { input_tokens: 37, output_tokens: 60 }
  });
  const c = out.choices[0].message.content;
  assert.notEqual(c.trim(), '', 'tidak boleh kosong');
  assert.match(c, /token/i);
  assert.equal(out.choices[0].finish_reason, 'length');
});

test('anthropic: jawaban normal tidak disisipi pesan peringatan', () => {
  const out = a.toOpenAiResponse({
    id: 'm', model: 'claude-opus-5',
    content: [{ type: 'thinking', thinking: 'x' }, { type: 'text', text: 'Jawaban nyata' }],
    stop_reason: 'end_turn', usage: {}
  });
  assert.equal(out.choices[0].message.content, 'Jawaban nyata');
});
