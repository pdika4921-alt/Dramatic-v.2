// ═══════════════════════════════════════════════════════════════════════════
// driveApi.js — kirim eviden ke Google Drive lewat Drive API v3 (murni Node)
//
// Dibuat untuk hosting (Hostinger dsb) yang TIDAK bisa mengeksekusi binary
// rclone. Jalur lokal (komputer) tetap memakai rclone seperti biasa — modul
// ini hanya dipakai bila rclone tidak tersedia / DRIVE_SYNC_MODE=api.
//
// Kredensial OAuth (client_id + client_secret + refresh_token) diambil dari:
//   1. environment: DRIVE_API_CLIENT_ID / DRIVE_API_CLIENT_SECRET /
//      DRIVE_API_REFRESH_TOKEN / DRIVE_API_ROOT_FOLDER_ID
//   2. file: drive-api.json di root proyek (lihat drive-api.example.json)
//
// Sifat:
//   * Tidak pernah menghapus file di Drive (hanya menambah/menimpa).
//   * Antrian serial + jeda ≥340 ms antar operasi tulis (batas Drive ±3/detik).
//   * Token di-refresh otomatis saat kedaluwarsa / kena 401.
// ═══════════════════════════════════════════════════════════════════════════
const fs = require('fs');
const path = require('path');
const https = require('https');

const CONFIG_PATH = process.env.DRIVE_API_CONFIG || path.join(__dirname, '..', 'drive-api.json');
const TOKEN_URL   = 'https://oauth2.googleapis.com/token';
const API_BASE    = 'https://www.googleapis.com/drive/v3';
const UPLOAD_BASE = 'https://www.googleapis.com/upload/drive/v3/files';
const FOLDER_MIME = 'application/vnd.google-apps.folder';

const JEDA_TULIS  = 340;    // ms minimal antar operasi tulis Drive (±3/detik)
const MAX_COBA    = 5;      // percobaan per operasi

let cfg = null;             // { client_id, client_secret, refresh_token, root_folder_id }
let accessToken = null;
let tokenExpAt = 0;
let gateTerakhir = 0;       // waktu operasi tulis terakhir

const folderCache = {};     // "<parentId>|<nama>" -> folderId
const stat = { terkirim: 0, gagal: 0, terakhir: null };

// ── util ─────────────────────────────────────────────────────────────────────
function jeda(ms) { return new Promise(r => setTimeout(r, ms)); }

// Gerbang rate-limit: pastikan ≥ JEDA_TULIS ms sejak operasi tulis terakhir.
async function gerbangTulis() {
  const tunggu = gateTerakhir + JEDA_TULIS - Date.now();
  if (tunggu > 0) await jeda(tunggu);
  gateTerakhir = Date.now();
}

function esc(s) {
  return String(s).replace(/\\/g, '\\\\').replace(/'/g, "\\'");
}

function req(urlStr, opts = {}) {
  return new Promise((resolve, reject) => {
    const u = new URL(urlStr);
    const r = https.request({
      hostname: u.hostname,
      path: u.pathname + u.search,
      method: opts.method || 'GET',
      headers: opts.headers || {},
      timeout: opts.timeout || 60000,
    }, res => {
      const potongan = [];
      res.on('data', d => potongan.push(d));
      res.on('end', () => {
        const buf = Buffer.concat(potongan);
        let json = null;
        try { json = JSON.parse(buf.toString('utf8')); } catch (e) { /* bukan JSON */ }
        resolve({ code: res.statusCode, json, text: buf.toString('utf8') });
      });
    });
    r.on('error', reject);
    r.on('timeout', () => { r.destroy(new Error('timeout')); });
    if (opts.body) r.write(opts.body);
    r.end();
  });
}

// ── kredensial ───────────────────────────────────────────────────────────────
function muatConfig() {
  if (cfg) return { ok: true };
  const env = {
    client_id: process.env.DRIVE_API_CLIENT_ID,
    client_secret: process.env.DRIVE_API_CLIENT_SECRET,
    refresh_token: process.env.DRIVE_API_REFRESH_TOKEN,
    root_folder_id: process.env.DRIVE_API_ROOT_FOLDER_ID,
  };
  if (env.client_id && env.client_secret && env.refresh_token && env.root_folder_id) {
    cfg = env;
    return { ok: true };
  }
  try {
    const j = JSON.parse(fs.readFileSync(CONFIG_PATH, 'utf8'));
    if (j.client_id && j.client_secret && j.refresh_token && j.root_folder_id) {
      cfg = j;
      return { ok: true };
    }
    return { ok: false, alasan: 'drive-api.json belum lengkap (butuh client_id, client_secret, refresh_token, root_folder_id)' };
  } catch (e) {
    return { ok: false, alasan: 'config Drive API tidak ditemukan: ' + CONFIG_PATH };
  }
}

async function ambilToken() {
  if (accessToken && Date.now() < tokenExpAt - 60000) return accessToken;
  const body = new URLSearchParams({
    grant_type: 'refresh_token',
    refresh_token: cfg.refresh_token,
    client_id: cfg.client_id,
    client_secret: cfg.client_secret,
  }).toString();
  const r = await req(TOKEN_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded', 'Content-Length': Buffer.byteLength(body) },
    body,
  });
  if (r.code !== 200 || !r.json || !r.json.access_token) {
    throw new Error('refresh token gagal (HTTP ' + r.code + '): ' + String(r.text || '').slice(0, 300));
  }
  accessToken = r.json.access_token;
  tokenExpAt = Date.now() + (Number(r.json.expires_in || 3500) * 1000);
  return accessToken;
}

// ── panggilan API dengan retry (401 → refresh, 429/403 rate → tunggu) ───────
// Catatan: 403 rateLimitExceeded pada client ID bersama rclone bisa berlangsung
// puluhan detik sampai menit — karena itu backoff untuk rate jauh lebih sabar
// (sampai ±2 menit total) daripada error lain.
async function api(method, urlStr, body, headers, opsi = {}) {
  let terakhirErr = null;
  const MAX_RATE = 8;                    // percobaan ulang khusus rate limit
  for (let i = 0; i < MAX_RATE; i++) {
    let tok;
    try { tok = await ambilToken(); }
    catch (e) { terakhirErr = e; await jeda(2000 * (i + 1)); continue; }
    let r;
    try {
      r = await req(urlStr, {
        method,
        body,
        headers: Object.assign({ Authorization: 'Bearer ' + tok }, headers || {}),
      });
    } catch (e) { terakhirErr = e; await jeda(1500 * (i + 1)); continue; }

    if (r.code >= 200 && r.code < 300) return r;
    if (r.code === 401) { accessToken = null; terakhirErr = new Error('401'); continue; }
    if (r.code === 429 || (r.code === 403 && /rate|quota|userRate|PERMISSION_DENIED/i.test(r.text || ''))) {
      terakhirErr = new Error('HTTP ' + r.code + ' rate limit: ' +
        String((r.json && r.json.error && r.json.error.message) || r.text || '').slice(0, 160));
      // tunggu makin lama: 10s, 20s, 30s … 80s (total ±6 menit)
      await jeda(Math.min(80, 10 * (i + 1)) * 1000);
      continue;
    }
    if (r.code >= 500) { terakhirErr = new Error('HTTP ' + r.code); await jeda(1500 * (i + 1)); continue; }
    // error permanen (400/403 non-rate/404) → tidak usah diulang
    const e = new Error('HTTP ' + r.code + ': ' + String(r.text || '').slice(0, 300));
    e.permanen = true;
    throw e;
  }
  throw terakhirErr || new Error('gagal tanpa keterangan');
}

// ── folder ───────────────────────────────────────────────────────────────────
async function folderId(parentId, nama) {
  const kunci = parentId + '|' + nama;
  if (folderCache[kunci]) return folderCache[kunci];

  const q = `'${parentId}' in parents and name='${esc(nama)}' and mimeType='${FOLDER_MIME}' and trashed=false`;
  const list = await api('GET', `${API_BASE}/files?q=${encodeURIComponent(q)}&fields=files(id)&pageSize=1`);
  if (list.json && list.json.files && list.json.files.length) {
    folderCache[kunci] = list.json.files[0].id;
    return folderCache[kunci];
  }

  await gerbangTulis();
  const meta = JSON.stringify({ name: nama, parents: [parentId], mimeType: FOLDER_MIME });
  const dibuat = await api('POST', `${API_BASE}/files?fields=id`, meta,
    { 'Content-Type': 'application/json' });
  if (!dibuat.json || !dibuat.json.id) throw new Error('gagal membuat folder ' + nama);
  folderCache[kunci] = dibuat.json.id;
  return dibuat.json.id;
}

// Pastikan rantai folder "a/b/c" ada di bawah root → id folder terbawah.
async function folderIdDariRel(relDir) {
  if (!relDir || relDir === '.') return cfg.root_folder_id;
  let parent = cfg.root_folder_id;
  for (const seg of relDir.split('/')) {
    if (!seg) continue;
    parent = await folderId(parent, seg);
  }
  return parent;
}

// ── upload satu file (multipart: metadata + isi file) ────────────────────────
async function upload(abs, rel) {
  await muatConfig();
  const isi = fs.readFileSync(abs);
  if (!isi.length) return false;

  const dirRel = path.posix.dirname(rel.replace(/\\/g, '/'));
  const nama   = path.posix.basename(rel.replace(/\\/g, '/'));
  const parentId = await folderIdDariRel(dirRel === '.' ? '' : dirRel);

  // sudah ada dengan ukuran sama? → lewati (idempoten)
  const q = `'${parentId}' in parents and name='${esc(nama)}' and mimeType!='${FOLDER_MIME}' and trashed=false`;
  const ada = await api('GET', `${API_BASE}/files?q=${encodeURIComponent(q)}&fields=files(id,size)&pageSize=1`);
  if (ada.json && ada.json.files && ada.json.files.length &&
      Number(ada.json.files[0].size) === isi.length) {
    return true;
  }

  const batas = '----driveapi' + Date.now().toString(16) + Math.random().toString(16).slice(2, 10);
  const meta  = JSON.stringify({ name: nama, parents: [parentId] });
  const body = Buffer.concat([
    Buffer.from(`--${batas}\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n${meta}\r\n--${batas}\r\nContent-Type: application/octet-stream\r\n\r\n`),
    isi,
    Buffer.from(`\r\n--${batas}--`),
  ]);

  await gerbangTulis();
  const r = await api('POST',
    `${UPLOAD_BASE}?uploadType=multipart&fields=id,name,size`,
    body,
    { 'Content-Type': `multipart/related; boundary=${batas}`, 'Content-Length': body.length });

  const ok = !!(r.json && r.json.id);
  stat.terakhir = { rel, ok, ukuran: isi.length, detik: null };
  if (ok) { stat.terkirim++; }
  else { stat.gagal++; }
  return ok;
}

// ── daftar isi Drive dari root_folder_id: { "rel/path": size } ───────────────
// Jalan berjenjang per folder (bukan listing seluruh Drive) — hemat kuota API.
async function daftar() {
  await muatConfig();
  const hasil = {};
  let antre = [{ id: cfg.root_folder_id, prefix: '' }];
  let halaman = 0;
  while (antre.length) {
    const kini = antre.shift();
    let pageToken = '';
    do {
      const q = `'${kini.id}' in parents and trashed=false`;
      const r = await api('GET',
        `${API_BASE}/files?pageSize=1000&q=${encodeURIComponent(q)}` +
        `&fields=nextPageToken,files(id,name,parents,mimeType,size)` +
        (pageToken ? `&pageToken=${pageToken}` : ''));
      if (!r.json || !r.json.files) {
        throw new Error('daftar Drive gagal (HTTP ' + r.code + '): ' + String(r.text || '').slice(0, 200));
      }
      for (const f of r.json.files) {
        if (f.mimeType === FOLDER_MIME) antre.push({ id: f.id, prefix: kini.prefix + f.name + '/' });
        else hasil[kini.prefix + f.name] = Number(f.size || 0);
      }
      pageToken = r.json.nextPageToken || '';
      if (++halaman % 20 === 0) await jeda(350);   // jeda berkala anti-403
    } while (pageToken);
  }
  return hasil;
}

// ── uji akses (dipakai saat boot) ────────────────────────────────────────────
async function ujiAkses() {
  const muat = muatConfig();
  if (!muat.ok) return { ok: false, alasan: muat.alasan };
  try {
    await ambilToken();
    const q = `'${cfg.root_folder_id}' in parents and trashed=false`;
    const r = await api('GET', `${API_BASE}/files?q=${encodeURIComponent(q)}&fields=files(id)&pageSize=1`);
    if (r.code === 404 || (r.json && r.json.error)) {
      return { ok: false, alasan: 'Folder tujuan tidak bisa diakses akun Google ini: ' + String(r.text || '').slice(0, 200) };
    }
    return { ok: true };
  } catch (e) {
    return { ok: false, alasan: 'Drive API gagal diakses: ' + e.message };
  }
}

function status() {
  return { mode: 'api', terkirim: stat.terkirim, gagal: stat.gagal, terakhir: stat.terakhir, config: CONFIG_PATH };
}

module.exports = { muatConfig, upload, daftar, ujiAkses, status };
