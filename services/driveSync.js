// ═══════════════════════════════════════════════════════════════════════════
// driveSync.js — pencadangan eviden (folder uploads/) ke Google Drive
//
// Desain:
//   * Tujuan: remote rclone "gd" (lihat %APPDATA%\rclone\rclone.conf),
//     root_folder_id = folder Drive pilihan user.
//   * Metode: OAuth milik user sendiri (rclone) → file milik user, memakai
//     kuota Drive 15 GB user. Bukan service account (sejak 15 Apr 2025
//     service account baru tidak bisa memiliki file Drive).
//   * Cakupan: SELURUH isi uploads/ — dipantau lewat fs.watch (recursive)
//     sehingga semua jalur penulisan file (upload eviden, KML, BAUT-1, dll)
//     ikut tercadangkan tanpa harus menyentuh kode tiap endpoint.
//   * Startup: pindai uploads/ dan kirim file yang belum ada di manifest
//     (backfill 1.200+ file lama + file yang tertulis saat server mati).
//   * Tidak pernah menghapus: pakai `rclone copyto` (bukan sync) → file yang
//     dihapus lokal TETAP ada di Drive (backup aman dari salah hapus).
//   * Antrian serial + --tpslimit menjaga batas tulis Drive API (±3 req/detik).
//
// Status: aktif hanya bila remote rclone tersedia & token sudah di-otorisasi.
// ═══════════════════════════════════════════════════════════════════════════
const { spawn } = require('child_process');
const fs = require('fs');
const path = require('path');

const RCLONE   = process.env.RCLONE_PATH || 'D:\\Data\\dika baru\\rclone\\rclone.exe';
const REMOTE   = process.env.DRIVE_REMOTE || 'gd';
const ROOT     = path.join(__dirname, '..', 'uploads');
const STATE    = path.join(__dirname, '..', '.drive-sync', 'manifest.json');
const LOGFILE  = path.join(__dirname, '..', '.drive-sync', 'sync.log');
const DIR_STATE = path.join(__dirname, '..', '.drive-sync');

// Mode pengiriman:
//   'auto' → rclone bila binari ada, kalau tidak → Drive API (untuk hosting
//            yang tidak bisa mengeksekusi binary — lihat services/driveApi.js)
//   'rclone' / 'api' → paksa salah satu (uji/ debugging)
const MODE_PAKSA = (process.env.DRIVE_SYNC_MODE || 'auto').toLowerCase();
let MODE = 'rclone';        // dipilih saat mulai()

const TPSLIMIT   = Number(process.env.DRIVE_TPSLIMIT || 2);   // upload/detik
const MAX_Coba   = 5;                                          // percobaan per file
const JEDA_EVENT = 1500;                                       // debounce fs.watch (ms)

// Folder cache turunan (dibuat ulang otomatis oleh server) — jangan dicadangkan.
const LEWATI_DIR = new Set(['_pdfcache', '_dwgcache']);
const EXCLUDE_ARGS = ['--exclude', '/_pdfcache/**', '--exclude', '/_dwgcache/**'];

let manifest = {};        // "rel/path.jpg" -> { size, at }
let antrian   = [];       // { abs, rel, coba }
let sedangJalan = false;
let bulkJalan = false;    // true saat backfill massal (rclone copy) berjalan
let mati = false;
let watcher = null;
let timerScan = null;
let eventDebounce = {};

const stat = { terkirim: 0, gagal: 0, antre: 0, terakhir: null, aktif: false, alasan: null, bulk: null };

// ── util ─────────────────────────────────────────────────────────────────────
function catat(pesan) {
  const baris = `[${new Date().toISOString()}] ${pesan}\n`;
  try {
    fs.mkdirSync(DIR_STATE, { recursive: true });
    fs.appendFileSync(LOGFILE, baris);
  } catch (e) { /* log gagal bukan masalah fatal */ }
  console.log('[DriveSync] ' + pesan);
}

function bacaManifest() {
  try { manifest = JSON.parse(fs.readFileSync(STATE, 'utf8')) || {}; }
  catch (e) { manifest = {}; }
}
function simpanManifest() {
  try {
    fs.mkdirSync(DIR_STATE, { recursive: true });
    const tmp = STATE + '.tmp';
    fs.writeFileSync(tmp, JSON.stringify(manifest));
    fs.renameSync(tmp, STATE);
  } catch (e) { console.warn('[DriveSync] manifest gagal disimpan:', e.message); }
}

function relDari(abs) {
  const r = path.relative(ROOT, abs).replace(/\\/g, '/');
  if (r.startsWith('..')) return null;
  // Folder cache turunan tidak ikut dicadangkan
  const atas = r.split('/')[0];
  if (LEWATI_DIR.has(atas)) return null;
  return r;
}

function sudahTerkirim(rel, size) {
  const m = manifest[rel];
  return !!(m && m.size === size);
}

// ── proses antrian ───────────────────────────────────────────────────────────
function jalan() {
  if (sedangJalan || mati || bulkJalan) return;   // bulk dulu, antrian menyusul
  const item = antrian.shift();
  if (!item) { stat.antre = 0; return; }
  sedangJalan = true;
  stat.antre = antrian.length;

  let ukuran = 0;
  try { ukuran = fs.statSync(item.abs).size; } catch (e) { /* file hilang */ }
  if (!ukuran) {
    sedangJalan = false;
    return setTimeout(jalan, 500);
  }
  if (sudahTerkirim(item.rel, ukuran)) {
    sedangJalan = false;
    return setTimeout(jalan, 0);
  }

  // ── jalur Drive API (hosting tanpa rclone) ────────────────────────────────
  if (MODE === 'api') {
    const driveApi = require('./driveApi');
    const t0 = Date.now();
    driveApi.upload(item.abs, item.rel).then(ok => {
      sedangJalan = false;
      if (ok) {
        manifest[item.rel] = { size: ukuran, at: Date.now() };
        stat.terkirim++;
        stat.terakhir = { rel: item.rel, ok: true, detik: ((Date.now() - t0) / 1000).toFixed(1) };
        if (stat.terkirim % 25 === 0 || antrian.length === 0) simpanManifest();
        catat(`OK  ${item.rel} (${(ukuran / 1024).toFixed(0)} KB, ${stat.terakhir.detik}s)`);
      } else {
        item.coba = (item.coba || 0) + 1;
        stat.gagal++;
        stat.terakhir = { rel: item.rel, ok: false, pesan: 'Drive API menolak' };
        if (item.coba < MAX_Coba && !mati) {
          const tunda = Math.min(60000, 3000 * Math.pow(2, item.coba));
          catat(`Coba ${item.coba}/${MAX_Coba} gagal ${item.rel}`);
          setTimeout(() => { antrian.push(item); stat.antre = antrian.length; jalan(); }, tunda);
        } else {
          catat(`GAGAL ${item.rel} setelah ${item.coba}x`);
          simpanManifest();
        }
      }
      setTimeout(jalan, 250);
    }).catch(e => {
      sedangJalan = false;
      item.coba = (item.coba || 0) + 1;
      stat.gagal++;
      stat.terakhir = { rel: item.rel, ok: false, pesan: e.message };
      if (item.coba < MAX_Coba && !mati) {
        const tunda = Math.min(60000, 3000 * Math.pow(2, item.coba));
        catat(`Coba ${item.coba}/${MAX_Coba} error ${item.rel}: ${e.message.slice(0, 200)}`);
        setTimeout(() => { antrian.push(item); stat.antre = antrian.length; jalan(); }, tunda);
      } else {
        catat(`GAGAL ${item.rel} setelah ${item.coba}x: ${e.message.slice(0, 300)}`);
        simpanManifest();
      }
      setTimeout(jalan, 250);
    });
    return;
  }

  const tujuan = `${REMOTE}:${item.rel}`;
  const args = [
    'copyto', item.abs, tujuan,
    '--tpslimit', String(TPSLIMIT),
    '--tpslimit-burst', '1',
    '--retries', '3',
    '--low-level-retries', '10',
    '--stats', '0',
    '--log-level', 'ERROR',
  ];

  const t0 = Date.now();
  const anak = spawn(RCLONE, args, { windowsHide: true });
  let err = '';
  anak.stderr.on('data', d => { err += d.toString(); });
  anak.on('error', e => { err = e.message; });
  anak.on('close', code => {
    sedangJalan = false;
    if (code === 0) {
      manifest[item.rel] = { size: ukuran, at: Date.now() };
      stat.terkirim++;
      stat.terakhir = { rel: item.rel, ok: true, detik: ((Date.now() - t0) / 1000).toFixed(1) };
      if (stat.terkirim % 25 === 0 || antrian.length === 0) simpanManifest();
      catat(`OK  ${item.rel} (${(ukuran / 1024).toFixed(0)} KB, ${stat.terakhir.detik}s)`);
    } else {
      item.coba = (item.coba || 0) + 1;
      stat.gagal++;
      stat.terakhir = { rel: item.rel, ok: false, pesan: err.trim().slice(0, 300) };
      if (item.coba < MAX_Coba && !mati) {
        const tunda = Math.min(60000, 3000 * Math.pow(2, item.coba));  // backoff
        catat(`Coba ${item.coba}/${MAX_Coba} gagal (kode ${code}) ${item.rel}: ${err.trim().slice(0, 200)}`);
        setTimeout(() => { antrian.push(item); stat.antre = antrian.length; jalan(); }, tunda);
      } else {
        catat(`GAGAL ${item.rel} setelah ${item.coba}x: ${err.trim().slice(0, 300)}`);
        simpanManifest();   // supaya tidak diulang terus di tiap restart
      }
    }
    setTimeout(jalan, 250);
  });
}

function antre(abs, sebab) {
  if (!stat.aktif) return;
  const rel = relDari(abs);
  if (!rel) return;
  if (!fs.existsSync(abs)) return;
  let size = 0;
  try { size = fs.statSync(abs).size; } catch (e) { return; }
  if (!size || sudahTerkirim(rel, size)) return;
  if (antrian.some(i => i.rel === rel)) return;
  antrian.push({ abs, rel, coba: 0, sebab: sebab || 'event' });
  stat.antre = antrian.length;
  if (antrian.length === 1 || antrian.length % 50 === 0) {
    catat(`Antre ${antrian.length} file (${sebab || 'event'})`);
  }
  jalan();
}

// ── pindai folder (backfill + file yang tertulis saat server mati) ──────────
// Kirim yang belum ada secara BULK (rclone copy paralel) — jauh lebih cepat
// daripada memanggil rclone per file (±7 dtk/file → ±2 jam untuk 1.200 file).
function pindai() {
  if (!stat.aktif || mati || bulkJalan) return;
  const belum = [];
  const jalanKini = (dir) => {
    let isi = [];
    try { isi = fs.readdirSync(dir, { withFileTypes: true }); } catch (e) { return; }
    for (const d of isi) {
      const p = path.join(dir, d.name);
      if (d.isDirectory()) jalanKini(p);
      else if (d.isFile()) {
        const rel = relDari(p);
        if (!rel) continue;
        let size = 0;
        try { size = fs.statSync(p).size; } catch (e) { continue; }
        if (!size || sudahTerkirim(rel, size)) continue;
        belum.push({ abs: p, rel, size });
      }
    }
  };
  jalanKini(ROOT);
  if (!belum.length) return;              // semua sudah tercatat — tidak perlu ke Drive
  bulk(belum);
}

function bulk(belum) {
  bulkJalan = true;
  stat.bulk = { status: 'berjalan', total: belum.length, mulai: Date.now() };

  // ── jalur Drive API: kirim lewat antrian serial (tanpa rclone) ────────────
  if (MODE === 'api') {
    catat(`Bulk: ${belum.length} file belum tercadangkan → Drive API (serial)`);
    (async () => {
      const driveApi = require('./driveApi');
      let gagal = 0;
      for (const b of belum) {
        if (mati) break;
        try {
          const ok = await driveApi.upload(b.abs, b.rel);
          if (ok) {
            manifest[b.rel] = { size: b.size, at: Date.now() };
            stat.terkirim++;
            if (stat.terkirim % 25 === 0) simpanManifest();
          } else { gagal++; antre(b.abs, 'bulk-gagal'); }
        } catch (e) { gagal++; catat(`Bulk gagal ${b.rel}: ${e.message.slice(0, 200)}`); antre(b.abs, 'bulk-error'); }
      }
      simpanManifest();
      const cocokSemua = await sinkronManifestDariDrive().catch(() => false);
      const detik = Math.round((Date.now() - stat.bulk.mulai) / 1000);
      catat(`Bulk selesai (${gagal} gagal) — ${detik}s`);
      stat.bulk = { status: (gagal === 0 && cocokSemua) ? 'selesai' : 'gagal', total: belum.length, detik };
      bulkJalan = false;
      jalan();
    })();
    return;
  }

  catat(`Bulk: ${belum.length} file belum tercadangkan → rclone copy (paralel)`);
  const anak = spawn(RCLONE, [
    'copy', ROOT, `${REMOTE}:`,
    '--size-only',                 // nama file unik → cukup bandingkan ukuran
    '--tpslimit', '3',             // batas tulis Drive API (±3 request/detik)
    '--transfers', '6',
    '--checkers', '8',
    '--log-level', 'ERROR',
  ].concat(EXCLUDE_ARGS), { windowsHide: true });
  let err = '';
  anak.stderr.on('data', d => { err += d.toString(); });
  anak.on('error', e => { err = e.message; });
  anak.on('close', async (code) => {
    catat(`Bulk selesai (kode ${code})${err.trim() ? ' — ' + err.trim().slice(0, 200) : ''}`);
    const cocokSemua = await sinkronManifestDariDrive();
    if (code !== 0 || !cocokSemua) {
      // masih ada yang kurang → kirim ulang per file (ada retry & backoff)
      belum.forEach(b => antre(b.abs, 'bulk-gagal'));
    }
    const detik = Math.round((Date.now() - stat.bulk.mulai) / 1000);
    stat.bulk = {
      status: (code === 0 && cocokSemua) ? 'selesai' : 'gagal',
      total: belum.length, detik,
    };
    simpanManifest();
    bulkJalan = false;
    jalan();
  });
}

// Samakan manifest dengan isi Drive (sumber kebenaran) memakai lsjson.
async function sinkronManifestDariDrive() {
  let peta;
  if (MODE === 'api') {
    const driveApi = require('./driveApi');
    try { peta = await driveApi.daftar(); }
    catch (e) { catat('daftar Drive gagal: ' + String(e.message || '').trim().slice(0, 200)); return false; }
  } else {
    const r = await jalankan(['lsjson', '-R', '--fast-list', `${REMOTE}:`, '--log-level', 'ERROR'], 180000);
    if (r.code !== 0) { catat('lsjson gagal: ' + String(r.err || '').trim().slice(0, 200)); return false; }
    let arr;
    try { arr = JSON.parse(r.out); } catch (e) { catat('lsjson: respon bukan JSON'); return false; }
    peta = {};
    for (const it of arr || []) if (!it.IsDir) peta[String(it.Path).replace(/\\/g, '/')] = it.Size;
  }

  let cocok = 0, hilang = 0;
  const jalanKini = (dir) => {
    let isi = [];
    try { isi = fs.readdirSync(dir, { withFileTypes: true }); } catch (e) { return; }
    for (const d of isi) {
      const p = path.join(dir, d.name);
      if (d.isDirectory()) jalanKini(p);
      else if (d.isFile()) {
        const rel = relDari(p);
        if (!rel) continue;
        let size = 0;
        try { size = fs.statSync(p).size; } catch (e) { continue; }
        if (!size) continue;
        if (peta[rel] === size) { manifest[rel] = { size, at: Date.now() }; cocok++; }
        else { hilang++; }
      }
    }
  };
  jalanKini(ROOT);
  catat(`Manifest sinkron dengan Drive: ${cocok} cocok, ${hilang} belum ada`);
  return hilang === 0;
}

// ── deteksi perubahan file (tulis selesai) ───────────────────────────────────
function saatBerubah(berkas, tipe) {
  if (tipe === 'rename' && !fs.existsSync(berkas)) {
    // file terhapus lokal → biarkan (backup tidak menghapus di Drive)
    return;
  }
  if (!fs.existsSync(berkas)) return;
  clearTimeout(eventDebounce[berkas]);
  // Tunda agar tulisan selesai dulu, lalu pastikan ukuran stabil 2x.
  eventDebounce[berkas] = setTimeout(() => {
    delete eventDebounce[berkas];
    let s1 = 0;
    try { s1 = fs.statSync(berkas).size; } catch (e) { return; }
    if (!s1) return;
    setTimeout(() => {
      let s2 = 0;
      try { s2 = fs.statSync(berkas).size; } catch (e) { return; }
      if (s2 === s1) antre(berkas, 'tulis');
    }, 800);
  }, JEDA_EVENT);
}

// ── cek remote ───────────────────────────────────────────────────────────────
function jalankan(args, timeoutMs) {
  return new Promise(resolve => {
    let out = '', err = '';
    const a = spawn(RCLONE, args, { windowsHide: true });
    const t = setTimeout(() => { try { a.kill(); } catch (e) {} resolve({ code: -1, out, err: err + ' (timeout)' }); }, timeoutMs || 30000);
    a.stdout.on('data', d => out += d.toString());
    a.stderr.on('data', d => err += d.toString());
    a.on('error', e => { clearTimeout(t); resolve({ code: -1, out, err: e.message }); });
    a.on('close', code => { clearTimeout(t); resolve({ code, out, err }); });
  });
}

async function mulai() {
  bacaManifest();

  // ── pilih mode pengiriman ─────────────────────────────────────────────────
  const rcloneAda = fs.existsSync(RCLONE);
  if (MODE_PAKSA === 'api') {
    MODE = 'api';
  } else if (MODE_PAKSA === 'rclone') {
    MODE = 'rclone';
  } else {
    MODE = rcloneAda ? 'rclone' : 'api';
  }

  if (MODE === 'api') {
    const driveApi = require('./driveApi');
    const uji = await driveApi.ujiAkses();
    if (!uji.ok) {
      stat.alasan = uji.alasan;
      catat('NONAKTIF — ' + stat.alasan);
      return false;
    }
    stat.aktif = true;
    stat.alasan = null;
    catat(`AKTIF (Drive API, tanpa rclone) — folder tujuan terakses, ${Object.keys(manifest).length} file sudah tercatat di manifest`);
    // Samakan manifest dengan isi Drive SEBELUM pindai — kalau tidak, hosting
    // baru (manifest kosong) akan menganggap semua file belum terkirim dan
    // meng-upload ulang 1.700+ file yang sebenarnya sudah ada di Drive.
    try {
      if (await sinkronManifestDariDrive()) {
        // semua sudah cocok — pindai tidak akan menemukan apa pun
      }
    } catch (e) { catat('sinkron manifest saat boot gagal (dilewati): ' + e.message); }
  } else {
    if (!fs.existsSync(RCLONE)) {
      stat.alasan = 'rclone tidak ditemukan: ' + RCLONE;
      catat('NONAKTIF — ' + stat.alasan);
      return false;
    }
    const rem = await jalankan(['listremotes']);
    if (rem.code !== 0 || !rem.out.split('\n').map(s => s.trim().replace(/:$/, '')).includes(REMOTE)) {
      stat.alasan = `remote rclone "${REMOTE}:" belum dikonfigurasi`;
      catat('NONAKTIF — ' + stat.alasan);
      return false;
    }
    // token ada? sekaligus membuktikan folder tujuan bisa diakses akun ini
    const cek = await jalankan(['lsf', `${REMOTE}:`, '--max-depth', '1', '--log-level', 'ERROR'], 45000);
    if (cek.code !== 0) {
      const detail = (cek.err || cek.out || '').trim().slice(0, 300);
      const kunci = /token|oauth|authori[sz]|consent|invalid_grant/i.test(detail);
      stat.alasan = kunci
        ? 'Google Drive belum diotorisasi — jalankan: rclone config reconnect ' + REMOTE + ':'
        : 'Folder tujuan tidak bisa diakses akun Google ini (cek share/edit access): ' + detail;
      catat('NONAKTIF — ' + stat.alasan);
      return false;
    }
    stat.aktif = true;
    stat.alasan = null;
    catat(`AKTIF (rclone) — remote ${REMOTE}:, root folder ID dari konfigurasi, ${Object.keys(manifest).length} file sudah tercatat di manifest`);
  }

  // 1) backfill / kejar-tertinggal
  pindai();

  // 2) pantau perubahan file baru
  try {
    watcher = fs.watch(ROOT, { recursive: true }, saatBerubah);
    watcher.on('error', e => catat('watcher error: ' + e.message));
    catat('watcher uploads/ aktif');
  } catch (e) {
    catat('watcher tidak bisa dijalankan (' + e.message + ') — andalkan pindai berkala');
  }

  // 3) pindai ulang tiap 10 menit (jaga-jaga event watcher terlewat)
  timerScan = setInterval(pindai, 10 * 60 * 1000);
  return true;
}

function berhenti() {
  mati = true;
  if (watcher) { try { watcher.close(); } catch (e) {} watcher = null; }
  if (timerScan) { clearInterval(timerScan); timerScan = null; }
}

function status() {
  return {
    aktif: stat.aktif,
    alasan: stat.alasan,
    mode: MODE,
    remote: MODE === 'api' ? 'Google Drive API' : REMOTE + ':',
    manifest: Object.keys(manifest).length,
    terkirim: stat.terkirim,
    gagal: stat.gagal,
    antre: antrian.length,
    bulk: stat.bulk,
    terakhir: stat.terakhir,
    log: LOGFILE,
  };
}

// uji cepat dari CLI: node -e "require('./services/driveSync').mulai()"
module.exports = { mulai, berhenti, antre, pindai, status };
