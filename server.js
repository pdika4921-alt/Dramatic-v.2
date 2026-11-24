const express = require('express');
const session = require('express-session');
const db = require('./db-wrapper');
const multer = require('multer');
const sharp = require('sharp');
const fs = require('fs');
const path = require('path');
const os = require('os');
const crypto = require('crypto');
const { execFile } = require('child_process');
const zlib = require('zlib');
const UPLOADS_DIR = path.join(__dirname, 'uploads');
const bcrypt = require('bcryptjs'); // pakai bcryptjs sesuai package.json
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');
const ImageModule = require('docxtemplater-image-module-free');
const rekapOpmService = require('./services/rekapOpm');
// Sinkronisasi nilai BA Tambah Kurang ke BOQ TAMKUR.xlsx (tidak mengubah hitungan)
const { tulisBoqTamkur, bacaHargaBoqTamkur } = require('./services/boqTamkur');
const { bacaRekonDariExcel, bacaDrmDariExcel } = require('./services/boqRekonUpload');
// Data master BOQ DRM ALL — baca saja (dipakai tombol "CEK DATA MASTER BOQ").
const { cariLokasiDiMasterDrm } = require('./services/boqDrmMaster');
// Modul laporan MITRATEL berdiri sendiri; tidak mengubah alur dokumen lama.
const mitratelService = require('./services/mitratel');
const sizeOfMod = require('image-size');
const sizeOf = sizeOfMod.imageSize || sizeOfMod.default || sizeOfMod;
const { DOMParser: DocPrParser, XMLSerializer: DocPrSerializer } = require('@xmldom/xmldom');
const { PDFDocument } = require('pdf-lib');
const { bacaKml, parseKml, renderPeta } = require('./kml-renderer');
const godcSync = require('./godc-sync');
const verif = require('./services/verifyEviden');
const verifCfg = require('./services/verifyConfig');
const geoCheck = require('./services/geoCheck');
// Pencadangan eviden (folder uploads/) ke Google Drive — lihat services/driveSync.js
const driveSync = require('./services/driveSync');

const app = express();
// Helper Parser Docxtemplater untuk membaca variabel di dalam nested loop
// Catatan: nilai NULL dari SQL harus dikembalikan "" — bila dikembalikan null,
// docxtemplater merendernya jadi teks literal "undefined" di laporan.
const smartParser = (tag) => ({
  get: (scope, context) => {
    const cleanTag = tag.replace(/[​-‍﻿]/g, "").trim();
    // Cari dari scope terdalam hingga terluar (scope chain)
    if (context && context.scopeList) {
      for (let i = context.scopeList.length - 1; i >= 0; i--) {
        const s = context.scopeList[i];
        if (s && s[cleanTag] !== undefined && s[cleanTag] !== null) {
          return s[cleanTag];
        }
      }
    }
    if (scope && scope[cleanTag] !== undefined && scope[cleanTag] !== null) return scope[cleanTag];
    return "";
  }
});
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

// ── KOMPRESI JSON ────────────────────────────────────────────────────────────
// Respons Data Master berisi ±11.000 proyek (±7 MB JSON). Dikompresi dengan
// zlib bawaan Node (tanpa dependency baru) supaya hemat bandwidth; payload kecil
// tetap dikirim apa adanya.
app.use((req, res, next) => {
  res.setHeader('Vary', 'Accept-Encoding');
  if (!/gzip/.test(req.headers['accept-encoding'] || '')) return next();
  const json = res.json.bind(res);
  res.json = (body) => {
    let payload;
    try { payload = JSON.stringify(body); } catch (e) { return json(body); }
    res.setHeader('Content-Type', 'application/json; charset=utf-8');
    res.setHeader('Cache-Control', 'no-store');
    if (payload.length < 2048) { res.setHeader('Content-Length', Buffer.byteLength(payload)); return res.end(payload); }
    const gz = zlib.gzipSync(payload);
    res.setHeader('Content-Encoding', 'gzip');
    res.setHeader('Content-Length', gz.length);
    return res.end(gz);
  };
  next();
});

// ── SESSION STORE PERSISTEN (sqlite) ──────────────────────────────────────────
// Sesi login disimpan ke database sehingga TIDAK hilang saat server restart.


class SqliteSessionStore extends session.Store {
  get(sid, cb) {
    this._run('SELECT data FROM sessions WHERE id = ?', [sid])
      .then(rows => {
        if (!rows.length) return cb(null, null);
        const sess = JSON.parse(rows[0].data);
        const exp = sess.cookie && sess.cookie.expires ? new Date(sess.cookie.expires) : null;
        if (exp && exp < new Date()) return cb(null, null);
        cb(null, sess);
      })
      .catch(cb);
  }
  set(sid, sess, cb) {
    const data = JSON.stringify(sess);
    db.run(
      `INSERT INTO sessions (id, data, updated_at) VALUES (?, ?, datetime('now','localtime'))
       ON CONFLICT(id) DO UPDATE SET data = excluded.data, updated_at = excluded.updated_at`,
      [sid, data]
    ).then(() => cb && cb()).catch(cb);
  }
  destroy(sid, cb) {
    db.run('DELETE FROM sessions WHERE id = ?', [sid])
      .then(() => cb && cb()).catch(cb);
  }
  touch(sid, sess, cb) { this.set(sid, sess, cb); }
  _run(sql, params) {
    return db ? db.all(sql, params) : Promise.resolve([]);
  }
}

app.use(session({
  secret: 'rahasia-kontruksi-aman-bogor-2026',
  resave: false,
  saveUninitialized: false,
  store: new SqliteSessionStore(),
  cookie: {
    maxAge: 24 * 60 * 60 * 1000,
    secure: false
  }
}));

// Jangan-cache halaman & sw.js: mencegah halaman lama (mis. service worker
// BOP Konstruksi dari aplikasi sebelumnya di origin yang sama) tetap tampil.
app.use((req, res, next) => {
  if (req.path === '/sw.js' || /\/(index|login|admin|user|rab|404|sw-cleanup)\.html$/i.test(req.path)) {
    res.setHeader('Cache-Control', 'no-store, no-cache, must-revalidate, max-age=0');
    res.setHeader('Pragma', 'no-cache');
  }
  next();
});

app.use(express.static(path.join(__dirname, 'public'), {
  etag: false,
  lastModified: false,
  setHeaders: (res, filePath) => {
    if (filePath.endsWith('.html')) {
      res.setHeader('Cache-Control', 'no-store, no-cache, must-revalidate, proxy-revalidate, max-age=0, s-maxage=0');
      res.setHeader('Pragma', 'no-cache');
      res.setHeader('Expires', '0');
      res.setHeader('Surrogate-Control', 'no-store');
    }
  }
}));
// Tambahkan baris di bawah ini agar gambar bisa diakses dari web
app.use('/uploads', express.static(UPLOADS_DIR));

function pastikanFolderEksis(dirPath) {
  if (!fs.existsSync(dirPath)) {
    fs.mkdirSync(dirPath, { recursive: true });
  }
}

// ── PDF → GAMBAR (laporan: berkas ABD .pdf / INPUT RFC .pdf dirender per halaman) ──
// Cache PNG: uploads/_pdfcache/<kunci>_p<N>.png; kunci = hash(path+mtime+size) sehingga
// berlaku sampai berkasnya diganti. Gagal render → kembali [] dan nama berkas tetap
// ditampilkan seperti sebelumnya ({"Berkas terlampir: ..."}).
const PDF_CACHE_DIR = path.join(UPLOADS_DIR, '_pdfcache');
let _pdfjsMod = null;
async function muatPdfjs() {
  if (_pdfjsMod) return _pdfjsMod;
  const { pathToFileURL } = require('url');
  const { DOMMatrix, ImageData, Path2D } = require('@napi-rs/canvas');
  globalThis.DOMMatrix = globalThis.DOMMatrix || DOMMatrix;
  globalThis.ImageData = globalThis.ImageData || ImageData;
  globalThis.Path2D = globalThis.Path2D || Path2D;
  _pdfjsMod = await import(pathToFileURL(path.join(__dirname, 'node_modules', 'pdfjs-dist', 'legacy', 'build', 'pdf.mjs')).href);
  return _pdfjsMod;
}

async function renderPdfHalaman(pdfPath) {
  try {
    const st = fs.statSync(pdfPath);
    const kunci = require('crypto').createHash('sha1')
      .update(pdfPath + '|' + st.mtimeMs + '|' + st.size + '|pdf30').digest('hex').slice(0, 16);
    pastikanFolderEksis(PDF_CACHE_DIR);
    const cache = [];
    for (let i = 1; ; i++) {
      const p = path.join(PDF_CACHE_DIR, `${kunci}_p${i}.png`);
      if (!fs.existsSync(p)) break;
      cache.push(p);
    }
    if (cache.length > 0) return cache;

    const pdfjs = await muatPdfjs();
    const { createCanvas } = require('@napi-rs/canvas');
    const doc = await pdfjs.getDocument({
      data: new Uint8Array(fs.readFileSync(pdfPath)),
      // Di Node pdfjs memuat font via fs.readFile, jadi harus path file
      // (slash miring, diakhiri "/") — bukan http:// maupun file://
      standardFontDataUrl: path.join(__dirname, 'node_modules', 'pdfjs-dist', 'standard_fonts').replace(/\\/g, '/') + '/',
      isEvalSupported: false,
      // Pengaman: PDF berbahaya tidak boleh menjalankan JavaScript (CVE pdfjs < 6.2.108)
      enableScripting: false,
    }).promise;
    const hasil = [];
    const maks = Math.min(doc.numPages, 30); // batas wajar, jaga waktu generate
    for (let i = 1; i <= maks; i++) {
      const page = await doc.getPage(i);
      // Skala 3.0 ≈ 216 DPI (dulu 1.6 ≈ 115 DPI) supaya teks/gambar PDF tidak buram saat zoom.
      const viewport = page.getViewport({ scale: 3.0 });
      const canvas = createCanvas(Math.ceil(viewport.width), Math.ceil(viewport.height));
      const ctx = canvas.getContext('2d');
      ctx.fillStyle = '#ffffff';
      ctx.fillRect(0, 0, canvas.width, canvas.height);
      await page.render({ canvasContext: ctx, viewport }).promise;
      const out = path.join(PDF_CACHE_DIR, `${kunci}_p${i}.png`);
      fs.writeFileSync(out, canvas.toBuffer('image/png'));
      hasil.push(out);
    }
    await doc.destroy();
    console.log('[PDF] ' + path.basename(pdfPath) + ' → ' + hasil.length + ' halaman');
    return hasil;
  } catch (e) {
    console.warn('[PDF] gagal render ' + path.basename(pdfPath) + ': ' + e.message);
    return [];
  }
}

// ── DWG/DXF → PDF (AutoCAD 2025 accoreconsole, sudah terpasang di mesin) ──
// Cache PDF: uploads/_dwgcache/<kunci>.pdf (kunci = hash path+mtime+size).
// Gagal → null → nama berkas tetap tampil seperti sebelumnya.
const DWG_CACHE_DIR = path.join(UPLOADS_DIR, '_dwgcache');
const ACCORE = 'C:\\Program Files\\Autodesk\\AutoCAD 2025\\accoreconsole.exe';

function jalankanAccore(dwgPath, scrPath) {
  return new Promise((resolve) => {
    try {
      const { spawn } = require('child_process');
      let selesai = false;
      const anak = spawn(ACCORE, ['/i', dwgPath, '/s', scrPath], {
        cwd: DWG_CACHE_DIR, windowsHide: true, stdio: ['pipe', 'pipe', 'pipe'],
      });
      let keluaran = '';
      const telan = (d) => { try { keluaran += d.toString('utf8'); if (keluaran.length > 8000) keluaran = keluaran.slice(-8000); } catch (e) {} };
      anak.stdout.on('data', telan);
      anak.stderr.on('data', telan);
      try { anak.stdin.end(); } catch (e) {}
      const timer = setTimeout(() => {
        if (selesai) return;
        selesai = true;
        console.warn('[DWG] accoreconsole timeout — ekor log: ' + keluaran.replace(/\s+/g, ' ').slice(-400));
        try { anak.kill(); } catch (e) {}
        resolve(false);
      }, 190000);
      anak.on('error', (e) => {
        if (selesai) return;
        selesai = true; clearTimeout(timer);
        console.warn('[DWG] spawn error: ' + e.message);
        resolve(false);
      });
      anak.on('close', () => {
        if (selesai) return;
        selesai = true; clearTimeout(timer); resolve(true);
      });
    } catch (e) { resolve(false); }
  });
}

async function konversiDwgKePdf(dwgPath) {
  try {
    if (!fs.existsSync(ACCORE)) return null;
    const st = fs.statSync(dwgPath);
    const kunci = require('crypto').createHash('sha1')
      .update(dwgPath + '|' + st.mtimeMs + '|' + st.size).digest('hex').slice(0, 16);
    pastikanFolderEksis(DWG_CACHE_DIR);
    const pdfOut = path.join(DWG_CACHE_DIR, kunci + '.pdf');
    if (fs.existsSync(pdfOut) && fs.statSync(pdfOut).size > 0) return pdfOut;

    const scrPath = path.join(DWG_CACHE_DIR, kunci + '.scr');
    // Hapus PDF parsial (crash sebelumnya) supaya tidak memicu prompt "already exists".
    try { if (fs.existsSync(pdfOut)) fs.unlinkSync(pdfOut); } catch (e) {}
    // Alur prompt accoreconsole (diverifikasi): format → plot area → detailed? → nama file → quit.
    // NB: path PDF WAJIB diapit tanda kutip — AutoCAD script memisahkan token di spasi.
    const skrip = ['-EXPORT', 'PDF', 'Extents', 'No', '"' + pdfOut.replace(/\\/g, '/') + '"', '_QUIT', 'N'];
    fs.writeFileSync(scrPath, skrip.join('\r\n') + '\r\n');
    const jalan = await jalankanAccore(dwgPath, scrPath);
    if (jalan && fs.existsSync(pdfOut) && fs.statSync(pdfOut).size > 0) {
      console.log('[DWG] ' + path.basename(dwgPath) + ' → PDF ok');
      return pdfOut;
    }
    console.warn('[DWG] konversi gagal untuk ' + path.basename(dwgPath));
    return null;
  } catch (e) {
    console.warn('[DWG] error konversi: ' + e.message);
    return null;
  }
}

// Pangkas tepi putih PNG (hasil plot DWG penuh margin) supaya eviden rapat.
function gambarTerpotong(canvas) {
  try {
    const { createCanvas } = require('@napi-rs/canvas');
    const w = canvas.width, h = canvas.height;
    const ctx = canvas.getContext('2d');
    const data = ctx.getImageData(0, 0, w, h).data;
    let x0 = w, y0 = h, x1 = -1, y1 = -1;
    for (let y = 0; y < h; y += 2) {
      for (let x = 0; x < w; x += 2) {
        const i = (y * w + x) * 4;
        if (data[i] < 245 || data[i + 1] < 245 || data[i + 2] < 245) {
          if (x < x0) x0 = x;
          if (x > x1) x1 = x;
          if (y < y0) y0 = y;
          if (y > y1) y1 = y;
        }
      }
    }
    if (x1 < 0 || (x1 - x0) < w * 0.15 || (y1 - y0) < h * 0.15) return canvas;
    const pad = 14;
    x0 = Math.max(0, x0 - pad); y0 = Math.max(0, y0 - pad);
    x1 = Math.min(w - 1, x1 + pad); y1 = Math.min(h - 1, y1 + pad);
    const cw = x1 - x0 + 1, ch = y1 - y0 + 1;
    const out = createCanvas(cw, ch);
    const octx = out.getContext('2d');
    octx.fillStyle = '#ffffff';
    octx.fillRect(0, 0, cw, ch);
    octx.drawImage(canvas, -x0, -y0);
    return out;
  } catch (e) {
    console.warn('[PNG] pangkas gagal: ' + e.message);
    return canvas;
  }
}

// Render PDF hasil plot DWG: skala lebih tajam (3.2) + pangkas tepi putih.
// Cache terpisah: uploads/_dwgcache/<kunci>_p<N>.png
async function renderPdfDwg(pdfPath) {
  try {
    const st = fs.statSync(pdfPath);
    const kunci = require('crypto').createHash('sha1')
      .update(pdfPath + '|' + st.mtimeMs + '|' + st.size + '|dwg32').digest('hex').slice(0, 16);
    pastikanFolderEksis(DWG_CACHE_DIR);
    const cache = [];
    for (let i = 1; ; i++) {
      const p = path.join(DWG_CACHE_DIR, `${kunci}_p${i}.png`);
      if (!fs.existsSync(p)) break;
      cache.push(p);
    }
    if (cache.length > 0) return cache;

    const pdfjs = await muatPdfjs();
    const { createCanvas } = require('@napi-rs/canvas');
    const doc = await pdfjs.getDocument({
      data: new Uint8Array(fs.readFileSync(pdfPath)),
      standardFontDataUrl: path.join(__dirname, 'node_modules', 'pdfjs-dist', 'standard_fonts').replace(/\\/g, '/') + '/',
      isEvalSupported: false,
      enableScripting: false, // lihat renderPdfHalaman — tolak JavaScript di PDF
    }).promise;
    const hasil = [];
    const maks = Math.min(doc.numPages, 10);
    for (let i = 1; i <= maks; i++) {
      const page = await doc.getPage(i);
      // 3.2 ≈ 230 DPI, lebih tajam saat gambar hasil plot DWG di-zoom di Word/PDF.
      const viewport = page.getViewport({ scale: 3.2 });
      const canvas = createCanvas(Math.ceil(viewport.width), Math.ceil(viewport.height));
      const ctx = canvas.getContext('2d');
      ctx.fillStyle = '#ffffff';
      ctx.fillRect(0, 0, canvas.width, canvas.height);
      await page.render({ canvasContext: ctx, viewport }).promise;
      const potong = gambarTerpotong(canvas);
      const out = path.join(DWG_CACHE_DIR, `${kunci}_p${i}.png`);
      fs.writeFileSync(out, potong.toBuffer('image/png'));
      hasil.push(out);
    }
    await doc.destroy();
    console.log('[DWG] render ' + path.basename(pdfPath) + ' → ' + hasil.length + ' gambar (dipangkas)');
    return hasil;
  } catch (e) {
    console.warn('[DWG] gagal render PDF ' + path.basename(pdfPath) + ': ' + e.message);
    return [];
  }
}

// ── .SOR (trace OTDR, Telcordia SR-4731) → GAMBAR grafik trace ──
// Parser: @diolan12/js-otdr. Cache: uploads/_pdfcache/<kunci>_sor.png.
// Gagal parse → null → nama file tetap tampil (fallback lama).
async function renderSorGambar(sorPath) {
  try {
    const st = fs.statSync(sorPath);
    const kunci = require('crypto').createHash('sha1')
      .update(sorPath + '|' + st.mtimeMs + '|' + st.size + '|v2').digest('hex').slice(0, 16);
    pastikanFolderEksis(PDF_CACHE_DIR);
    const png = path.join(PDF_CACHE_DIR, `${kunci}_sor.png`);
    if (fs.existsSync(png)) return png;

    const { SorParser } = require('@diolan12/js-otdr');
    const buf = fs.readFileSync(sorPath);
    const sor = new SorParser(buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength)).parse();
    const fxd = sor.FxdParams || {};
    const gen = sor.GenParams || {};
    const sup = sor.SupParams || {};
    const ke = sor.KeyEvents || {};
    const dp = sor.DataPts || {};
    const db = dp.dbPoints;
    if (!db || !db.length) throw new Error('trace kosong');

    const { createCanvas } = require('@napi-rs/canvas');
    const W = 1600, H = 970;
    const cv = createCanvas(W, H);
    const c = cv.getContext('2d');
    c.fillStyle = '#ffffff'; c.fillRect(0, 0, W, H);
    c.strokeStyle = '#cbd5e1'; c.lineWidth = 2; c.strokeRect(3, 3, W - 6, H - 6);

    // ── Data turunan ──
    // dbPoints = nilai LOSS positif (naik seiring jarak, noise = saturasi).
    // Agar seperti OTDR asli: sumbu Y ditulis positif dengan 0 di bawah,
    // trace turun ke kanan, floor noise di bawah.
    const n = db.length;
    const abs = new Array(n);
    for (let i = 0; i < n; i++) abs[i] = Math.abs(Number(db[i]));
    const urut = abs.slice().sort((a, b) => a - b);
    const pct = (p) => urut[Math.min(n - 1, Math.max(0, Math.round(p * (n - 1))))];
    let lo = pct(0.01), hi = pct(0.995);          // buang outlier (pulse depan)
    if (!(hi - lo > 0.5)) { lo = urut[0]; hi = urut[n - 1]; }
    if (!(hi - lo > 0)) { lo -= 1; hi += 1; }
    const span = hi - lo;
    const langkahKasar = (s) => {
      if (!(s > 0)) return 1;
      const e = Math.pow(10, Math.floor(Math.log10(s)));
      for (const m of [1, 2, 2.5, 5, 10]) if (m * e >= s) return m * e;
      return 10 * e;
    };
    const stepV = langkahKasar(span / 5);
    const topTick = stepV * Math.max(1, Math.ceil(span / stepV));
    const nTick = Math.max(1, Math.round(topTick / stepV));
    const decV = topTick >= 5 ? 0 : (topTick >= 0.5 ? 1 : 2);
    const fmt = (v, d = 2) => (v == null || !isFinite(Number(v))) ? '-' : Number(v).toFixed(d).replace('.', ',');
    const isi = (v, b = '-') => (v == null || String(v).trim() === '') ? b : String(v).trim();

    const unitAsli = String(fxd.unit || 'km');
    const sp = Number(fxd.sampleSpacingUnits) || 0;   // km per titik
    let totalSatuan = sp * n;
    if (!(totalSatuan > 0)) totalSatuan = Number(fxd.traceLengthMeters) || 0;
    let xUnit = unitAsli, total = totalSatuan;
    if (/^km$/i.test(unitAsli) && totalSatuan < 1) { xUnit = 'm'; total = totalSatuan * 1000; }
    const keX = (km) => (/^m$/i.test(xUnit) ? (Number(km) || 0) * 1000 : (Number(km) || 0));
    const decX = total >= 100 ? 0 : total >= 10 ? 1 : total >= 1 ? 2 : 3;

    // ── Utilisasi event / parameter ──
    const evts = (ke.events || []).slice()
      .sort((a, b) => (Number(a.eventNumber) || 0) - (Number(b.eventNumber) || 0));
    const splices = evts.filter((e) => /splice/i.test(String(e.eventType || '')));
    const lossSplice = splices.map((e) => Number(e.spliceLossDb)).filter((v) => isFinite(v) && v > 0);
    const rataSplice = lossSplice.length ? lossSplice.reduce((a, b) => a + b, 0) / lossSplice.length : null;
    const maksSplice = lossSplice.length ? Math.max(...lossSplice) : null;
    const slopes = evts.map((e) => Number(e.slopeDbPerKm)).filter((v) => isFinite(v) && v > 0).sort((a, b) => a - b);
    let attenuasi = slopes.length ? slopes[Math.floor(slopes.length / 2)] : null;
    const idxKm = (km) => Math.max(0, Math.min(n - 1, Math.round((Number(km) || 0) / (sp || 1))));
    if (attenuasi == null && sp > 0 && n > 2) {
      const i1 = Math.floor(n * 0.1), i2 = Math.floor(n * 0.9);
      attenuasi = Math.abs((abs[i2] - abs[i1]) / ((i2 - i1) * sp));
    }
    const eof = evts.filter((e) => /end of fiber/i.test(String(e.eventType || ''))).pop()
      || evts[evts.length - 1] || null;
    const iAwal = evts.length ? idxKm(evts[0].distanceKm) : 0;
    const iAkhir = eof ? idxKm(eof.distanceKm) : n - 1;
    const totalLoss = abs[iAkhir] - abs[iAwal];
    const fiberKm = (eof && eof.distanceKm != null) ? Number(eof.distanceKm) : (sp * n);
    const orl = Number((ke.summary || {}).orlDb) || 0;
    const resolusiM = (/km/i.test(unitAsli) ? totalSatuan * 1000 : totalSatuan) / (n || 1);
    const decRes = resolusiM >= 10 ? 2 : resolusiM >= 1 ? 3 : 4;
    const dObj = fxd.dateTime instanceof Date ? fxd.dateTime : new Date(fxd.dateTime);
    const p2 = (x) => String(x).padStart(2, '0');
    const tgl = isNaN(dObj) ? '-' : dObj.getFullYear() + '-' + p2(dObj.getMonth() + 1) + '-' + p2(dObj.getDate());
    const jam = isNaN(dObj) ? '-' : p2(dObj.getHours()) + ':' + p2(dObj.getMinutes()) + ':' + p2(dObj.getSeconds());

    // ── Judul ──
    c.textAlign = 'left';
    c.fillStyle = '#0f172a'; c.font = 'bold 38px sans-serif';
    c.fillText('TRACE OTDR — ' + path.basename(sorPath), 30, 52);
    c.fillStyle = '#475569'; c.font = '24px sans-serif';
    c.fillText('Kabel: ' + isi(gen.cableId || gen.cableCode) + '  ·  Serat: ' + isi(gen.fiberId) +
      '  ·  ' + (fxd.wavelengthNm || '-') + ' nm  ·  Pulse ' + (fxd.pulseWidthNs || '-') + ' ns  ·  ' +
      fmt(total, decX) + ' ' + xUnit + '  ·  ' + n + ' titik  ·  ' + evts.length + ' event', 30, 86);

    // ── Chart ──
    const L = 110, R = 1150, T = 100, B = 620;
    const cw = R - L, ch = B - T;
    c.fillStyle = '#ffffff'; c.fillRect(L, T, cw, ch);
    c.strokeStyle = '#94a3b8'; c.lineWidth = 1.5; c.setLineDash([]);
    c.strokeRect(L, T, cw, ch);

    c.strokeStyle = '#e2e8f0'; c.lineWidth = 1; c.setLineDash([5, 5]);
    c.font = '26px sans-serif';
    for (let g = 0; g <= nTick; g++) {
      const y = B - (ch * g) / nTick;
      c.beginPath(); c.moveTo(L, y); c.lineTo(R, y); c.stroke();
      c.textAlign = 'right'; c.fillStyle = '#334155';
      c.fillText(fmt(g * stepV, decV), L - 12, y + 9);
    }
    for (let g = 0; g <= 10; g++) {
      const x = L + (cw * g) / 10;
      if (g > 0 && g < 10) { c.beginPath(); c.moveTo(x, T); c.lineTo(x, B); c.stroke(); }
      c.textAlign = g === 0 ? 'left' : (g === 10 ? 'right' : 'center');
      c.fillStyle = '#334155';
      c.fillText(fmt((total * g) / 10, decX), g === 0 ? L : (g === 10 ? R : x), B + 34);
    }
    c.setLineDash([]);
    c.textAlign = 'center'; c.fillStyle = '#0f172a'; c.font = 'bold 28px sans-serif';
    c.fillText('Jarak (' + xUnit + ')', L + cw / 2, B + 68);
    c.save(); c.translate(38, T + ch / 2); c.rotate(-Math.PI / 2);
    c.fillText('Level (dB)', 0, 0); c.restore();

    // ── Trace (loss → ditampilkan terbalik: makin besar loss, makin ke bawah) ──
    const yDari = (a) => Math.max(T, Math.min(B, B - ((hi - a) / span) * ch));
    c.strokeStyle = '#111827'; c.lineWidth = 1.5; c.lineJoin = 'round';
    c.beginPath();
    const stepI = Math.max(1, Math.ceil(n / 20000));
    for (let i = 0; i < n; i += stepI) {
      const x = L + (i / (n - 1 || 1)) * cw, y = yDari(abs[i]);
      if (i === 0) c.moveTo(x, y); else c.lineTo(x, y);
    }
    c.stroke();

    // ── Penanda event (garis merah + nomor, anti tumpang tindih) ──
    const xsEvent = evts.map((e) => L + (keX(e.distanceKm) / (total || 1)) * cw);
    const barisLingkaran = [];
    evts.forEach((e, i) => {
      const x = xsEvent[i];
      if (x < L || x > R) return;
      let r = 0;
      while (barisLingkaran[r] != null && x - barisLingkaran[r] < 34) r++;
      barisLingkaran[r] = x;
      const y0 = T + 15 + r * 30;
      c.strokeStyle = '#dc2626'; c.lineWidth = 1.3; c.setLineDash([6, 5]);
      c.beginPath(); c.moveTo(x, T); c.lineTo(x, B); c.stroke();
      c.setLineDash([]);
      c.fillStyle = '#dc2626';
      c.beginPath(); c.arc(x, y0, 13, 0, Math.PI * 2); c.fill();
      c.fillStyle = '#fff'; c.font = 'bold 15px sans-serif'; c.textAlign = 'center';
      c.fillText(String(e.eventNumber != null ? e.eventNumber : i + 1), x, y0 + 5);
    });

    // ── Panel kanan (Parameter + Info) ──
    const PX0 = 1170, PX1 = 1580, PW = PX1 - PX0;
    const gambarPanel = (judul, baris, y0) => {
      c.fillStyle = '#111827'; c.fillRect(PX0, y0, PW, 34);
      c.fillStyle = '#fff'; c.textAlign = 'left'; c.font = 'bold 24px sans-serif';
      c.fillText(judul, PX0 + 12, y0 + 24);
      let y = y0 + 34;
      baris.forEach(([k, v], i) => {
        c.fillStyle = i % 2 ? '#ffffff' : '#f1f5f9';
        c.fillRect(PX0, y, PW, 29);
        c.font = '24px sans-serif';
        c.textAlign = 'left'; c.fillStyle = '#334155'; c.fillText(k, PX0 + 12, y + 20);
        c.textAlign = 'right'; c.fillStyle = '#0f172a';
        const maksV = PW - 28 - c.measureText(k).width;
        let vs = String(v);
        while (vs.length > 1 && c.measureText(vs).width > maksV) vs = vs.slice(0, -1);
        if (vs !== String(v)) vs = vs.slice(0, -2) + '…';
        c.fillText(vs, PX1 - 12, y + 20);
        c.strokeStyle = '#e2e8f0'; c.lineWidth = 1;
        c.beginPath(); c.moveTo(PX0, y + 29.5); c.lineTo(PX1, y + 29.5); c.stroke();
        y += 29;
      });
      return y;
    };

    gambarPanel('PARAMETER', [
      ['Panjang Serat', fmt(fiberKm, 3) + ' km'],
      ['Rentang', fmt(totalSatuan, 3) + ' ' + unitAsli],
      ['Panjang Gelombang', (fxd.wavelengthNm || '-') + ' nm'],
      ['Lebar Pulse', (fxd.pulseWidthNs || '-') + ' ns'],
      ['Redaman', fmt(attenuasi, 3) + ' dB/km'],
      ['Total Loss', fmt(totalLoss, 3) + ' dB'],
      ['ORL', orl > 0 ? fmt(orl, 2) + ' dB' : '-'],
      ['Jumlah Splice', String(splices.length)],
      ['Rata-rata loss splice', rataSplice != null ? fmt(rataSplice, 3) + ' dB' : '-'],
      ['Loss splice maks', maksSplice != null ? fmt(maksSplice, 3) + ' dB' : '-'],
      ['Jumlah Sampel', String(fxd.numDataPoints || n)],
      ['Resolusi', fmt(resolusiM, decRes) + ' m'],
      ['Indeks Bias (IOR)', fmt(fxd.ior, 6)],
      ['Ambang Loss', fmt(fxd.lossThreshold, 2) + ' dB'],
      ['Ambang Refleksi', fmt(fxd.reflectionThreshold, 2) + ' dB'],
    ], 90);

    gambarPanel('INFO ALAT & KABEL', [
      ['Kabel', isi(gen.cableId || gen.cableCode)],
      ['Serat', isi(gen.fiberId)],
      ['Tipe Serat', (gen.fiberType && gen.fiberType !== 'unknown') ? gen.fiberType : '-'],
      ['Tanggal', tgl],
      ['Waktu', jam],
      ['Operator', isi(gen.operator)],
      ['OTDR', isi(sup.otdr || sup.supplier)],
      ['SN OTDR', isi(sup.otdrSn)],
      ['Modul', isi(sup.module)],
      ['Software', isi(sup.software)],
      ['Komentar', isi(gen.comment)],
    ], 579);

    // ── Tabel event ──
    c.textAlign = 'left'; c.fillStyle = '#0f172a'; c.font = 'bold 26px sans-serif';
    c.fillText('EVENT (' + evts.length + ')', 30, 716);
    const kolom = [
      ['Tipe', 230], ['№', 60], ['Jarak (' + xUnit + ')', 170], ['Loss (dB)', 150],
      ['Refleksi (dB)', 190], ['Slope (dB/km)', 180], ['Komentar', 140],
    ];
    const HT = 724, RH = 30, maksBaris = 7;
    let xg = 30;
    const xKol = kolom.map(([h, w]) => { const x = xg; xg += w; return [x, w, h]; });
    const lebarTabel = xg - 30;
    const txKol = (x, w, h) => (h === '№') ? x + w / 2 : (x === 30 ? x + 12 : x + w - 12);
    c.fillStyle = '#111827'; c.fillRect(30, HT, lebarTabel, 34);
    c.font = 'bold 24px sans-serif'; c.fillStyle = '#fff';
    xKol.forEach(([x, w, h]) => {
      c.textAlign = h === '№' ? 'center' : (x === 30 ? 'left' : 'right');
      c.fillText(h, txKol(x, w, h), HT + 24);
    });
    let yBaris = HT + 34;
    const lebih = evts.length > maksBaris;
    const tampil = evts.slice(0, maksBaris - (lebih ? 1 : 0));
    for (let i = 0; i < maksBaris; i++) {
      c.fillStyle = i % 2 ? '#fff' : '#f8fafc';
      c.fillRect(30, yBaris, lebarTabel, RH);
      const e = tampil[i];
      if (e) {
        c.font = '24px sans-serif'; c.fillStyle = '#0f172a';
        const isiKol = [
          isi(e.eventType),
          e.eventNumber != null ? String(e.eventNumber) : String(i + 1),
          fmt(keX(e.distanceKm), 3),
          Number(e.spliceLossDb) ? fmt(e.spliceLossDb, 3) : '-',
          Number(e.reflectionLossDb) ? fmt(e.reflectionLossDb, 2) : '-',
          Number(e.slopeDbPerKm) ? fmt(e.slopeDbPerKm, 3) : '-',
          isi(e.comment),
        ];
        xKol.forEach(([x, w, h], j) => {
          c.textAlign = h === '№' ? 'center' : (x === 30 ? 'left' : 'right');
          c.fillText(isiKol[j], txKol(x, w, h), yBaris + 21);
        });
      } else if (i === 0 && !evts.length) {
        c.font = '24px sans-serif'; c.fillStyle = '#64748b'; c.textAlign = 'left';
        c.fillText('Tidak ada data event pada file ini.', 42, yBaris + 21);
      } else if (lebih && i === maksBaris - 1) {
        c.font = '24px sans-serif'; c.fillStyle = '#64748b'; c.textAlign = 'left';
        c.fillText('… + ' + (evts.length - tampil.length) + ' event lainnya', 42, yBaris + 21);
      }
      yBaris += RH;
    }
    c.strokeStyle = '#cbd5e1'; c.lineWidth = 1.5; c.setLineDash([]);
    c.strokeRect(30, HT, lebarTabel, yBaris - HT);

    fs.writeFileSync(png, cv.toBuffer('image/png'));
    console.log('[SOR] ' + path.basename(sorPath) + ' → grafik trace ok');
    return png;
  } catch (e) {
    console.warn('[SOR] gagal render ' + path.basename(sorPath) + ': ' + e.message);
    return null;
  }
}

async function init() {
  if (!fs.existsSync(UPLOADS_DIR)) fs.mkdirSync(UPLOADS_DIR, { recursive: true });
  await db.exec(`CREATE TABLE IF NOT EXISTS users (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    username TEXT UNIQUE,
    password TEXT,
    role TEXT
  )`);

  // ── LOGIN TEMATIK: daftar tematik yang boleh dibuka tiap akun ──────────────
  // Isi kolom: JSON array, mis. ["HEM","MICRODEMAND"].
  // Kosong/null = akun BELUM punya tematik → tidak boleh mengakses data apa pun
  // (admin tidak terpengaruh: selalu bebas).
  try { await db.exec(`ALTER TABLE users ADD COLUMN tematik TEXT`); } catch (e) { /* kolom sudah ada */ }

  await db.exec(`CREATE TABLE IF NOT EXISTS sessions (
    id TEXT PRIMARY KEY,
    data TEXT,
    updated_at TEXT
  )`);

  await db.exec(`CREATE TABLE IF NOT EXISTS lokasi (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    pid TEXT,
    nama_lokasi TEXT UNIQUE,
    tematik TEXT,
    detail_tambahan TEXT,
    status_pekerjaan TEXT,
    no_sp TEXT
  )`);
  try { await db.exec(`ALTER TABLE lokasi ADD COLUMN no_sp TEXT`); } catch (e) { /* kolom sudah ada */ }
  try { await db.exec(`ALTER TABLE lokasi ADD COLUMN status_eviden TEXT DEFAULT 'need'`); } catch (e) { /* kolom sudah ada */ }

  await db.exec(`CREATE TABLE IF NOT EXISTS berkas (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    lokasi_id INTEGER,
    kategori TEXT,
    sub_kategori TEXT,
    file_name TEXT
  )`);

  // ── TABEL OPM: rekam ODP + kapasitas + status tiap port (ukur/isi) ──
  // status 'ukur'  = port kosong, ada foto eviden (masuk ke berkas)
  // status 'isi'   = port ada user, tidak ada foto (tulisan "ISI" di laporan)
  await db.exec(`CREATE TABLE IF NOT EXISTS opm_odp (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    lokasi_id INTEGER NOT NULL,
    odp TEXT NOT NULL,
    kapasitas INTEGER DEFAULT 8,
    UNIQUE(lokasi_id, odp)
  )`);

  await db.exec(`CREATE TABLE IF NOT EXISTS opm_port (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    lokasi_id INTEGER NOT NULL,
    odp TEXT NOT NULL,
    port INTEGER NOT NULL,
    status TEXT NOT NULL DEFAULT 'ukur',
    UNIQUE(lokasi_id, odp, port)
  )`);

  // ── TABEL REKAP OPM RELOKASI (nilai hasil ukur per port, bisa dikoreksi manual) ──
  await db.exec(`CREATE TABLE IF NOT EXISTS rekap_opm (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    lokasi_id INTEGER NOT NULL,
    odp TEXT NOT NULL,
    port INTEGER NOT NULL,
    nilai TEXT DEFAULT '',
    UNIQUE(lokasi_id, odp, port)
  )`);
  await db.exec(`CREATE TABLE IF NOT EXISTS rekap_opm_meta (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    lokasi_id INTEGER UNIQUE NOT NULL,
    catuan TEXT DEFAULT '-',
    updated_at DATETIME DEFAULT CURRENT_TIMESTAMP
  )`);
  // ODP mana saja yang dipilih user untuk dimasukkan ke lampiran rekap OPM
  await db.exec(`CREATE TABLE IF NOT EXISTS rekap_opm_pilih (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    lokasi_id INTEGER NOT NULL,
    odp TEXT NOT NULL,
    UNIQUE(lokasi_id, odp)
  )`);

  // ── TABEL REDAMAN ODP/OTB: rekam ODP + kapasitas + status tiap port ──
  // (mirip OPM, terpisah supaya EVIDEN REDAMAN tidak tercampur dengan
  //  EVIDENT HASIL UKUR OPM pada ODP yang sama)
  await db.exec(`CREATE TABLE IF NOT EXISTS redaman_odp (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    lokasi_id INTEGER NOT NULL,
    odp TEXT NOT NULL,
    kapasitas INTEGER DEFAULT 8,
    UNIQUE(lokasi_id, odp)
  )`);

  await db.exec(`CREATE TABLE IF NOT EXISTS redaman_port (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    lokasi_id INTEGER NOT NULL,
    odp TEXT NOT NULL,
    port INTEGER NOT NULL,
    status TEXT NOT NULL DEFAULT 'ukur',
    UNIQUE(lokasi_id, odp, port)
  )`);

  // ── TABEL BOQ AKTUAL: baris hasil paste Excel (seperti input RAB) ──
  await db.exec(`CREATE TABLE IF NOT EXISTS boq_aktual (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    lokasi_id INTEGER NOT NULL,
    designator TEXT,
    uraian TEXT,
    satuan TEXT,
    material REAL DEFAULT 0,
    jasa REAL DEFAULT 0,
    volume REAL DEFAULT 0,
    created_at TEXT DEFAULT (datetime('now','localtime'))
  )`);

  // ── TABEL BOQ BARANG TIBA (input TEMATIK Node B): no, jenis material,
  //    satuan, volume, keterangan — tampil sebagai tabel di TAMPLATE_NODE_B ──
  await db.exec(`CREATE TABLE IF NOT EXISTS barang_tiba (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    lokasi_id INTEGER NOT NULL,
    no INTEGER,
    jenis_material TEXT,
    satuan TEXT,
    volume TEXT,
    keterangan TEXT,
    created_at TEXT DEFAULT (datetime('now','localtime'))
  )`);

  // ── TABEL CLOSURE CUT OVER: nama + jenis closure (24/48/96/dll) ──
  await db.exec(`CREATE TABLE IF NOT EXISTS co_closure (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    lokasi_id INTEGER NOT NULL,
    nama TEXT NOT NULL,
    jenis TEXT,
    UNIQUE(lokasi_id, nama)
  )`);

  await db.exec(`CREATE TABLE IF NOT EXISTS laporan_rekon (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    lokasi_id INTEGER UNIQUE,
    proyek TEXT,
    kontrak TEXT,
    surat_pesanan TEXT,
    district TEXT,
    pelaksana TEXT,
    hari_ct TEXT,
    tgl TEXT,
    bulan TEXT,
    nama_waspang TEXT,
    nik_waspang TEXT,
    lokasi_smile TEXT,
    ct_hari TEXT,
    ct_tgl TEXT,
    ct_bulan TEXT,
    dismantling_hari TEXT,
    dismantling_tanggal TEXT,
    dismantling_bulan TEXT,
    matrial_aset1 TEXT, volume_aset1 TEXT, kondisi_aset1 TEXT,
    matrial_aset2 TEXT, volume_aset2 TEXT, kondisi_aset2 TEXT,
    matrial_aset3 TEXT, volume_aset3 TEXT, kondisi_aset3 TEXT,
    matrial_aset4 TEXT, volume_aset4 TEXT, kondisi_aset4 TEXT,
    matrial_aset5 TEXT, volume_aset5 TEXT, kondisi_aset5 TEXT,
    ibooster_hari TEXT, ibooster_tanggal TEXT, ibooster_bulan TEXT,
    ibooster_lokasi TEXT
  )`);

  // ── TABEL RAB (Rencana Anggaran Biaya) — BARU ──────────────────────────────
  await db.exec(`CREATE TABLE IF NOT EXISTS rab (
    id            INTEGER PRIMARY KEY AUTOINCREMENT,
    lokasi_id     INTEGER NOT NULL,
    no            INTEGER,
    designator    TEXT NOT NULL,
    des_material  TEXT,
    des_jasa      TEXT,
    kategori      TEXT DEFAULT 'Gabungan',
    satuan        TEXT,
    material      REAL DEFAULT 0,
    jasa          REAL DEFAULT 0,
    volume        REAL DEFAULT 0,
    created_at    TEXT DEFAULT (datetime('now','localtime')),
    FOREIGN KEY (lokasi_id) REFERENCES lokasi(id)
  )`);

    await db.exec(`CREATE TABLE IF NOT EXISTS rab_rekon (
    id            INTEGER PRIMARY KEY AUTOINCREMENT,
    lokasi_id     INTEGER NOT NULL,
    designator    TEXT,
    drm           REAL DEFAULT 0,
    rekon         REAL DEFAULT 0,
    tambah        REAL DEFAULT 0,
    kurang        REAL DEFAULT 0,
    satuan        TEXT,
    FOREIGN KEY (lokasi_id) REFERENCES lokasi(id)
  )`);

  // ── MIGRASI: tambah kolom baru jika tabel laporan_rekon sudah ada sebelumnya ──
  const kolomBaru = ['ct_hari', 'ct_tgl', 'ct_bulan'];
  for (const kolom of kolomBaru) {
    try {
      await db.run(`ALTER TABLE laporan_rekon ADD COLUMN ${kolom} TEXT`);
      console.log(` Kolom "${kolom}" berhasil ditambahkan ke laporan_rekon.`);
    } catch (e) {
      // Kolom sudah ada — lewati saja
    }
  }
// ── MIGRASI KHUSUS BA TAMBAH KURANG ──
  const kolomTK = ['proyektk', 'kontraktk', 'sptk', 'nosptk', 'haritk', 'tgltk', 'blntk', 'thntk', 'loktk'];
  for (const kolom of kolomTK) {
    try { 
      await db.run(`ALTER TABLE laporan_rekon ADD COLUMN ${kolom} TEXT`); 
    } catch (e) {}
  }

  // ── MIGRASI INPUT TEMATIK NODE B (LACT) — field laporan tambahan ──
  const kolomNodeB = [
    'waspang_ta', 'nik_waspang_ta',
    'tgl_bulan_thn_ct',
    'hari_barang_tiba', 'tgl_barang_tiba', 'bulan_barang_tiba',
    'tgl_bulan_thn_barang_tiba', 'tgl_bulan_thn_opm',
    'pm_area', 'site', 'ruas_pekerjaan', 'izin'
  ];
  for (const kolom of kolomNodeB) {
    try {
      await db.run(`ALTER TABLE laporan_rekon ADD COLUMN ${kolom} TEXT`);
      console.log(` Kolom "${kolom}" berhasil ditambahkan ke laporan_rekon.`);
    } catch (e) { /* Kolom sudah ada — lewati */ }
  }

  // ── MIGRASI KHUSUS BA CO (Laporan Cut Over) ──
  const kolomBAco = ['harico', 'tglco', 'bulanco', 'sambungan1co', 'sambungan2co', 'sambungan3co', 'sambungan4co'];
  for (const kolom of kolomBAco) {
    try {
      await db.run(`ALTER TABLE laporan_rekon ADD COLUMN ${kolom} TEXT`);
    } catch (e) {}
  }

  // ── MIGRASI LAPORAN SERAH TERIMA HASIL DISMANTLING (TEMATIK RELOKASI) ──
  // gudang_sto       → mengisi placeholder {STO} ("Gudang STO ...")
  // pilih_dismantling→ 'DISMANTLING' / 'TIDAK DISMANTLING' (yang tidak dipilih dicoret)
  const kolomDismantling = ['gudang_sto', 'pilih_dismantling'];
  for (const kolom of kolomDismantling) {
    try {
      await db.run(`ALTER TABLE laporan_rekon ADD COLUMN ${kolom} TEXT`);
      console.log(` Kolom "${kolom}" berhasil ditambahkan ke laporan_rekon.`);
    } catch (e) { /* kolom sudah ada — lewati */ }
  }

  // ── MIGRASI: tambah kolom harga_satuan ke rab_rekon jika belum ada ──      ← TAMBAHAN BARU
  try {
    await db.run(`ALTER TABLE rab_rekon ADD COLUMN harga_satuan REAL DEFAULT 0`);
    console.log(' Kolom "harga_satuan" berhasil ditambahkan ke rab_rekon.');
  } catch (e) { /* sudah ada, abaikan */ }

  // ── MIGRASI: sumber harga (AMD4/AMD5) yang dipakai saat rekon disimpan ──
  // Dipakai supaya download dari menu mana pun memakai master yang sama
  // dengan saat hitung realisasi (tidak lagi default ke AMD4).
  let kolomSumberBaru = false;
  try {
    await db.run(`ALTER TABLE rab_rekon ADD COLUMN sumber_master TEXT DEFAULT 'amd4'`);
    kolomSumberBaru = true;
    console.log(' Kolom "sumber_master" berhasil ditambahkan ke rab_rekon.');
  } catch (e) { /* sudah ada */ }

  if (kolomSumberBaru) {
    // Data lama belum punya sumber master. Cocokkan harga tersimpan dengan
    // kedua BOQ TAMKUR; bila jelas lebih cocok ke AMD5, tandai AMD5.
    try {
      const kunciHarga = sumber => {
        const out = {};
        const peta = bacaHargaBoqTamkur(sumber);
        Object.keys(peta).forEach(k => {
          const kk = String(k).toUpperCase().replace(/[^A-Z0-9]/g, '');
          if (kk && out[kk] == null) out[kk] = peta[k];
        });
        return out;
      };
      const p4 = kunciHarga('amd4'), p5 = kunciHarga('amd5');
      const lokasiList = await db.all('SELECT DISTINCT lokasi_id FROM rab_rekon');
      for (const rowLok of lokasiList) {
        const items = await db.all(
          'SELECT designator, harga_satuan FROM rab_rekon WHERE lokasi_id = ?', [rowLok.lokasi_id]
        );
        let cocok4 = 0, cocok5 = 0;
        items.forEach(it => {
          const k = String(it.designator || '').toUpperCase().replace(/[^A-Z0-9]/g, '');
          const h4 = p4[k], h5 = p5[k];
          if (typeof h4 !== 'number' || typeof h5 !== 'number' || h4 === h5) return;
          const hs = Number(it.harga_satuan) || 0;
          if (Math.abs(hs - h4) < 0.5) cocok4++;
          if (Math.abs(hs - h5) < 0.5) cocok5++;
        });
        if (cocok5 > cocok4) {
          await db.run('UPDATE rab_rekon SET sumber_master = ? WHERE lokasi_id = ?',
            ['amd5', rowLok.lokasi_id]);
        }
      }
    } catch (e) {
      console.warn(' [rab_rekon] gagal menginferensi sumber_master lama:', e.message);
    }
  }

  // ── MIGRASI: tambah created_at ke rab jika belum ada ──
  try { await db.run(`ALTER TABLE rab ADD COLUMN created_at TEXT DEFAULT (datetime('now','localtime'))`); } catch {}

  // ── TABEL INPUT KHUSUS (BOQ AKTUAL / INPUT RFC / EVIDEN BON MATERIAL) ──
  await db.exec(`CREATE TABLE IF NOT EXISTS input_khusus (
    id         INTEGER PRIMARY KEY AUTOINCREMENT,
    lokasi_id  INTEGER NOT NULL,
    kategori   TEXT NOT NULL,
    judul      TEXT,
    volume     TEXT,
    isi        TEXT,
    created_at TEXT DEFAULT (datetime('now','localtime')),
    updated_at TEXT
  )`);

  // ── MIGRASI: jejak "terakhir diubah" pada input khusus (fitur edit & simpan ulang) ──
  try {
    await db.run(`ALTER TABLE input_khusus ADD COLUMN updated_at TEXT`);
    console.log(' Kolom "updated_at" berhasil ditambahkan ke input_khusus.');
  } catch (e) { /* kolom sudah ada — lewati */ }


  // ── TABEL RAB SUMMARY — cache nilai RAB per lokasi ─────────────────────
  await db.exec(`CREATE TABLE IF NOT EXISTS rab_summary (
    lokasi_id      INTEGER PRIMARY KEY,
    total_material REAL    DEFAULT 0,
    total_jasa     REAL    DEFAULT 0,
    grand_total    REAL    DEFAULT 0,
    rab_count      INTEGER DEFAULT 0,
    updated_at     TEXT    DEFAULT (datetime('now','localtime')),
    FOREIGN KEY (lokasi_id) REFERENCES lokasi(id)
  )`);
  // Migrasi kolom rab_summary jika tabel sudah ada tapi kolom belum
  for (const [col, def] of [['total_material','REAL DEFAULT 0'],['total_jasa','REAL DEFAULT 0'],['grand_total','REAL DEFAULT 0'],['rab_count','INTEGER DEFAULT 0']]) {
    try { await db.run(`ALTER TABLE rab_summary ADD COLUMN ${col} ${def}`); } catch(e) {}
  }
  console.log(' Tabel rab_summary siap.');

  // ── TABEL AKTIVITAS — log dokumen yang digenerate (widget Dashboard) ─────
  await db.exec(`CREATE TABLE IF NOT EXISTS aktivitas (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    jenis       TEXT    NOT NULL,
    lokasi_id   INTEGER,
    lokasi_nama TEXT,
    pid         TEXT,
    status      TEXT    DEFAULT 'Selesai',
    timestamp   TEXT    DEFAULT (datetime('now','localtime'))
  )`);
  console.log(' Tabel aktivitas siap.');

  // ── TABEL DOKUMEN STATUS — status 6 dokumen pemberkasan per LOP ────────────
  await db.exec(`CREATE TABLE IF NOT EXISTS dokumen_status (
    id           INTEGER PRIMARY KEY AUTOINCREMENT,
    lokasi_id    INTEGER NOT NULL,
    jenis        TEXT    NOT NULL,
    status       TEXT    NOT NULL DEFAULT 'BELUM',
    generated_at TEXT,
    file_path    TEXT,
    updated_by   TEXT,
    updated_at   TEXT    DEFAULT (datetime('now','localtime')),
    UNIQUE (lokasi_id, jenis)
  )`);
  console.log(' Tabel dokumen_status siap.');
  try {
    const nSeed = await seedDokumenStatus();
    if (nSeed) console.log(' Seed dokumen_status: +' + nSeed + ' baris BELUM.');
  } catch (e) {
    console.warn('seed dokumen_status:', e.message);
  }

  // ── TABEL AUDIT SKIP — berkas "tidak berlaku" per lokasi (panel Audit Berkas) ──
  await db.exec(`CREATE TABLE IF NOT EXISTS audit_skip (
    id         INTEGER PRIMARY KEY AUTOINCREMENT,
    lokasi_id  INTEGER NOT NULL,
    item       TEXT    NOT NULL,
    skipped_by TEXT,
    created_at TEXT    DEFAULT (datetime('now','localtime')),
    UNIQUE (lokasi_id, item)
  )`);
  console.log(' Tabel audit_skip siap.');

  // ── TABEL VERIFIKASI EVIDEN — hasil verifikasi keaslian per foto ─────────
  try {
    await verif.initTabel(db);
  } catch (e) {
    console.warn('verifikasi_eviden:', e.message);
  }

  // ── TABEL CACHE GEOCODING — koordinat → nama wilayah (Nominatim) ─────────
  try {
    await geoCheck.initTabel(db);
  } catch (e) {
    console.warn('geo_cache:', e.message);
  }

  // ── TABEL PROJECT PLAN — form input rencana proyek (designator + evidensi) ─
  await db.exec(`CREATE TABLE IF NOT EXISTS project_plan (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    nama_lokasi TEXT NOT NULL,
    created_at DATETIME DEFAULT CURRENT_TIMESTAMP
  )`);

  await db.exec(`CREATE TABLE IF NOT EXISTS project_plan_item (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    project_id INTEGER NOT NULL,
    designator TEXT,
    jenis TEXT,
    volume REAL,
    satuan TEXT,
    FOREIGN KEY (project_id) REFERENCES project_plan(id)
  )`);

  await db.exec(`CREATE TABLE IF NOT EXISTS project_plan_evidensi (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    project_id INTEGER NOT NULL,
    keterangan TEXT,
    file_path TEXT,
    FOREIGN KEY (project_id) REFERENCES project_plan(id)
  )`);
  try { await db.run('ALTER TABLE project_plan_evidensi ADD COLUMN kategori TEXT'); } catch (e) { /* kolom sudah ada */ }
  // eviden melekat per baris designator (NULL = eviden umum tanpa baris)
  try { await db.run('ALTER TABLE project_plan_evidensi ADD COLUMN item_id INTEGER'); } catch (e) { /* kolom sudah ada */ }
  console.log(' Tabel project_plan siap.');

  const userCount = await db.get('SELECT COUNT(*) as c FROM users');
  if (userCount.c === 0) {
    const hashAdmin = await bcrypt.hash('admin123', 10);
    const hashUser  = await bcrypt.hash('user123', 10);
    await db.run('INSERT INTO users (username, password, role) VALUES (?, ?, ?)', ['admin', hashAdmin, 'admin']);
    await db.run('INSERT INTO users (username, password, role) VALUES (?, ?, ?)', ['field_user', hashUser, 'user']);
    console.log(' Akun bawaan berhasil dibuat (admin:admin123, field_user:user123)');
  }
}

// ── DOKUMEN PEREKASAN — Pick up berkas + 6 dokumen, statusnya per LOP ────────
// Status: BELUM | TERGENERATE | NA. Terisi otomatis saat generate lewat aplikasi,
// bisa dikoreksi manual oleh admin (PUT /api/dokumen-status).
// `pickup` bukan dokumen: menandai lokasi sudah diambil/dikerjakan (status eviden
// progres/done), jadi tidak ikut menghitung "Done Pemberkasan".
const DOK_STATUS_JENIS  = ['pickup', 'lact', 'baut', 'ba_prelim', 'ba_perubahan', 'mom_redesign', 'ba_tambahkurang'];
const DOK_STATUS_LABEL  = {
  pickup: 'Pick up berkas',
  lact: 'LACT', baut: 'BAUT', ba_prelim: 'BA PRELIM', ba_perubahan: 'BA REDISGN',
  mom_redesign: 'MOM REDISGN', ba_tambahkurang: 'BA TAMBAH KURANG'
};
const DOK_STATUS_VALUES = ['BELUM', 'TERGENERATE', 'NA'];

// Laporan_*.docx (menu LACT HEM/MICRODEMAND/NODE B, tombol Unduh Audit Berkas,
// Simpan & Cetak Word RAB) tercatat di tabel aktivitas sebagai "microdemand_hem";
// untuk checklist Progres dokumen itu dihitung sebagai LACT.
const DOK_DARI_AKTIVITAS = { microdemand_hem: 'lact' };

// Gabungkan proyek trigger (butuh eviden) dengan master lokasi & No. SP.
async function ambilLopTrigger() {
  const rows  = await db.all('SELECT sumber_rev, branch, nama_lop, tematik_project, status_progress, butuh_eviden, po_non_po, nilai_total_drm, drm_nilai FROM godc_project');
  const items = godcSync.uniqByLokasi(rows).filter(needEvidence);
  const spRows = await db.all("SELECT nama_lokasi, no_sp FROM godc_sp WHERE COALESCE(no_sp,'') <> ''");
  const spMap  = new Map(spRows.map(r => [godcSync.normName(r.nama_lokasi), r.no_sp]));
  const lokRows = await db.all('SELECT id, nama_lokasi FROM lokasi');
  const lokMap  = new Map(lokRows.map(l => [godcSync.normName(l.nama_lokasi), l.id]));
  return items.map(it => ({
    nama: it.nama_lop,
    branch: godcSync.normBranch(it.branch),
    tematik: godcTematikRowServer(it.tematik_project),
    sumber: (it.sumber_rev || '').trim() || 'Data Master',
    po: godcSync.poClass(it.po_non_po) === 'PO' ? 'PO' : 'NON PO',
    sp: spMap.get(godcSync.normName(it.nama_lop)) || '',
    drm: Number(it.drm_nilai) || Number(it.nilai_total_drm) || 0,
    lokasi_id: lokMap.get(godcSync.normName(it.nama_lop)) || null
  }));
}

// Peta status dokumen: Map(lokasi_id -> { jenis -> {status, generated_at} })
async function petaDokumen() {
  const rows = await db.all('SELECT lokasi_id, jenis, status, generated_at FROM dokumen_status');
  const m = new Map();
  rows.forEach(r => {
    if (!m.has(r.lokasi_id)) m.set(r.lokasi_id, {});
    m.get(r.lokasi_id)[r.jenis] = { status: r.status, generated_at: r.generated_at || null };
  });
  return m;
}

// Sisipkan baris 'BELUM' untuk setiap LOP x jenis yang belum punya baris.
// Idempoten — aman dipanggil berkali-kali (dipanggil saat boot & saat rekap dibuka).
async function seedDokumenStatus() {
  const tabel = await db.get("SELECT name FROM sqlite_master WHERE type='table' AND name='godc_project'");
  if (!tabel) return 0;
  const lops = await ambilLopTrigger();
  const ids  = Array.from(new Set(lops.map(l => l.lokasi_id).filter(Boolean)));
  if (!ids.length) return 0;
  const sudahRows = await db.all(
    `SELECT lokasi_id, jenis FROM dokumen_status WHERE lokasi_id IN (${ids.map(() => '?').join(',')})`,
    ids
  );
  const sudah = new Set(sudahRows.map(r => r.lokasi_id + '|' + r.jenis));
  let tambah = 0;
  // SELURUH baris disimpan dalam SATU transaksi. Tanpa ini tiap baris memanggil
  // saveDb() yang mengekspor penuh database (~13 MB) ke disk — saat boot ada
  // belasan ribu baris, event loop tersendat belasan menit dan localhost:3000
  // tidak merespons (koneksi menggantung) sampai seed selesai.
  await db.batch(async () => {
    for (const id of ids) {
      for (const j of DOK_STATUS_JENIS) {
        if (sudah.has(id + '|' + j)) continue;
        await db.run("INSERT OR IGNORE INTO dokumen_status (lokasi_id, jenis, status) VALUES (?, ?, 'BELUM')", [id, j]);
        tambah++;
      }
    }
  });

  // Pick up berkas: LOP yang status evidennya sudah progres/done berarti sudah
  // diambil user. Hanya menyentuh baris yang belum pernah disentuh (updated_by
  // NULL) supaya koreksi manual admin tidak tertimpa.
  try {
    await db.run(
      `UPDATE dokumen_status
          SET status = 'TERGENERATE',
              updated_by = 'sistem',
              updated_at = datetime('now','localtime')
        WHERE jenis = 'pickup'
          AND status = 'BELUM'
          AND updated_by IS NULL
          AND lokasi_id IN (SELECT id FROM lokasi WHERE status_eviden IN ('progres','done'))`
    );
  } catch (e) { /* tabel lokasi belum siap — dilewati */ }

  // Dokumen yang sudah pernah digenerate lewat aplikasi tetap tercentang walau
  // barisnya baru dibuat — jejaknya ada di tabel aktivitas (hasil generate
  // sebelum fitur ini ada / sebelum server di-restart). Riwayat menang atas
  // status BELUM; koreksi manual NA tetap dihormati.
  try {
    const aliasAktivitas = Object.keys(DOK_DARI_AKTIVITAS)
      .map(k => `(dokumen_status.jenis = '${DOK_DARI_AKTIVITAS[k]}' AND a.jenis = '${k}')`)
      .join(' OR ');
    const cocokJenis = `a.jenis = dokumen_status.jenis${aliasAktivitas ? ' OR (' + aliasAktivitas + ')' : ''}`;
    await db.run(
      `UPDATE dokumen_status
          SET status = 'TERGENERATE',
              generated_at = COALESCE((SELECT MAX(a.timestamp) FROM aktivitas a
                                        WHERE a.lokasi_id = dokumen_status.lokasi_id
                                          AND (${cocokJenis})),
                                      datetime('now','localtime')),
              updated_by = 'sistem',
              updated_at = datetime('now','localtime')
        WHERE status = 'BELUM'
          AND jenis <> 'pickup'
          AND EXISTS (SELECT 1 FROM aktivitas a
                       WHERE a.lokasi_id = dokumen_status.lokasi_id
                         AND (${cocokJenis}))`
    );
  } catch (e) { console.warn('backfill dokumen dari aktivitas:', e.message); }

  return tambah;
}

// Tulis satu status dokumen. Cek dulu apakah barisnya sudah ada supaya tidak
// bergantung pada nilai "baris terubah" (lihat catatan di db-wrapper.js).
async function tandaiDokumen(lokasi_id, jenis, status, opts) {
  opts = opts || {};
  const ada = await db.get('SELECT id FROM dokumen_status WHERE lokasi_id = ? AND jenis = ?', [lokasi_id, jenis]);
  if (ada) {
    await db.run(
      `UPDATE dokumen_status
          SET status = ?,
              generated_at = CASE WHEN ? = 'TERGENERATE' THEN COALESCE(?, datetime('now','localtime')) ELSE NULL END,
              updated_by = ?,
              updated_at = datetime('now','localtime')
        WHERE lokasi_id = ? AND jenis = ?`,
      [status, status, opts.generated_at || null, opts.updated_by || null, lokasi_id, jenis]
    );
    return;
  }
  await db.run(
    `INSERT INTO dokumen_status (lokasi_id, jenis, status, generated_at, updated_by)
     VALUES (?, ?, ?, CASE WHEN ? = 'TERGENERATE' THEN COALESCE(?, datetime('now','localtime')) ELSE NULL END, ?)`,
    [lokasi_id, jenis, status, status, opts.generated_at || null, opts.updated_by || null]
  );
}

init();

// Middleware Proteksi Hak Akses
// IP klien — aman dari spoofing X-Forwarded-For (CWE-290 / advisory proxy-addr):
//   * Socket PEER publik  → klien menyambung langsung → pakai socket (tidak ada
//     yang bisa memalsukannya).
//   * Socket PEER privat/loopback (127.x, 10.x, 192.168.x, 172.16-31.x) → ada
//     proxy (nginx Hostinger / local) di depan → pakai baris TERAKHIR XFF,
//     yaitu alamat yang dijamin proxy; baris depan bisa dikarang klien.
// Express sendiri tidak mempercayai XFF (trust proxy = false) sehingga req.ip
// selalu alamat socket.
function ipKlien(req) {
  const peer = String(req.ip || (req.socket && req.socket.remoteAddress) || '').replace(/^::ffff:/, '');
  const privat = peer === '::1' || peer === 'localhost' || peer === '127.0.0.1' ||
    /^10\./.test(peer) || /^192\.168\./.test(peer) || /^172\.(1[6-9]|2\d|3[01])\./.test(peer);
  if (privat) {
    const xf = req.headers['x-forwarded-for'];
    if (xf) {
      const rantai = String(xf).split(',').map(s => s.trim()).filter(Boolean);
      if (rantai.length) return rantai[rantai.length - 1].replace(/^::ffff:/, '');
    }
  }
  return peer || '—';
}

function checkAuth(role) {
  return (req, res, next) => {
    if (!req.session || !req.session.user) {
      return res.status(401).json({ message: 'Sesi Anda telah berakhir. Silakan login kembali.' });
    }
    // Catat aktivitas terakhir (throttle 30 dtk) — penanda "online" di panel
    // /api/sesi-aktif. Menulis req.session.user menandai sesi berubah, sehingga
    // express-session menyimpannya + updated_at ikut diperbarui.
    const now = Date.now();
    if (!req.session.user.akhir || now - req.session.user.akhir > 30000) {
      req.session.user.akhir = now;
    }
    if (role && req.session.user.role !== role) {
      return res.status(403).json({ message: 'Hak akses ditolak: Peran Anda tidak diizinkan.' });
    }
    next();
  };
}

// Helper format angka untuk log
function fmtLog(n) { return Number(n || 0).toLocaleString('id-ID'); }

// ─────────────────────────────────────────────────────────────────────────────
// LOGIN TEMATIK — pembatasan akses per Tematik Project
// ─────────────────────────────────────────────────────────────────────────────
// Setiap akun lapangan (role 'user') menyimpan daftar tematik yang boleh ia
// buka pada kolom users.tematik (JSON array). Setelah login berlaku dua aturan:
//   1. Permintaan yang menunjuk ke SATU lokasi (/api/.../:lokasi_id, body
//      lokasi_id, berkas, dst) ditolak 403 bila lokasi itu di luar tematiknya.
//   2. Jawaban daftar (res.json) dibersihkan — baris ber-tematik lain dibuang,
//      termasuk baris Data Master (tematik_project).
// Akun admin SELALU bebas; akun tanpa tematik dianggap TIDAK punya akses.
// Penjaga dipasang otomatis ke semua route lewat lapisan paling bawah, jadi
// route/rfungsi lama tidak perlu diubah satu per satu.
// ─────────────────────────────────────────────────────────────────────────────

function kunciTematik(v) { return String(v === null || v === undefined ? '' : v).trim().toLowerCase(); }

// Baca kolom users.tematik → array nama tematik (menerima JSON array, "A,B", atau kosong)
function bacaTematikUser(row) {
  const raw = row && row.tematik;
  if (Array.isArray(raw)) return raw.map(x => String(x).trim()).filter(Boolean);
  if (typeof raw !== 'string' || !raw.trim()) return [];
  const s = raw.trim();
  if (s.charAt(0) === '[') {
    try {
      const a = JSON.parse(s);
      if (Array.isArray(a)) return a.map(x => String(x).trim()).filter(Boolean);
    } catch (e) { /* JSON rusak → ulangi dengan pemisah koma */ }
  }
  return s.split(',').map(x => String(x).trim()).filter(Boolean);
}

// Seluruh tematik yang benar-benar dipakai di sistem (satu-satunya sumber
// validasi saat admin mengaitkan akun ke tematik).
async function daftarTematikSistem() {
  try {
    const rows = await db.all("SELECT DISTINCT tematik AS t FROM lokasi WHERE COALESCE(tematik,'') <> ''");
    return Array.from(new Set(rows.map(r => String(r.t).trim()).filter(Boolean)))
      .sort((a, b) => a.localeCompare(b));
  } catch (e) { return []; }
}

// null  = bebas (admin / belum login → biarkan checkAuth yang menilai)
// Set   = daftar tematik yang boleh (boleh kosong = tidak boleh apa-apa)
function tematikDiizinkan(req) {
  const u = req.session && req.session.user;
  if (!u || u.role === 'admin') return null;
  const daftar = Array.isArray(u.tematik) ? u.tematik : [];
  return new Set(daftar.map(kunciTematik).filter(Boolean));
}

// Baris daftar yang punya kolom tematik/tematik_project → kunci perbandingan.
// Hanya nilai string/angka yang dianggap penanda baris — bila "tematik" berbentuk
// array/objek (mis. grup rekap: {tematik:[…]}) itu wadah, bukan baris, jadi
// objeknya tetap diproses rekursif agar isinya tersaring.
function nilaiTematikBaris(obj) {
  if (!obj || typeof obj !== 'object' || Array.isArray(obj)) return undefined;
  const t = obj.tematik;
  if (typeof t === 'string' || typeof t === 'number') {
    const s = String(t).trim();
    if (s) return kunciTematik(s);
  }
  const p = obj.tematik_project;
  if (typeof p === 'string' || typeof p === 'number') {
    const s = String(p).trim();
    if (s) return kunciTematik(godcTematikRowServer(s));
  }
  return undefined;
}

// Buang baris di luar tematik pengguna — berlaku rekursif pada seluruh payload
function saringJsonTematik(v, izin) {
  if (Array.isArray(v)) {
    const hasil = [];
    for (const item of v) {
      const t = nilaiTematikBaris(item);
      if (t !== undefined && !izin.has(t)) continue;   // baris tematik lain → buang
      hasil.push(saringJsonTematik(item, izin));
    }
    return hasil;
  }
  if (v && typeof v === 'object') {
    const hasil = {};
    for (const k of Object.keys(v)) hasil[k] = saringJsonTematik(v[k], izin);
    return hasil;
  }
  return v;
}

// Bungkus res.json agar jawaban selalu disaring dulu (lapisan kompresi di
// atas tetap bekerja karena kita memanggil res.json miliknya)
function pasangFilterTematik(res, izin) {
  const asli = res.json.bind(res);
  res.json = (body) => asli(saringJsonTematik(body, izin));
}

async function lokasiDariBerkas(id) {
  const n = Number(id);
  if (!Number.isInteger(n) || n <= 0) return null;
  const b = await db.get('SELECT lokasi_id FROM berkas WHERE id = ?', [n]);
  return b && b.lokasi_id ? Number(b.lokasi_id) : null;
}

// Kumpulkan id lokasi yang diminta request ini (parameter, body, query, turunan)
async function idLokasiDariRequest(req) {
  const ids = new Set();
  const angka = (v) => {
    const n = Number(v);
    if (Number.isInteger(n) && n > 0) ids.add(n);
  };
  const p = req.params || {};

  if (p.lokasi_id !== undefined) angka(p.lokasi_id);
  if (p.lokasiId  !== undefined) angka(p.lokasiId);

  // berkas/foto → lokasi pemiliknya
  for (const k of ['berkas_id', 'fotoId']) {
    if (p[k] !== undefined) {
      const lid = await lokasiDariBerkas(p[k]);
      if (lid) ids.add(lid);
    }
  }

  // param bernama "id" yang berarti lokasi: /api/lokasi/:id
  if (p.id !== undefined && /^\/api\/lokasi\//.test(req.path)) angka(p.id);

  // /api/input-khusus/:id → input_khusus.lokasi_id
  if (p.id !== undefined && /^\/api\/input-khusus\//.test(req.path)) {
    const n = Number(p.id);
    if (Number.isInteger(n) && n > 0) {
      const ik = await db.get('SELECT lokasi_id FROM input_khusus WHERE id = ?', [n]);
      if (ik && ik.lokasi_id) ids.add(Number(ik.lokasi_id));
    }
  }

  // body (JSON maupun form multipart) & query
  for (const sumber of [req.body, req.query]) {
    if (!sumber) continue;
    if (sumber.lokasi_id !== undefined) angka(sumber.lokasi_id);
    if (sumber.lokasiId  !== undefined) angka(sumber.lokasiId);
  }

  return ids;
}

// Penjaga yang ditanam di depan tiap handler route
async function batasAksesTematik(req, res, next) {
  let izin;
  try { izin = tematikDiizinkan(req); }
  catch (e) {
    console.error('[tematik] gagal membaca sesi:', e);
    return res.status(500).json({ message: 'Gagal memverifikasi hak akses Anda.' });
  }
  if (izin === null) return next();          // admin / belum login → jangan diutak-atik

  try {
    const ids = await idLokasiDariRequest(req);
    for (const id of ids) {
      const lokasi = await db.get('SELECT tematik FROM lokasi WHERE id = ?', [id]);
      if (!lokasi) continue;                 // id tak dikenal → biarkan handler menilai
      if (!izin.has(kunciTematik(lokasi.tematik))) {
        return res.status(403).json({
          message: 'Akses ditolak: lokasi "' + (lokasi.tematik || 'Tanpa Tematik') +
                   '" di luar tematik yang ditugaskan kepada Anda.'
        });
      }
    }
  } catch (e) {
    console.error('[tematik] gagal memeriksa lokasi:', e);
    return res.status(500).json({ message: 'Gagal memverifikasi hak akses Anda.' });
  }

  pasangFilterTematik(res, izin);
  next();
}

// ── Semua endpoint /api wajib login (kecuali login & logout) ─────────────────
// Sebelumnya sebagian endpoint baca terbuka tanpa sesi, sehingga pengguna yang
// logout tetap bisa melihat seluruh data lewat pintu belakang.
app.use('/api', (req, res, next) => {
  const p = req.path;                        // path relatif terhadap /api → "/login"
  if (p === '/login' || p === '/logout') return next();
  if (!req.session || !req.session.user) {
    return res.status(401).json({ message: 'Silakan login terlebih dahulu.' });
  }
  next();
});

// ── Pemasang otomatis: sisipkan batasAksesTematik tepat sebelum handler ─────
// Urutan tiap route jadi: middleware lama… → penjaga tematik → handler.
// Karena dipasang sesudah multer/body-parser, body form (lokasi_id) sudah terbaca.
(function pasangPenjagaTematik() {
  ['get', 'post', 'put', 'patch', 'delete', 'head'].forEach(function (metode) {
    const asli = app[metode].bind(app);
    app[metode] = function (path) {
      const handlers = [];
      for (let i = 1; i < arguments.length; i++) {
        const h = arguments[i];
        if (Array.isArray(h)) handlers.push.apply(handlers, h);
        else handlers.push(h);
      }
      const siap = handlers.filter(function (h) { return typeof h === 'function'; });
      if (!siap.length) return asli(path);
      const handlerAkhir = siap.pop();   // handler terminal selalu terakhir
      siap.push(batasAksesTematik, handlerAkhir);
      return asli.apply(app, [path].concat(siap));
    };
  });
})();


// ── KATEGORI PETA RUTE (KML/KMZ) ─────────────────────────────────────────────
const KATEGORI_PETA = '-PETA RUTE (KML/KMZ)';

// Kategori yang ditampilkan sebagai gambar halaman penuh (1 gambar per halaman)
const KATEGORI_KML_FULL = new Set(['KML (KML)', KATEGORI_PETA]);

// ── KML ASLI (untuk Google Earth) ─────────────────────────────────────────────
// Sejak perubahan ini sistem MENYIMPAN file .kml/.kmz asli hasil unggahan,
// sebangun dengan file render PNG-nya (nama sama: base + ekstensi). Berkas lama
// (sebelum penyimpanan asli) tidak punya pasangan → fungsi ini mengembalikan null.
function baseTanpaExt(nama) {
  return String(nama || '').replace(/\.[^.]+$/, '');
}

function cariKmlAsli(folder, base) {
  const kandidat = [path.join(folder, base + '.kml'), path.join(folder, base + '.kmz')];
  for (const p of kandidat) {
    if (fs.existsSync(p)) return p;
  }
  return null;
}

function escKml(s) {
  return String(s == null ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&apos;');
}

// Placemark KML dari hasil parseKml ({ name, type, points:[[lon,lat]] })
function pmKml(f) {
  const nm = '<name>' + escKml(f.name || 'Tanpa Nama') + '</name>';
  const pts = (f.points || []).filter(Boolean);
  if (!pts.length) return '';
  if (f.type === 'point' || pts.length === 1) {
    return '<Placemark>' + nm + '<styleUrl>#sTitik</styleUrl>'
      + '<Point><coordinates>' + pts[0][0] + ',' + pts[0][1] + '</coordinates></Point></Placemark>';
  }
  const koord = pts.map(p => p[0] + ',' + p[1]).join(' ');
  if (f.type === 'polygon') {
    return '<Placemark>' + nm + '<styleUrl>#sRute</styleUrl><Polygon><outerBoundaryIs><LinearRing>'
      + '<coordinates>' + koord + '</coordinates></LinearRing></outerBoundaryIs></Polygon></Placemark>';
  }
  return '<Placemark>' + nm + '<styleUrl>#sRute</styleUrl><LineString><coordinates>'
    + koord + '</coordinates></LineString></Placemark>';
}

// Daftar berkas wajib per tematik (dipakai /api/progres & cetak laporan).
const DATABASE_TEMATIK_WAJIB = {
  "MyRepublik": [
    "Myrep-EVIDEN PENARIKAN (MAX 6 FOTO)", "-Pemasangan Tiang (MAX 6 FOTO)",
    "-Pengecoran (6 foto)", "-Instalasi FDT( 6 foto)", "-Instalasi JC( 6 foto)",
    "-Istalasi Per FAT", "-Instalasi Pole",
    "-Eviden Redaman Spliter (foto ada 9 foto)", "-Eviden Redaman FAT (1 FAT ADA 8 FOTO)"
  ],
  "Jlm": [
    "-Eviden Marking Kabel 3 foto", "-Eviden Subdact 40/34 3 foto",
    "-Eviden Subdact 32/28 3 foto", "-Penarikan Kabel KU 3 foto",
    "-Pemasangan HH Portable 30*30 3 foto", "-Instalasi ODP 3 foto",
    "-Tali Tambang 3 Foto", "-Instalasi Roset 3 foto",
    "-Pemasangan J- Closure 3 foto", "-Penarikan Kabel DC 3 foto",
    "-Pembuatan HH 60*60 3 foto", "-Galian 4 foto", "-Boring 3 foto",
    "-Pemasangan Clean out 3 foto", "-Instalasi ODC 3 foto",
    "-Bobok Beton 3 foto", "-Eviden Redaman"
  ],
  "Relokasi": [
    "MASUKAN EVIDEN PEKERJAAN",
    "-EVIDEN PEMASANGAN ODP",
    "-EVIDENT HASIL UKUR OPM",
    "INPUT REKAP OPM",
    "-EVIDEN COLUSER",
    "-EVIDEN DISMANTLING",
    "-LAMPIRAN DATA PENGUKURAN OPM",
    "-MANCORE",
    "-LAMPIRAN HASIL UJI KEDALAMAN GALIAN",
    "-LAMPIRAN CO"
  ],
  "M-tel": [
    "-EVIDENT HASIL UKUR OPM", "-LAMPIRAN DATA PENGUKURAN OPM",
    "-MANCORE", "-GROUNDING"
  ],
  "HEM": [
    "MANCORE", "ABD (PDF/DWG/DXF)", "KML (KML)", "BOQ AKTUAL",
    "EVIDEN BON MATERIAL", "INPUT RFC", "EVIDEN PENARIKAN KABEL",
    "EVIDEN PEMASANGAN ACC", "EVIDEN PEMASANGAN TIANG",
    "EVIDEN INSTALASI ODP/OTB", "EVIDEN JUMPER",
    "EVIDEN REDAMAN ODP/OTB", "EVIDEN HDPE",
    "EVIDEN OTDR", "EVIDEN GALIAN", "EVIDEN HH AKSES", "EVIDEN HH ODP"
  ],
  "MICRODEMAND": [
    "ABD (PDF/DWG/DXF)", "KML (KML)", "BOQ AKTUAL",
    "EVIDEN BON MATERIAL", "INPUT RFC", "EVIDEN PENARIKAN KABEL",
    "EVIDEN PEMASANGAN ACC", "EVIDEN PEMASANGAN TIANG",
    "EVIDEN INSTALASI ODP/OTB", "EVIDEN JUMPER",
    "EVIDEN REDAMAN ODP/OTB", "EVIDEN HDPE",
    "EVIDEN OTDR", "EVIDEN GALIAN", "EVIDEN HH AKSES", "EVIDEN HH ODP"
  ]
};

// Tematik Node B: daftar HEM + tiga inputan khusus Node B (LACT).
DATABASE_TEMATIK_WAJIB["Node B"] = [
  ...DATABASE_TEMATIK_WAJIB["HEM"],
  "BOQ BARANG TIBA",
  "FORM PENGUKURAN OPM",
  "OTDR (.SOR)"
];

// Nama file template .docx per tematik (pakai pembanding bila tematiknya tak punya template sendiri).
const TEMPLATE_ALIAS_TEMATIK = { "Node B": "HEM" };
function namaTemplateDocx(tematik) {
  // Node B (LACT) punya master template sendiri: TAMPLATE_NODE_B.docx
  if (tematik === "Node B") return "TAMPLATE_NODE_B.docx";
  return `template_${TEMPLATE_ALIAS_TEMATIK[tematik] || tematik}.docx`;
}

// Konversi gambar PNG menjadi PDF satu halaman (pdf-lib).
async function pngToPdf(pngPath) {
  const pngBytes = fs.readFileSync(pngPath);
  const doc = await PDFDocument.create();
  const img = await doc.embedPng(pngBytes);
  const page = doc.addPage([img.width, img.height]);
  page.drawImage(img, { x: 0, y: 0, width: img.width, height: img.height });
  return doc.save();
}

// ═══════════════════════════════════════════════════════════════════════════
// OUTPUT PDF — konversi .docx hasil generate menjadi PDF memakai MS Word
// (COM/Word.Application). Dipakai bila request membawa ?format=pdf.
// ═══════════════════════════════════════════════════════════════════════════
let antreanKonversiPdf = Promise.resolve();

function jalankanPs(perintah, timeoutMs) {
  return new Promise((resolve, reject) => {
    execFile('powershell.exe',
      ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-Command', perintah],
      { timeout: timeoutMs || 120000, maxBuffer: 16 * 1024 * 1024, windowsHide: true },
      (err, stdout, stderr) => {
        if (err) reject(new Error(String(stderr || err.message || '').trim().slice(0, 400) || 'Proses konversi gagal.'));
        else resolve(stdout);
      });
  });
}

async function konversiDocxKePdf(buf) {
  const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'dokpdf-'));
  const docxPath = path.join(tmp, 'dokumen.docx');
  const pdfPath = path.join(tmp, 'dokumen.pdf');
  try {
    fs.writeFileSync(docxPath, buf);
    const q = p => "'" + String(p).replace(/'/g, "''") + "'";
    const perintah = [
      "$ErrorActionPreference='Stop'",
      "$w=New-Object -ComObject Word.Application",
      "$w.Visible=$false",
      "$w.DisplayAlerts=0",
      "$wpid=0",
      "try { $s='[DllImport(\"user32.dll\")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint p);'; Add-Type -MemberDefinition $s -Namespace W32 -Name Np | Out-Null; [W32.Np]::GetWindowThreadProcessId([IntPtr]$w.Hwnd,[ref]$wpid) | Out-Null } catch {}",
      "try {",
      "  $d=$w.Documents.Open(" + q(docxPath) + ",$false,$true,$false)",
      "  $d.ExportAsFixedFormat(" + q(pdfPath) + ",17)",
      "  $d.Close($false)",
      "} finally {",
      "  try { $w.Quit() } catch {}",
      "  try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($w) } catch {}",
      "  if ($wpid -gt 0) { $pp=Get-Process -Id $wpid -ErrorAction SilentlyContinue; if ($pp) { Stop-Process -Id $wpid -Force -ErrorAction SilentlyContinue } }",
      "}"
    ].join('; ');
    await jalankanPs(perintah, 120000);
    if (!fs.existsSync(pdfPath) || !fs.statSync(pdfPath).size) {
      throw new Error('MS Word tidak menghasilkan file PDF.');
    }
    return fs.readFileSync(pdfPath);
  } finally {
    try { fs.rmSync(tmp, { recursive: true, force: true }); } catch (e) {}
  }
}

// Kirim dokumen hasil generate: .docx bila biasa, .pdf bila ?format=pdf.
function kirimDokumen(req, res, buf, namaDocx) {
  const mauPdf = String((req.query && (req.query.format || req.query.output)) || '').toLowerCase() === 'pdf';
  const kirim = (isi, tipe, nama) => {
    res.setHeader('Content-Type', tipe);
    res.setHeader('Content-Disposition', 'attachment; filename="' + nama + '"');
    res.send(isi);
  };
  if (!mauPdf) {
    return kirim(buf, 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', namaDocx);
  }
  // Satu-satu, agar instance MS Word tidak saling bentrok
  const tugas = antreanKonversiPdf.then(() => konversiDocxKePdf(buf));
  antreanKonversiPdf = tugas.catch(() => {});
  return tugas.then(pdf => {
    kirim(pdf, 'application/pdf', String(namaDocx).replace(/\.docx$/i, '.pdf'));
  }).catch(err => {
    console.error('Gagal konversi PDF:', err.message);
    if (!res.headersSent) {
      res.status(500).json({
        message: 'Gagal membuat PDF (dikonversi lewat MS Word). Pastikan MS Word terpasang di server. Detail: ' + err.message
      });
    }
  });
}

// Cari berkas peta rute terbaru milik sebuah lokasi.
async function cariPetaRute(lokasi_id) {
  const rows = await db.all(
    'SELECT file_name, sub_kategori FROM berkas WHERE lokasi_id = ? AND kategori = ? ORDER BY id DESC LIMIT 1',
    [lokasi_id, KATEGORI_PETA]
  );
  if (!rows.length) return null;
  const lokasi = await db.get('SELECT tematik, nama_lokasi FROM lokasi WHERE id = ?', [lokasi_id]);
  if (!lokasi) return null;
  const ntc = (lokasi.tematik || 'Umum').replace(/[^a-zA-Z0-9_-]/g, '_');
  const nlc = lokasi.nama_lokasi.replace(/[^a-zA-Z0-9_-]/g, '_');
  const kc  = KATEGORI_PETA.replace(/[^a-zA-Z0-9_-]/g, '_');
  const fp  = path.join(__dirname, 'uploads', ntc, nlc, kc, rows[0].file_name);
  return fs.existsSync(fp)
    ? { path: fp, nama: rows[0].sub_kategori || 'Peta Rute' }
    : null;
}

// Deteksi format gambar dari magic bytes (awalan file).
function deteksiEkstensiMedia(buf) {
  if (!buf || buf.length < 4) return null;
  const b = buf[0], c = buf[1], d = buf[2], e = buf[3];
  if (buf.length >= 8 && b === 0x89 && c === 0x50 && d === 0x4E && e === 0x47) return 'png';
  if (buf.length >= 3 && b === 0xFF && c === 0xD8 && d === 0xFF) return 'jpg';
  if (buf.length >= 6 && b === 0x47 && c === 0x49 && d === 0x46 && e === 0x38) return 'gif';
  if (buf.length >= 2 && b === 0x42 && c === 0x4D) return 'bmp';
  if (buf.length >= 12 && b === 0x52 && c === 0x49 && d === 0x46 && e === 0x46 &&
      buf[8] === 0x57 && buf[9] === 0x45 && buf[10] === 0x42 && buf[11] === 0x50) return 'webp';
  if ((b === 0x49 && c === 0x49 && d === 0x2A && e === 0x00) ||
      (b === 0x4D && c === 0x4D && d === 0x00 && e === 0x2A)) return 'tif';
  return null;
}

// Perbaiki part media hasil inject: ekstensi file word/media/* WAJIB sesuai isi data.
// ImageModule menamai semua gambar "image_generated_N.png", padahal isinya bisa JPG/GIF
// (foto upload, blank.png). Word menolak membuka DOCX bila isi part media tidak cocok
// dengan ekstensi/content-type — inilah penyebab "Word experienced an error trying to open the file".
function perbaikiMediaParts(zip) {
  const renameDisini = {};
  Object.keys(zip.files).forEach((n) => {
    if (zip.files[n].dir || !/^word\/media\//.test(n)) return;
    const f = zip.files[n];
    let raw;
    try { raw = f.asNodeBuffer ? f.asNodeBuffer() : Buffer.from(f.asBinary(), 'binary'); }
    catch (err) { raw = null; }
    if (!raw || !raw.length) return;
    const kanonik = deteksiEkstensiMedia(raw);
    if (!kanonik) return;
    const cur = (n.split('.').pop() || '').toLowerCase();
    const curKanonik = cur === 'jpeg' || cur === 'tiff' ? cur.replace('jpeg', 'jpg').replace('tiff', 'tif') : cur;
    if (kanonik === curKanonik) return;
    const namaBaru = n.slice(0, n.lastIndexOf('.') + 1) + kanonik;
    if (zip.files[namaBaru]) return; // jangan timpa part lain
    renameDisini[n] = namaBaru;
  });

  if (Object.keys(renameDisini).length === 0) return;

  Object.keys(renameDisini).forEach((oldN) => {
    const src = zip.files[oldN];
    const data = src.asNodeBuffer ? src.asNodeBuffer() : Buffer.from(src.asBinary(), 'binary');
    zip.file(renameDisini[oldN], data);
    zip.file(oldN, Buffer.from(''));
    zip.remove(oldN);
  });

  // Perbarui Target/relasi di semua .rels (document.xml.rels, header*.xml.rels, dsb.)
  const daftarRels = Object.keys(zip.files).filter(k => /\.rels$/.test(k) && !zip.files[k].dir);
  daftarRels.forEach((r) => {
    let txt = zip.files[r].asText();
    let berubah = false;
    Object.keys(renameDisini).forEach((oldN) => {
      // oldN = "word/media/image_generated_1.png" → filename = "image_generated_1.png"
      // Rels pakai path relatif: Target="media/image_generated_1.png"
      const fileName = oldN.split('/').pop();
      const newFileName = renameDisini[oldN].split('/').pop();
      // Cobabagai variasi path: "media/file.ext", "file.ext", atau full path
      const targets = [
        'Target="media/' + fileName + '"',
        'Target="' + fileName + '"',
      ];
      for (const lama of targets) {
        if (!txt.includes(lama)) continue;
        txt = txt.split(lama).join('Target="media/' + newFileName + '"');
        berubah = true;
      }
    });
    if (berubah) zip.file(r, txt);
  });

  // Daftarkan ekstensi baru di [Content_Types].xml bila belum ada
  const ctypes = zip.file('[Content_Types].xml');
  if (ctypes) {
    const petaType = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp', webp: 'image/webp', tif: 'image/tiff', tiff: 'image/tiff' };
    let ct = ctypes.asText();
    const dibutuhkan = new Set(Object.keys(renameDisini).map(n => renameDisini[n].split('.').pop().toLowerCase()));
    let berubah = false;
    dibutuhkan.forEach((ext) => {
      const re = new RegExp('<Default\\s+Extension="' + ext + '"', 'i');
      if (re.test(ct)) return;
      const ctype = petaType[ext] || 'application/octet-stream';
      ct = ct.replace('</Types>', '<Default Extension="' + ext + '" ContentType="' + ctype + '"/></Types>');
      berubah = true;
    });
    if (berubah) zip.file('[Content_Types].xml', ct);
  }
}

// Word 2021 (dan Office versi baru) menolak membuka DOCX bila ada ID gambar yang duplikat.
// docxtemplater-image-module-free selalu menulis wp:docPr id="2" dan pic:cNvPr id="0"
// pada SETIAP gambar yang disisipkan → bila ada 100+ foto, id=2 terulang 100x.
// WPS menerimanya, tapi Word menandai "unreadable content". Normalisasi di bawah
// menjamin semua ID gambar UNIK di seluruh part (document + header + footer).
function normalisasiDocPrIds(zip) {
  const parts = Object.keys(zip.files).filter(n => !zip.files[n].dir &&
    /^word\/(document|header\d*|footer\d*)\.xml$/i.test(n));
  if (!parts.includes('word/document.xml')) return;

  let nextId = 0;
  const used = new Set();

  parts.forEach((p) => {
    let xml = zip.files[p].asText();
    const re = /<(?:wp:docPr|pic:cNvPr)\s+id="(\d+)"/g;
    let m;
    let ditemukan = false;
    const daftar = [];
    while ((m = re.exec(xml))) daftar.push({ raw: m[0], id: m[1] });
    if (!daftar.length) return;
    daftar.forEach(({ raw, id }) => {
      // Pertahankan id lama bila masih unik; jika tidak, beri id baru berurutan.
      if (!used.has(Number(id))) {
        used.add(Number(id));
        return;
      }
      let baru = nextId;
      while (used.has(baru)) baru++;
      used.add(baru);
      nextId = baru + 1;
      xml = xml.replace(raw, raw.replace('id="' + id + '"', 'id="' + baru + '"'));
      ditemukan = true;
    });
    if (ditemukan) zip.file(p, xml);
  });
}

// Word 2021 (dan Office modern) menolak membuka DOCX bila skema OOXML-nya dilanggar.
// Pelanggaran yang paling sering lolos WPS tapi ditolak Word: elemen blok (w:p, w:tbl,
// w:r, w:tr, w:tc, w:drawing) yang tersarang DI DALAM elemen properti (w:pPr, w:rPr,
// w:tblPr, dsb). Fungsi ini memindahkan blok tersebut keluar menjadi sibling, persis
// seperti yang diperbaiki pada templateRelokasi*.docx (section {#eep_list}).
function perbaikiPPrInvalid(zip) {
  const p = 'word/document.xml';
  if (!zip.files[p]) return;
  const BLOCK_BY_PROP = {
    pPr: ['p', 'tbl', 'tr', 'tc', 'r', 'drawing', 'pict', 'sdt'],
    rPr: ['p', 'tbl', 'tr', 'tc', 'drawing', 'pict'],
    tblPr: ['p', 'tbl', 'tr', 'tc'],
    tcPr: ['p', 'tr', 'tbl', 'tc', 'r', 'drawing'],
    trPr: ['p', 'tr', 'tc', 'tbl'],
    tblGrid: ['p', 'tbl', 'tr', 'tc'],
  };
  let dom;
  const xml = zip.files[p].asText();
  try { dom = new DocPrParser().parseFromString(xml, 'text/xml'); }
  catch (e) { return; }
  const doc = dom.documentElement;
  let fixed = 0;
  // kumpulkan semua node dulu (jangan mutasi saat traversal)
  const nodes = [];
  const visit = (n) => {
    if (!n || n.nodeType !== 1) return;
    nodes.push(n);
    for (let i = 0; i < n.childNodes.length; i++) visit(n.childNodes[i]);
  };
  visit(doc);
  nodes.forEach((propEl) => {
    if (!BLOCK_BY_PROP[propEl.localName]) return;
    const badNames = BLOCK_BY_PROP[propEl.localName];
    const blocks = [];
    for (let i = 0; i < propEl.childNodes.length; i++) {
      const c = propEl.childNodes[i];
      if (c.nodeType === 1 && badNames.includes(c.localName)) blocks.push(c);
    }
    if (!blocks.length) return;
    const home = propEl.parentNode;
    if (!home) return;
    blocks.forEach((b) => { home.parentNode.insertBefore(b, home); fixed++; });
    const masihAdaElm = Array.prototype.some.call(propEl.childNodes, c => c.nodeType === 1);
    if (!masihAdaElm) propEl.parentNode.removeChild(propEl);
  });
  if (fixed) {
    zip.file(p, new DocPrSerializer().serializeToString(doc));
    console.log('[DOCX] perbaikiPPrInvalid: ' + fixed + ' blok dipindahkan keluar dari properti.');
  }
}

// Sisipkan halaman lampiran peta rute ke akhir dokumen DOCX yang sudah jadi.
// Bekerja langsung pada XML word/document.xml sehingga template tidak perlu diubah.
async function lampirkanPetaKeDocx(docxBuf, lokasi_id, opts) {
  // opts.sisipPeta = false → hanya normalisasi media, JANGAN tambahkan lampiran peta
  // (dipakai bila template sudah menampilkan peta lewat {#kml_list} agar tidak dobel).
  const sisipPeta = !opts || opts.sisipPeta !== false;
  let zip;
  try {
    zip = new PizZip(docxBuf);
  } catch (e) {
    console.error('Gagal buka DOCX zip:', e.message);
    return docxBuf;
  }

  // Selalu normalisasi media (fix ekstensi PNG/JPG/GIF) — ini WAJIB untuk semua kasus.
  const mediaCount = Object.keys(zip.files).filter(n => /^word\/media\//.test(n)).length;
  console.log('[DOCX] lampirkanPetaKeDocx dipanggil, media parts:', mediaCount);
  try { perbaikiMediaParts(zip); } catch (e) { console.error('perbaikiMediaParts:', e.message); }
  try { normalisasiDocPrIds(zip); } catch (e) { console.error('normalisasiDocPrIds:', e.message); }
  try { perbaikiPPrInvalid(zip); } catch (e) { console.error('perbaikiPPrInvalid:', e.message); }

  // Jika tidak ada kolom koneksi atau tidak ada peta, generate langsung.
  let peta;
  try {
    peta = await cariPetaRute(lokasi_id);
  } catch (e) {
    console.error('Gagal cari peta:', e.message);
    peta = null;
  }

  if (!peta || !sisipPeta) {
    return zip.generate({ type: 'nodebuffer', compression: 'DEFLATE' });
  }

  try {
    if (!zip.file('word/document.xml') || !zip.file('word/_rels/document.xml.rels')) {
      return zip.generate({ type: 'nodebuffer', compression: 'DEFLATE' });
    }

    const docXml = zip.file('word/document.xml').asText();
    if (!docXml.includes('</w:body>')) {
      return zip.generate({ type: 'nodebuffer', compression: 'DEFLATE' });
    }

    const pngBytes = fs.readFileSync(peta.path);
    const dim = sizeOf(pngBytes) || { width: 1000, height: 700 };

    let wEmu = Math.round(5.9 * 914400);
    let hEmu = Math.round((dim.height / dim.width) * wEmu);
    const maxHEmu = 9 * 914400;
    if (hEmu > maxHEmu) { hEmu = maxHEmu; wEmu = Math.round(hEmu * (dim.width / dim.height)); }

    const mediaName = 'peta_' + Date.now() + '.png';
    zip.file('word/media/' + mediaName, pngBytes);

    if (zip.file('[Content_Types].xml')) {
      let ctypes = zip.file('[Content_Types].xml').asText();
      if (!/<Default\s+Extension="png"\s+ContentType="image\/png"\/>/.test(ctypes)) {
        ctypes = ctypes.replace('</Types>', '<Default Extension="png" ContentType="image/png"/></Types>');
        zip.file('[Content_Types].xml', ctypes);
      }
    }

    const rels = zip.file('word/_rels/document.xml.rels').asText();
    let maxRid = 0;
    const reRid = /rId(\d+)/g;
    let m;
    while ((m = reRid.exec(rels))) maxRid = Math.max(maxRid, parseInt(m[1], 10));
    const rid = 'rId' + (maxRid + 1);
    const relEntry = '<Relationship Id="' + rid + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/' + mediaName + '"/>';
    zip.file('word/_rels/document.xml.rels', rels.replace('</Relationships>', relEntry + '</Relationships>'));

    const docPrId = maxRid + 100;
    const drawing =
      '<w:drawing xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
      '<wp:inline distT="0" distB="0" distL="0" distR="0">' +
      '<wp:extent cx="' + wEmu + '" cy="' + hEmu + '"/>' +
      '<wp:effectExtent l="0" t="0" r="0" b="0"/>' +
      '<wp:docPr id="' + docPrId + '" name="Peta Rute"/>' +
      '<wp:cNvGraphicFramePr/>' +
      '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
      '<pic:pic><pic:nvPicPr><pic:cNvPr id="' + docPrId + '" name="Peta Rute"/><pic:cNvPicPr/></pic:nvPicPr>' +
      '<pic:blipFill><a:blip r:embed="' + rid + '"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>' +
      '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="' + wEmu + '" cy="' + hEmu + '"/></a:xfrm>' +
      '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>' +
      '</a:graphicData></a:graphic></wp:inline></w:drawing>';

    const caption = String(peta.nama || 'Peta Rute').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
    const insert =
      '<w:p><w:pPr><w:pageBreakBefore/></w:pPr></w:p>' +
      '<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>Lampiran: Peta Rute &#8212; ' + caption + '</w:t></w:r></w:p>' +
      '<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r>' + drawing + '</w:r></w:p>';

    zip.file('word/document.xml', docXml.replace('</w:body>', insert + '</w:body>'));
  } catch (e) {
    console.error('Gagal sisipkan peta (dilewati):', e.message);
  }

  return zip.generate({ type: 'nodebuffer', compression: 'DEFLATE' });
}

const storage = multer.memoryStorage();
const upload  = multer({ storage, limits: { fileSize: 15 * 1024 * 1024 } });

// ── Pencatat aktivitas dokumen (untuk widget "Aktivitas Terbaru" di Dashboard) ──
async function catatAktivitas(jenis, lokasi_id, status = 'Selesai') {
  try {
    const lok = await db.get('SELECT nama_lokasi, pid FROM lokasi WHERE id = ?', [lokasi_id]);
    await db.run(
      'INSERT INTO aktivitas (jenis, lokasi_id, lokasi_nama, pid, status) VALUES (?,?,?,?,?)',
      [jenis, lokasi_id || null, lok ? lok.nama_lokasi : null, lok ? lok.pid : null, status]
    );
    // Generate berhasil -> status dokumen di tabel dokumen_status ikut TERGENERATE
    // (kolom Progres & dashboard per dokumen membaca tabel ini).
    const jenisDok = DOK_DARI_AKTIVITAS[jenis] || jenis;
    if (lokasi_id && DOK_STATUS_JENIS.indexOf(jenisDok) >= 0) {
      await tandaiDokumen(lokasi_id, jenisDok, 'TERGENERATE', { updated_by: 'sistem' });
    }
  } catch (e) {
    console.warn('catatAktivitas:', e.message);
  }
}

// ── Auto-set status_eviden ke 'done' jika 100% lengkap ──
async function autoSetDoneIfComplete(lokasi_id) {
  try {
    const lokasi = await db.get('SELECT tematik FROM lokasi WHERE id = ?', [lokasi_id]);
    if (!lokasi) return;
    const daftarWajib = DATABASE_TEMATIK_WAJIB[lokasi.tematik] || [];
    if (daftarWajib.length === 0) return;
    const berkasRows = await db.all('SELECT kategori FROM berkas WHERE lokasi_id = ?', [lokasi_id]);
    const inputRows = await db.all('SELECT kategori FROM input_khusus WHERE lokasi_id = ?', [lokasi_id]);
    const boqRows = await db.get('SELECT COUNT(*) AS n FROM boq_aktual WHERE lokasi_id = ?', [lokasi_id]);
    const btRows  = await db.get('SELECT COUNT(*) AS n FROM barang_tiba WHERE lokasi_id = ?', [lokasi_id]);
    const kategoriAda = [...berkasRows.map(b => b.kategori), ...inputRows.map(b => b.kategori)];
    if (boqRows && boqRows.n > 0) kategoriAda.push('BOQ AKTUAL');
    if (btRows && btRows.n > 0) kategoriAda.push('BOQ BARANG TIBA');
    const setWajib = new Set(daftarWajib);
    const sudah = kategoriAda.filter(k => setWajib.has(k)).length;
    const progres = Math.round((sudah / daftarWajib.length) * 100);
    if (progres >= 100) {
      await db.run("UPDATE lokasi SET status_eviden = 'done' WHERE id = ? AND (status_eviden IS NULL OR status_eviden != 'done')", [lokasi_id]);
    }
  } catch (e) {
    console.error('autoSetDoneIfComplete:', e.message);
  }
}

// [API] GET — Aktivitas/generate dokumen terbaru (dipakai widget Dashboard)
app.get('/api/aktivitas', checkAuth('admin'), async (req, res) => {
  try {
    const limit = Math.min(parseInt(req.query.limit, 10) || 12, 50);
    const rows  = await db.all(
      'SELECT id, jenis, lokasi_id, lokasi_nama, pid, status, timestamp FROM aktivitas ORDER BY id DESC LIMIT ?',
      [limit]
    );
    res.json(rows || []);
  } catch (e) {
    console.error('GET /api/aktivitas:', e.message);
    res.status(500).json({ message: 'Gagal memuat aktivitas.' });
  }
});

// ─────────────────────────────────────────────────────────────────────────────
// [API] SESI — peran pengguna yang sedang login (dipakai admin.html untuk
// mengunci koreksi manual status dokumen ke admin saja).
// ─────────────────────────────────────────────────────────────────────────────
app.get('/api/sesi', checkAuth(), (req, res) => {
  const u = req.session.user;
  res.json({
    ok: true,
    username: u.username,
    role: u.role,
    tematik: u.role === 'admin' ? [] : (Array.isArray(u.tematik) ? u.tematik : []),
    master: !!u.master,
    pengguna_unlock: !!(u.master || u.penggunaUnlock)
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// [API] STATUS DOKUMEN PER LOP — dipakai tabel Detail LOP & dashboard dokumen
//   GET  /api/dokumen-status        -> daftar LOP + status 6 dokumen
//   PUT  /api/dokumen-status        -> koreksi manual (admin)
//   PATCH/api/dokumen-status        -> koreksi manual (admin)
// ─────────────────────────────────────────────────────────────────────────────
function dokumenStatusDari(lops, peta) {
  return lops.map(l => {
    const d = peta.get(l.lokasi_id) || {};
    const dok = {}, gen_at = {};
    DOK_STATUS_JENIS.forEach(j => {
      dok[j] = d[j] ? d[j].status : 'BELUM';
      gen_at[j] = d[j] ? d[j].generated_at : null;
    });
    return Object.assign({}, l, { dok: dok, gen_at: gen_at });
  });
}

app.get('/api/dokumen-status', checkAuth(), async (req, res) => {
  try {
    await seedDokumenStatus();
    const lops = await ambilLopTrigger();
    const peta = await petaDokumen();
    res.json({
      ok: true,
      jenis: DOK_STATUS_JENIS,
      labels: DOK_STATUS_LABEL,
      rows: dokumenStatusDari(lops, peta)
    });
  } catch (e) {
    console.error('GET /api/dokumen-status:', e.message);
    res.status(500).json({ ok: false, message: 'Gagal memuat status dokumen.', detail: e.message });
  }
});

async function ubahStatusDokumen(req, res) {
  try {
    const b = req.body || {};
    const lokasiId = Number(b.lokasi_id);
    const jenis = String(b.jenis || '');
    const status = String(b.status || '').toUpperCase();
    if (!lokasiId) return res.status(400).json({ ok: false, message: 'lokasi_id wajib diisi.' });
    if (DOK_STATUS_JENIS.indexOf(jenis) < 0) {
      return res.status(400).json({ ok: false, message: 'Jenis dokumen tidak dikenal.' });
    }
    if (DOK_STATUS_VALUES.indexOf(status) < 0) {
      return res.status(400).json({ ok: false, message: 'Status harus BELUM, TERGENERATE, atau NA.' });
    }
    const lok = await db.get('SELECT id FROM lokasi WHERE id = ?', [lokasiId]);
    if (!lok) return res.status(404).json({ ok: false, message: 'Lokasi tidak ditemukan.' });
    await tandaiDokumen(lokasiId, jenis, status, { updated_by: req.session.user.username });
    res.json({ ok: true, lokasi_id: lokasiId, jenis: jenis, status: status, label: DOK_STATUS_LABEL[jenis] });
  } catch (e) {
    console.error('PUT /api/dokumen-status:', e.message);
    res.status(500).json({ ok: false, message: 'Gagal menyimpan status dokumen.', detail: e.message });
  }
}
app.put('/api/dokumen-status', checkAuth('admin'), ubahStatusDokumen);
app.patch('/api/dokumen-status', checkAuth('admin'), ubahStatusDokumen);

// Daftar lokasi pada satu jenis dokumen yang sudah TERGENERATE (untuk dashboard).
// Filter: ?branch= &tematik= &sp= (sama seperti halaman rekap).
app.get('/api/dokumen/:jenis/lokasi', checkAuth(), async (req, res) => {
  try {
    const jenis = String(req.params.jenis || '');
    if (DOK_STATUS_JENIS.indexOf(jenis) < 0) {
      return res.status(400).json({ ok: false, message: 'Jenis dokumen tidak dikenal.' });
    }
    await seedDokumenStatus();
    const lops = await ambilLopTrigger();
    const peta = await petaDokumen();

    const reqBranch = godcSync.normBranch(req.query.branch || '');
    const reqBranchKey = reqBranch ? godcSync.normBranchKey(reqBranch) : '';
    const reqTematik = String(req.query.tematik || '').trim();
    const reqSp = String(req.query.sp || '').trim();

    const rows = dokumenStatusDari(lops, peta)
      .filter(l => l.dok[jenis] === 'TERGENERATE')
      .filter(l => !reqBranchKey || godcSync.normBranchKey(l.branch) === reqBranchKey)
      .filter(l => !reqTematik || l.tematik === reqTematik)
      .filter(l => !reqSp || l.sp === reqSp)
      .map(l => ({
        lokasi_id: l.lokasi_id,
        nama: l.nama,
        sp: l.sp,
        drm: l.drm,
        po: l.po,
        branch: l.branch,
        tematik: l.tematik,
        generated_at: l.gen_at[jenis]
      }))
      .sort((a, b) => String(a.nama || '').localeCompare(String(b.nama || '')));

    const total = rows.reduce((s, r) => s + (Number(r.drm) || 0), 0);
    const uniq = (arr) => Array.from(new Set(arr)).sort((a, b) => String(a).localeCompare(String(b)));
    const seluruh = dokumenStatusDari(lops, peta);

    res.json({
      ok: true,
      jenis: jenis,
      label: DOK_STATUS_LABEL[jenis],
      filter: { branch: reqBranch || '', tematik: reqTematik, sp: reqSp },
      jumlah: rows.length,
      total_drm: total,
      branches: uniq(seluruh.map(l => l.branch)).filter(Boolean),
      tematiks: uniq(seluruh.map(l => l.tematik)).filter(Boolean),
      sps: uniq(seluruh.map(l => l.sp)).filter(Boolean),
      rows: rows
    });
  } catch (e) {
    console.error('GET /api/dokumen/:jenis/lokasi:', e.message);
    res.status(500).json({ ok: false, message: 'Gagal memuat daftar dokumen.', detail: e.message });
  }
});

// ─────────────────────────────────────────────────────────────────────────────
// [API] Lokasi
// ─────────────────────────────────────────────────────────────────────────────
app.get('/api/lokasi', async (req, res) => {
  try {
    const rows = await db.all('SELECT * FROM lokasi');
    // ?trigger=1 → hanya lokasi yang punya baris butuh eviden di Data Master
    // Proyek trigger/butuh eviden (AF = 03. MATDEV, 04. INSTALASI,
    // FINISH INSTALASI / OA, GO LIVE, 07. TC, 08. UT, atau 09. REKON).
    // Dipakai form Upload Eviden Lapangan; tanpa parameter tetap seluruh
    // lokasi agar admin aman.
    if (req.query.trigger !== '1') return res.json(rows);
    const trg = await db.all("SELECT nama_lop FROM godc_project WHERE butuh_eviden = 'YA'");
    const set = new Set(trg.map(r => godcSync.normName(r.nama_lop)));
    res.json(rows.filter(r => set.has(godcSync.normName(r.nama_lokasi))));
  } catch (e) {
    console.error(e);
    res.status(500).json({ message: 'Gagal memproses permintaan data lokasi.' });
  }
});

app.post('/api/lokasi', checkAuth('admin'), async (req, res) => {
  try {
    const { nama_lokasi, pid, tematik, no_sp } = req.body;
    if (!nama_lokasi || !tematik) {
      return res.status(400).json({ message: 'Nama lokasi dan Tematik Project wajib diisi!' });
    }
    await db.run(
      'INSERT INTO lokasi (pid, nama_lokasi, tematik, no_sp) VALUES (?, ?, ?, ?)',
      [pid || null, nama_lokasi, tematik, no_sp || null]
    );
    res.json({ success: true });
  } catch (e) {
    console.error(e);
    res.status(400).json({ message: 'Gagal menyimpan. Nama lokasi sudah terdaftar atau format salah.' });
  }
});

app.delete('/api/lokasi/:id', checkAuth('admin'), async (req, res) => {
  try {
    const lokasiId = req.params.id;
    // Dalam satu transaksi: bila salah satu DELETE gagal, semua batalkan —
    // jangan sampai baris anak (berkas dll) terhapus tetapi lokasi tetap hidup.
    const result = await db.batch(async () => {
      await db.run('DELETE FROM berkas WHERE lokasi_id = ?', [lokasiId]);
      await db.run('DELETE FROM laporan_rekon WHERE lokasi_id = ?', [lokasiId]);
      await db.run('DELETE FROM rab           WHERE lokasi_id = ?', [lokasiId]);
      await db.run('DELETE FROM rab_summary   WHERE lokasi_id = ?', [lokasiId]);
      await db.run('DELETE FROM input_khusus  WHERE lokasi_id = ?', [lokasiId]);
      return await db.run('DELETE FROM lokasi WHERE id = ?', [lokasiId]);
    });
    if (result.changes === 0) return res.status(404).json({ message: 'Data lokasi tidak ditemukan.' });
    res.json({ success: true, message: 'Lokasi dan seluruh data terkait berhasil dihapus.' });
  } catch (e) {
    console.error(e);
    res.status(500).json({ message: 'Gagal menghapus lokasi dari sistem.' });
  }
});

// ─────────────────────────────────────────────────────────────────────────────
// [API] UPDATE STATUS EVIDEN LOKASI
// ─────────────────────────────────────────────────────────────────────────────
app.patch('/api/lokasi/:id/status-eviden', checkAuth('admin'), async (req, res) => {
  try {
    const lokasiId = Number(req.params.id);
    if (!lokasiId) return res.status(400).json({ message: 'ID lokasi tidak valid.' });
    const { status } = req.body;
    if (!['need', 'progres', 'done'].includes(status)) {
      return res.status(400).json({ message: 'Status harus need, progres, atau done.' });
    }
    const existing = await db.get('SELECT id FROM lokasi WHERE id = ?', [lokasiId]);
    if (!existing) return res.status(404).json({ message: 'Lokasi tidak ditemukan.' });
    await db.run('UPDATE lokasi SET status_eviden = ? WHERE id = ?', [status, lokasiId]);

    // Pick up berkas ikut status eviden: progres/done = lokasi sudah diambil
    // (sedang dikerjakan), need = dikembalikan (angka kembali ke Need Evidence).
    try {
      const pickupSekarang = await db.get(
        "SELECT status FROM dokumen_status WHERE lokasi_id = ? AND jenis = 'pickup'", [lokasiId]
      );
      const statusPickup = status === 'need'
        ? (pickupSekarang && pickupSekarang.status === 'TERGENERATE' ? 'BELUM' : null)
        : 'TERGENERATE';
      if (statusPickup) {
        await tandaiDokumen(lokasiId, 'pickup', statusPickup, { updated_by: req.session.user.username });
      }
    } catch (e) { /* tabel dokumen_status belum siap — status eviden tetap tersimpan */ }

    res.json({ ok: true, status });
  } catch (e) {
    console.error(e);
    res.status(500).json({ message: 'Gagal update status eviden.' });
  }
});

// ─────────────────────────────────────────────────────────────────────────────
// [API] LOKASI DENGAN RAB
// ─────────────────────────────────────────────────────────────────────────────
app.get('/api/lokasi-dengan-rab', checkAuth('admin'), async (req, res) => {
  try {
    const rows = await db.all(`
      SELECT
        l.id, l.pid, l.nama_lokasi, l.tematik, l.detail_tambahan,
        COALESCE(rs.total_material, 0) AS total_material,
        COALESCE(rs.total_jasa,     0) AS total_jasa,
        COALESCE(rs.grand_total,    0) AS grand_total,
        COALESCE(rs.rab_count,      0) AS rab_count,
        rs.updated_at                  AS rab_updated_at
      FROM lokasi l
      LEFT JOIN rab_summary rs ON rs.lokasi_id = l.id
      ORDER BY l.id ASC
    `);
    res.json(rows);
  } catch (e) {
    console.error('Error /api/lokasi-dengan-rab:', e);
    res.status(500).json({ message: 'Gagal memuat data lokasi dengan RAB.', error: e.message });
  }
});


// ─────────────────────────────────────────────────────────────────────────────
// [API] Upload Berkas Lapangan
// ─────────────────────────────────────────────────────────────────────────────
app.post('/api/upload', checkAuth(), upload.array('files'), async (req, res) => {
  try {
    const { lokasi_id, kategori, sub_kategori } = req.body;

    if (!lokasi_id || !kategori) {
      return res.status(400).json({ message: 'Parameter data Lokasi dan Kategori wajib ditentukan.' });
    }

    const lokasiData = await db.get('SELECT nama_lokasi, tematik FROM lokasi WHERE id = ?', [lokasi_id]);
    if (!lokasiData) {
      return res.status(400).json({ message: 'ID Lokasi yang dipilih tidak terdaftar di sistem.' });
    }

    const namaTematikClean = (lokasiData.tematik || 'Umum').replace(/[^a-zA-Z0-9_-]/g, '_');
    const namaLokasiClean  = lokasiData.nama_lokasi.replace(/[^a-zA-Z0-9_-]/g, '_');
    const kategoriClean    = kategori.replace(/[^a-zA-Z0-9_-]/g, '_');
    const folderTujuan     = path.join(__dirname, 'uploads', namaTematikClean, namaLokasiClean, kategoriClean);
    pastikanFolderEksis(folderTujuan);

    if (!req.files || req.files.length === 0) {
      return res.status(400).json({ message: 'Harap pilih minimal satu berkas gambar.' });
    }

    // Form Pengukuran OPM (Node B): hanya gambar — PDF/Excel harus dikonversi
    // ke JPG/PNG dulu oleh user agar bisa ditampilkan penuh di laporan.
    if (kategori === 'FORM PENGUKURAN OPM') {
      const IMG_OK = ['.jpg', '.jpeg', '.png', '.webp', '.gif', '.bmp'];
      const buruk = req.files.find(f => !IMG_OK.includes((path.extname(f.originalname || '') || '').toLowerCase()));
      if (buruk) {
        return res.status(400).json({
          message: `Form Pengukuran OPM hanya menerima gambar (JPG/PNG/WEBP). "${buruk.originalname}" bukan gambar — simpan file Excel/PDF sebagai JPG/PNG dulu.`
        });
      }
    }

    const mapping = {
      '-EVIDEN PEMASANGAN TIANG':            'EPT',
      '-EVIDEN PENARIKAN':                   'EPEN',
      '-EVIDEN PEMASANGAN ODP':              'EPODP',
      '-EVIDENT HASIL UKUR OPM':             'EOPM',
      '-LAMPIRAN DATA PENGUKURAN OPM':       'LDP',
      '-MANCORE':                            'MCR',
      '-GROUNDING':                          'GR',
      '-LAMPIRAN HASIL UJI KEDALAMAN GALIAN':'LKDL',
      '-EVIDEN EVIDEN GALIAN DAN DIMENSI HANDHOLE': 'EGAL',
      '-EVIDEN DISMANTLING':                 'EDIST',
      '-EVIDEN COLUSER':                    'ECO',
      '-LAMPIRAN CO':                        'LCO',
      '-Instalasi Roset':                    'ROSET',
      '-SURAT PENUNJUKAN':                   'SURATPENUNJUKAN',
      '-SSBOQ':                              'SSBOQ',
      'MASUKAN EVIDEN PEKERJAAN':            'EEP',
      // ── KATEGORI TEMATIK HEM / MICRODEMAND ──
      'MANCORE':                             'MCR',
      'ABD (PDF/DWG/DXF)':                   'ABD',
      'EVIDEN BON MATERIAL':                 'EBM',
      'EVIDEN PENARIKAN KABEL':              'EPKABEL',
      'EVIDEN PEMASANGAN ACC':               'EACC',
      'EVIDEN PEMASANGAN TIANG':             'EPT',
      'EVIDEN INSTALASI ODP/OTB':            'EIODP',
      'EVIDEN JUMPER':                       'EJUMP',
      'EVIDEN REDAMAN ODP/OTB':              'ERED',
      'EVIDEN HDPE':                         'EHDPE',
      'EVIDEN OTDR':                         'EOTDR',
      'EVIDEN GALIAN':                       'EGAL',
      'EVIDEN HH AKSES':                     'EHHA',
      'EVIDEN HH ODP':                       'EHHO',
      // ── INPUT TEMATIK NODE B (LACT) ──
      'FORM PENGUKURAN OPM':                 'OPMFORM',
      'OTDR (.SOR)':                         'SOR',
    };

    const prefix = mapping[kategori] || 'FILE';

    const KATEGORI_PAKAI_TITIK = ['-EVIDENT HASIL UKUR OPM', '-EVIDEN PEMASANGAN ODP', '-EVIDEN COLUSER', 'EVIDEN INSTALASI ODP/OTB', 'EVIDEN REDAMAN ODP/OTB'];
    const pakaiTitik  = KATEGORI_PAKAI_TITIK.includes(kategori);
    const titikMentah = pakaiTitik ? String(sub_kategori || '1') : (sub_kategori || null);
    const titik       = titikMentah ? titikMentah.replace(/[\\/:*?"<>|]+/g, '-').trim() : titikMentah;

    // Ambil daftar port jika ada (untuk penamaan manual port OPM)
    let explicitPorts = [];
    if (req.body.ports) {
      try {
        explicitPorts = JSON.parse(req.body.ports);
      } catch (e) {
        explicitPorts = Array.isArray(req.body.ports) ? req.body.ports : [req.body.ports];
      }
    }

    // Ambil daftar step jika ada (untuk penamaan manual eviden Pemasangan ODP:
    // EPODP-{ODP}-SBL / -PRG / -SDH). Misal body.steps = '["SBL","PRG","SDH"]'.
    let stepList = [];
    if (req.body.steps) {
      try {
        stepList = JSON.parse(req.body.steps);
      } catch (e) {
        stepList = Array.isArray(req.body.steps) ? req.body.steps : [req.body.steps];
      }
    }

    const namaBaru = [];
    for (let i = 0; i < req.files.length; i++) {
      const file = req.files[i];
      // Ekstensi asli berkas: khusus kategori ABD, PDF/DWG/DXF disimpan apa adanya;
      // kategori lain tetap dikonversi ke JPG seperti sebelumnya (termasuk PDF lama).
      let namaFileFinal;
      const ekstAsli = (path.extname(file.originalname || '') || '').toLowerCase();
      const RAW_EXT = kategori === 'ABD (PDF/DWG/DXF)'
        ? ['.pdf', '.dwg', '.dxf']
        : kategori === 'INPUT RFC' ? ['.pdf']
        // Node B: file OTDR .SOR disimpan apa adanya; Form OPM gambar tanpa
        // resize agar teks form tetap terbaca.
        : kategori === 'OTDR (.SOR)' ? ['.sor']
        // EVIDEN OTDR (HEM / MICRODEMAND / Node B): .sor disimpan apa adanya,
        // bukan dikonversi ke JPG (konversi binary .sor mustahil — sharp gagal).
        : kategori === 'EVIDEN OTDR' ? ['.sor']
        : kategori === 'FORM PENGUKURAN OPM' ? ['.jpg', '.jpeg', '.png', '.webp', '.gif', '.bmp']
        : [];
      const ekstFinal = RAW_EXT.includes(ekstAsli) ? ekstAsli : '.jpg';

      if (pakaiTitik) {
        // 1) Step eksplisit (Pemasangan ODP: SBL/PRG/SDH)
        const step = stepList[i];
        // 2) Nomor port eksplisit dari user (untuk OPM)
        const portUrut = explicitPorts[i];
        if (step) {
          const stepBersih = String(step).replace(/[\\/:*?"<>|]+/g, '-').trim();
          namaFileFinal = `${prefix}-${titik}-${stepBersih}${ekstFinal}`;
        } else if (portUrut) {
          namaFileFinal = `${prefix}-${titik}-P${portUrut}${ekstFinal}`;
        } else {
          const countResult = await db.get(
            'SELECT COUNT(*) as total FROM berkas WHERE lokasi_id = ? AND kategori = ? AND sub_kategori = ?',
            [lokasi_id, kategori, titik]
          );
          const fotoKe = countResult.total + i + 1;
          namaFileFinal = `${prefix}-${titik}-${fotoKe}${ekstFinal}`;
        }
      } else {
        const countResult = await db.get(
          'SELECT COUNT(*) as total FROM berkas WHERE lokasi_id = ? AND kategori = ?',
          [lokasi_id, kategori]
        );
        // Nomor urut per kategori bisa TERPAKAI ULANG setelah ada penghapusan,
        // sehingga nama file (mis. EEP27.jpg) dipakai oleh beberapa sub-grup.
        // Karena semua sub-grup berbagi SATU folder, file lama tertimpa → foto
        // yang sama muncul di banyak grup (dobel) dan tiap grup menampilkan foto
        // yang salah (pindah kategori). Cari nomor yang benar-benar belum dipakai
        // (cek DB maupun file di disk) sebelum menulis.
        let nomorUrut = countResult.total + i + 1;
        let namaCoba  = `${prefix}${nomorUrut}${ekstFinal}`;
        const sudahDipakai = async (nm) => {
          const r = await db.get(
            'SELECT 1 x FROM berkas WHERE lokasi_id = ? AND kategori = ? AND file_name = ?',
            [lokasi_id, kategori, nm]);
          return !!r || fs.existsSync(path.join(folderTujuan, nm));
        };
        while (await sudahDipakai(namaCoba)) {
          nomorUrut++;
          namaCoba = `${prefix}${nomorUrut}${ekstFinal}`;
        }
        namaFileFinal = namaCoba;
      }

      const fileDestination = path.join(folderTujuan, namaFileFinal);
      if (RAW_EXT.includes(ekstFinal)) {
        // PDF / DWG / DXF: simpan byte asli tanpa konversi (untuk ABD & berkas teknis)
        fs.writeFileSync(fileDestination, file.buffer);
      } else {
        await sharp(file.buffer)
          .resize({
              width: 2480,
              height: 3508,
              fit: 'inside', // Wajib 'inside' agar tidak ada bagian yang terpotong
              withoutEnlargement: true // Mencegah gambar kecil jadi pecah karena di-upscale
          })
          .withMetadata() // Pertahankan EXIF (tanggal/GPS/software) untuk verifikasi keaslian
          // A4 @300 DPI + kualitas tinggi: gambar tetap tajam saat di-zoom di Word/PDF.
          .jpeg({ quality: 90, chromaSubsampling: '4:4:4' })
          .toFile(fileDestination);
      }

      namaBaru.push(namaFileFinal);

      await db.run(
        'INSERT INTO berkas (lokasi_id, kategori, sub_kategori, file_name) VALUES (?, ?, ?, ?)',
        [lokasi_id, kategori, titik, namaFileFinal]
      );

      // Rekam status port (OPM → opm_port, Redaman → redaman_port)
      if (pakaiTitik && explicitPorts[i]) {
        const portTable = kategori === 'EVIDEN REDAMAN ODP/OTB' ? 'redaman_port' : 'opm_port';
        await db.run(
          `INSERT INTO ${portTable} (lokasi_id, odp, port, status) VALUES (?, ?, ?, 'ukur')
           ON CONFLICT(lokasi_id, odp, port) DO UPDATE SET status = 'ukur'`,
          [lokasi_id, titik, Number(explicitPorts[i])]
        );
      }
    }

    await autoSetDoneIfComplete(lokasi_id);

    // Verifikasi otomatis foto baru (aktif bila VERIFY_AUTO=1 di .env)
    if (verifCfg.autoVerify && namaBaru.length) {
      verif.verifikasiOtomatisUpload(db, lokasi_id, namaBaru, req.session.user.username)
        .catch(e => console.warn('VERIFY_AUTO:', e.message || e));
    }

    res.json({ success: true, message: 'Foto berhasil disimpan dengan penamaan otomatis!' });
  } catch (e) {
    console.error('Kesalahan Sistem Saat Upload:', e);
    res.status(500).json({
      message: 'Terjadi kesalahan internal server saat menyimpan berkas.',
      detail: String(e && e.message || e)
    });
  }
});

// ─────────────────────────────────────────────────────────────────────────────
// [API] Upload KML/KMZ → render peta OSM → simpan PNG sebagai berkas biasa
// Kategori: -PETA RUTE (KML/KMZ)
// ─────────────────────────────────────────────────────────────────────────────
app.post('/api/upload/kml', checkAuth(), upload.single('file'), async (req, res) => {
  try {
    const lokasi_id = req.body.lokasi_id;
    if (!lokasi_id) return res.status(400).json({ message: 'Lokasi kerja wajib dipilih.' });
    if (!req.file) return res.status(400).json({ message: 'File KML/KMZ wajib diunggah.' });

    const namaAsli = req.file.originalname || '';
    const ext = path.extname(namaAsli).toLowerCase();
    if (!['.kml', '.kmz'].includes(ext)) {
      return res.status(400).json({ message: 'Format berkas harus .kml atau .kmz.' });
    }

    const lokasi = await db.get('SELECT tematik, nama_lokasi FROM lokasi WHERE id = ?', [lokasi_id]);
    if (!lokasi) return res.status(400).json({ message: 'Lokasi tidak ditemukan.' });

    const xml = bacaKml(req.file.buffer);
    const features = parseKml(xml);
    if (!features.length) {
      return res.status(422).json({
        message: 'File tidak mengandung geometri yang bisa dipetakan (Point / LineString / Polygon / track GPS).',
      });
    }

    const kategoriKml = req.body.kategori || KATEGORI_PETA;
    const ntc = (lokasi.tematik || 'Umum').replace(/[^a-zA-Z0-9_-]/g, '_');
    const nlc = lokasi.nama_lokasi.replace(/[^a-zA-Z0-9_-]/g, '_');
    const kc  = kategoriKml.replace(/[^a-zA-Z0-9_-]/g, '_');
    const folder = path.join(__dirname, 'uploads', ntc, nlc, kc);
    pastikanFolderEksis(folder);

    const base = 'PETA_' + Date.now();
    const pngPath = path.join(folder, base + '.png');
    await renderPeta(features, pngPath);

    // Simpan KML/KMZ ASLI (untuk Google Earth) — nama setara dengan PNG render.
    fs.writeFileSync(path.join(folder, base + ext), req.file.buffer);

    const sub = (path.basename(namaAsli, ext) || 'Peta Rute').replace(/[\\/:*?"<>|]+/g, '-').trim();
    await db.run(
      'INSERT INTO berkas (lokasi_id, kategori, sub_kategori, file_name) VALUES (?, ?, ?, ?)',
      [lokasi_id, kategoriKml, sub, base + '.png']
    );
    const inserted = await db.get(
      'SELECT id FROM berkas WHERE lokasi_id = ? AND kategori = ? AND file_name = ? ORDER BY id DESC LIMIT 1',
      [lokasi_id, kategoriKml, base + '.png']
    );

    const ringkasan = {};
    features.forEach((f) => { ringkasan[f.type] = (ringkasan[f.type] || 0) + 1; });

    await autoSetDoneIfComplete(lokasi_id);

    res.json({
      success: true,
      message: 'Peta rute berhasil diproses dan disimpan.',
      file: base + '.png',
      url: '/uploads/' + ntc + '/' + nlc + '/' + kc + '/' + base + '.png',
      pdf: inserted ? '/api/download-peta/' + inserted.id + '/pdf' : null,
      kml: inserted ? '/api/kml/asli/' + inserted.id : null,
      ringkasan,
    });
  } catch (e) {
    console.error('Error upload KML/KMZ:', e);
    res.status(500).json({ message: 'Gagal memproses KML/KMZ.', detail: String(e && e.message || e) });
  }
});

// ─────────────────────────────────────────────────────────────────────────────
// [API] Download PDF hasil render peta rute (KML/KMZ)
// ─────────────────────────────────────────────────────────────────────────────
app.get('/api/download-peta/:berkas_id/pdf', checkAuth(), async (req, res) => {
  try {
    const b = await db.get('SELECT * FROM berkas WHERE id = ?', [req.params.berkas_id]);
    if (!b) return res.status(404).json({ message: 'Berkas tidak ditemukan.' });
    if (b.kategori !== KATEGORI_PETA && b.kategori !== 'KML (KML)') {
      return res.status(400).json({ message: 'Berkas ini bukan hasil render peta rute.' });
    }
    const lokasi = await db.get('SELECT tematik, nama_lokasi FROM lokasi WHERE id = ?', [b.lokasi_id]);
    if (!lokasi) return res.status(404).json({ message: 'Lokasi tidak ditemukan.' });

    const ntc = (lokasi.tematik || 'Umum').replace(/[^a-zA-Z0-9_-]/g, '_');
    const nlc = lokasi.nama_lokasi.replace(/[^a-zA-Z0-9_-]/g, '_');
    const kc  = b.kategori.replace(/[^a-zA-Z0-9_-]/g, '_');
    const pngPath = path.join(__dirname, 'uploads', ntc, nlc, kc, b.file_name);
    if (!fs.existsSync(pngPath)) return res.status(404).json({ message: 'File peta tidak ditemukan di server.' });

    const pdfBuf = await pngToPdf(pngPath);
    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', 'attachment; filename="PETA_' + nlc + '.pdf"');
    res.send(pdfBuf);
  } catch (e) {
    console.error('Error download PDF peta:', e);
    res.status(500).json({ message: 'Gagal membuat PDF peta.', detail: String(e && e.message || e) });
  }
});

// ─────────────────────────────────────────────────────────────────────────────
// [API] Download KML/KMZ ASLI — dibuka penuh layar di Google Earth
// ─────────────────────────────────────────────────────────────────────────────
app.get('/api/kml/asli/:berkas_id', checkAuth(), async (req, res) => {
  try {
    const b = await db.get('SELECT * FROM berkas WHERE id = ?', [req.params.berkas_id]);
    if (!b) return res.status(404).json({ message: 'Berkas tidak ditemukan.' });
    if (!KATEGORI_KML_FULL.has(b.kategori)) {
      return res.status(400).json({ message: 'Berkas ini bukan kategori peta/KML.' });
    }
    const lokasi = await db.get('SELECT tematik, nama_lokasi FROM lokasi WHERE id = ?', [b.lokasi_id]);
    if (!lokasi) return res.status(404).json({ message: 'Lokasi tidak ditemukan.' });

    const ntc = (lokasi.tematik || 'Umum').replace(/[^a-zA-Z0-9_-]/g, '_');
    const nlc = lokasi.nama_lokasi.replace(/[^a-zA-Z0-9_-]/g, '_');
    const kc  = b.kategori.replace(/[^a-zA-Z0-9_-]/g, '_');
    const folder = path.join(__dirname, 'uploads', ntc, nlc, kc);
    const asli = cariKmlAsli(folder, baseTanpaExt(b.file_name));
    if (!asli) {
      return res.status(404).json({
        message: 'File KML asli belum tersedia (berkas lama hanya menyimpan gambar render). Unggah ulang .kml/.kmz untuk menyimpan versi aslinya.',
      });
    }
    const ext = path.extname(asli).toLowerCase();
    const nama = ((b.sub_kategori || 'Peta_Rute') + ext).replace(/[\\/:*?"<>|]+/g, '-').replace(/\s+/g, '_');
    res.setHeader('Content-Type', ext === '.kmz'
      ? 'application/vnd.google-earth.kmz'
      : 'application/vnd.google-earth.kml+xml; charset=utf-8');
    res.setHeader('Content-Disposition', 'attachment; filename="' + nama + '"');
    fs.createReadStream(asli).pipe(res);
  } catch (e) {
    console.error('Error download KML asli:', e);
    res.status(500).json({ message: 'Gagal mengunduh KML asli.', detail: String(e && e.message || e) });
  }
});

// ─────────────────────────────────────────────────────────────────────────────
// [API] KML gabungan SELURUH hasil satu tematik → 1 file, tampil penuh layar
//       di Google Earth (folder per lokasi + LookAt membentang semua rute).
// ─────────────────────────────────────────────────────────────────────────────
app.get('/api/kml/tematik/:lokasi_id', checkAuth(), async (req, res) => {
  try {
    const lokasi = await db.get('SELECT tematik FROM lokasi WHERE id = ?', [req.params.lokasi_id]);
    if (!lokasi) return res.status(404).json({ message: 'Lokasi tidak ditemukan.' });
    const tematik = lokasi.tematik || 'Umum';

    const rows = await db.all(
      `SELECT b.*, l.nama_lokasi FROM berkas b
       JOIN lokasi l ON l.id = b.lokasi_id
       WHERE l.tematik = ? AND b.kategori IN ('KML (KML)', ?)
       ORDER BY l.nama_lokasi ASC, b.id ASC`,
      [tematik, KATEGORI_PETA]
    );

    // Kelompokkan per lokasi (nama tampilan), cari file KML asli tiap baris.
    const perLokasi = new Map();
    for (const b of rows) {
      if (!perLokasi.has(b.nama_lokasi)) perLokasi.set(b.nama_lokasi, []);
      perLokasi.get(b.nama_lokasi).push(b);
    }

    const ntc = (tematik || 'Umum').replace(/[^a-zA-Z0-9_-]/g, '_');
    const folders = [];
    const semuaTitik = [];
    let totalGeom = 0, lokasiDengan = 0;

    for (const [namaLok, list] of perLokasi) {
      const isi = [];
      for (const b of list) {
        const nlc = b.nama_lokasi.replace(/[^a-zA-Z0-9_-]/g, '_');
        const kc  = b.kategori.replace(/[^a-zA-Z0-9_-]/g, '_');
        const asli = cariKmlAsli(path.join(__dirname, 'uploads', ntc, nlc, kc), baseTanpaExt(b.file_name));
        if (!asli) continue;
        let feats = [];
        try { feats = parseKml(bacaKml(fs.readFileSync(asli))); } catch (err) { continue; }
        for (const f of feats) {
          const pm = pmKml(f);
          if (!pm) continue;
          isi.push(pm);
          (f.points || []).forEach(p => semuaTitik.push(p));
          totalGeom++;
        }
      }
      if (isi.length) {
        folders.push('<Folder><name>' + escKml(namaLok) + '</name>\n' + isi.join('\n') + '\n</Folder>');
        lokasiDengan++;
      }
    }

    if (!totalGeom) {
      return res.status(404).json({
        message: 'Belum ada file KML asli untuk tematik "' + tematik + '". Unggah ulang .kml/.kmz pada lokasi terkait (berkas lama hanya menyimpan gambar render).',
      });
    }

    // Kamera LookAt agar Google Earth langsung membentang seluruh area (full layar)
    let minLon = Infinity, maxLon = -Infinity, minLat = Infinity, maxLat = -Infinity;
    for (const p of semuaTitik) {
      if (p[0] < minLon) minLon = p[0];
      if (p[0] > maxLon) maxLon = p[0];
      if (p[1] < minLat) minLat = p[1];
      if (p[1] > maxLat) maxLat = p[1];
    }
    const ctrLon = (minLon + maxLon) / 2;
    const ctrLat = (minLat + maxLat) / 2;
    const lintang = (maxLat - minLat) * 111320;
    const bujur  = (maxLon - minLon) * 111320 * Math.max(0.2, Math.cos(ctrLat * Math.PI / 180));
    const range  = Math.round(Math.max(lintang, bujur) * 1.8 + 400);

    const kml =
      '<?xml version="1.0" encoding="UTF-8"?>\n' +
      '<kml xmlns="http://www.opengis.net/kml/2.2">\n<Document>\n' +
      '<name>KML ' + escKml(tematik) + ' (' + lokasiDengan + ' lokasi · ' + totalGeom + ' geometri)</name>\n' +
      '<description>Seluruh hasil peta/KML tematik ' + escKml(tematik) + ' — buka di Google Earth.</description>\n' +
      '<LookAt><longitude>' + ctrLon.toFixed(6) + '</longitude><latitude>' + ctrLat.toFixed(6) + '</latitude>' +
      '<range>' + range + '</range><tilt>0</tilt><heading>0</heading><altitudeMode>clampedToGround</altitudeMode></LookAt>\n' +
      '<Style id="sRute"><LineStyle><color>ffff0000</color><width>5</width></LineStyle>' +
      '<PolyStyle><color>4dff0000</color></PolyStyle></Style>\n' +
      '<Style id="sTitik"><IconStyle><scale>1.1</scale>' +
      '<Icon><href>http://maps.google.com/mapfiles/kml/pushpin/ylw-pushpin.png</href></Icon></IconStyle>' +
      '<LabelStyle><scale>0.8</scale></LabelStyle></Style>\n' +
      folders.join('\n') + '\n' +
      '</Document>\n</kml>\n';

    const namaFile = 'KML_' + tematik.replace(/[^A-Za-z0-9_-]+/g, '_') + '.kml';
    res.setHeader('Content-Type', 'application/vnd.google-earth.kml+xml; charset=utf-8');
    res.setHeader('Content-Disposition', 'attachment; filename="' + namaFile + '"');
    res.send(kml);
  } catch (e) {
    console.error('Error KML tematik:', e);
    res.status(500).json({ message: 'Gagal membuat KML tematik.', detail: String(e && e.message || e) });
  }
});

// ─────────────────────────────────────────────────────────────────────────────
// [API] Nomor titik berikutnya (EOPM & EPODP)
// ─────────────────────────────────────────────────────────────────────────────
app.get('/api/next-titik/:lokasi_id/:kategori', async (req, res) => {
  try {
    const { lokasi_id, kategori } = req.params;
    const rows = await db.all(
      'SELECT DISTINCT sub_kategori FROM berkas WHERE lokasi_id = ? AND kategori = ?',
      [lokasi_id, kategori]
    );
    const angka = rows.map(r => parseInt(r.sub_kategori, 10)).filter(n => !isNaN(n));
    const titikBerikutnya = angka.length > 0 ? Math.max(...angka) + 1 : 1;
    res.json({ titikBerikutnya });
  } catch (e) {
    console.error(e);
    res.status(500).json({ message: 'Gagal mengambil nomor titik berikutnya.' });
  }
});

// ═════════════════════════════════════════════════════════════════════════════
// [API] Riwayat / Info Port ODP (untuk validasi user & pemantauan admin)
// Kategori EVITEM HASIL UKUR OPM (EOPM): file dinamai EOPM-{ODP}-{port}.jpg
// di mana {port} menandakan nomor port yang sudah terisi.
// ═════════════════════════════════════════════════════════════════════════════
app.get('/api/opm/:lokasi_id', async (req, res) => {
  try {
    const lokasi_id = req.params.lokasi_id;
    const kategori  = req.query.kategori || '-EVIDENT HASIL UKUR OPM';

    // Ambil kapasitas & status port dari tabel OPM (sumber utama)
    const odps = await db.all('SELECT odp, kapasitas FROM opm_odp WHERE lokasi_id = ? ORDER BY odp ASC', [lokasi_id]);
    const portRows = await db.all('SELECT odp, port, status FROM opm_port WHERE lokasi_id = ? ORDER BY port ASC', [lokasi_id]);

    // Ambil nama file foto untuk port "ukur" dari tabel berkas
    const fotoRows = await db.all(
      'SELECT sub_kategori, file_name FROM berkas WHERE lokasi_id = ? AND kategori = ? ORDER BY id ASC',
      [lokasi_id, kategori]
    );

    // GABUNGKAN ODP: dari tabel opm_odp DAN dari foto berkas yang sudah ada.
    // (ODP lama yang diinput sebelum fitur OPM tidak ada di tabel opm_odp,
    //  sehingga sumber fallback-nya adalah sub_kategori di tabel berkas.)
    const odpMap = new Map();
    odps.forEach(g => odpMap.set(g.odp, g.kapasitas || 8));
    fotoRows.forEach(f => {
      const nm = (f.sub_kategori || '').trim();
      if (nm && !odpMap.has(nm)) odpMap.set(nm, 8);
    });

    const daftar = Array.from(odpMap.entries()).map(([nm, kap]) => {
      // Foto per ODP → {port { file }}
      const fotoMap = {};
      fotoRows.filter(f => (f.sub_kategori || '').trim() === nm).forEach(f => {
        const m = /-P?(\d+)\.(jpg|jpeg|png|webp|pdf)$/i.exec(f.file_name || '');
        const p = m ? parseInt(m[1], 10) : null;
        if (p) fotoMap[p] = f.file_name;
      });
      // Status per port (ukur = foto, isi = user).
      const portMap = {};
      portRows.filter(p => p.odp === nm).forEach(p => {
        portMap[p.port] = p.status;
      });

      // Kapasitas efektif: ikuti kapasitas terdaftar, tapi naikkan jika ada
      // foto / status port di nomor lebih besar agar tidak terpotong.
      let efektifKap = kap || 8;
      const maksFoto = Object.keys(fotoMap).map(Number);
      const maksStatus = Object.keys(portMap).map(Number);
      if (maksFoto.length)   efektifKap = Math.max(efektifKap, Math.max(...maksFoto));
      if (maksStatus.length) efektifKap = Math.max(efektifKap, Math.max(...maksStatus));

      const ports = [];
      for (let p = 1; p <= efektifKap; p++) {
        const st = portMap[p] || (fotoMap[p] ? 'ukur' : 'kosong');
        ports.push({ port: p, status: st, file: fotoMap[p] || null });
      }
      return {
        nama: nm,
        kapasitas: efektifKap,
        jumlah: ports.filter(p => p.status === 'ukur').length,
        maxPort: ports.filter(p => p.status === 'ukur').reduce((mx, p) => Math.max(mx, p.port), 0),
        ports,
      };
    }).sort((a, b) => a.nama.localeCompare(b.nama));

    const isiCount = portRows.filter(p => p.status === 'isi').length;
    res.json({ daftar, totalFoto: fotoRows.length, totalIsi: isiCount });
  } catch (e) {
    console.error(e);
    res.status(500).json({ message: 'Gagal memuat informasi port ODP.' });
  }
});

// Inisialisasi / kasih tahu server: ODP dengan kapasitas port tertentu
app.post('/api/opm/init', checkAuth(), async (req, res) => {
  try {
    const { lokasi_id, odp, kapasitas } = req.body;
    if (!lokasi_id || !odp) return res.status(400).json({ message: 'Lokasi & nama ODP wajib diisi.' });
    const kap = [8, 16].includes(Number(kapasitas)) ? Number(kapasitas) : 8;
    await db.run(
      `INSERT INTO opm_odp (lokasi_id, odp, kapasitas) VALUES (?, ?, ?)
       ON CONFLICT(lokasi_id, odp) DO UPDATE SET kapasitas = excluded.kapasitas`,
      [lokasi_id, odp, kap]
    );
    res.json({ success: true });
  } catch (e) {
    console.error(e);
    res.status(500).json({ message: 'Gagal menyimpan ODP.' });
  }
});

// Tandai satu port sebagai "isii" (ada user, tidak ada foto)
app.post('/api/opm/isi', checkAuth(), async (req, res) => {
  try {
    const { lokasi_id, odp, port } = req.body;
    if (!lokasi_id || !odp || !port) return res.status(400).json({ message: 'Lokasi, ODP, dan port wajib diisi.' });
    await db.run(
      `INSERT INTO opm_port (lokasi_id, odp, port, status) VALUES (?, ?, ?, 'isi')
       ON CONFLICT(lokasi_id, odp, port) DO UPDATE SET status = 'isi'`,
      [lokasi_id, odp, Number(port)]
    );
    res.json({ success: true });
  } catch (e) {
    console.error(e);
    res.status(500).json({ message: 'Gagal menyimpan status port.' });
  }
});

// ═════════════════════════════════════════════════════════════════════════════
// EVIDEN REDAMAN ODP/OTB — wizard per port (mirror OPM, tabel terpisah)
// ═════════════════════════════════════════════════════════════════════════════
const KATEGORI_REDAMAN = 'EVIDEN REDAMAN ODP/OTB';

app.get('/api/redaman/:lokasi_id', async (req, res) => {
  try {
    const lokasi_id = req.params.lokasi_id;
    const kategori  = KATEGORI_REDAMAN;

    const odps = await db.all('SELECT odp, kapasitas FROM redaman_odp WHERE lokasi_id = ? ORDER BY odp ASC', [lokasi_id]);
    const portRows = await db.all('SELECT odp, port, status FROM redaman_port WHERE lokasi_id = ? ORDER BY port ASC', [lokasi_id]);
    const fotoRows = await db.all(
      'SELECT sub_kategori, file_name FROM berkas WHERE lokasi_id = ? AND kategori = ? ORDER BY id ASC',
      [lokasi_id, kategori]
    );

    const odpMap = new Map();
    odps.forEach(g => odpMap.set(g.odp, g.kapasitas || 8));
    fotoRows.forEach(f => {
      const nm = (f.sub_kategori || '').trim();
      if (nm && !odpMap.has(nm)) odpMap.set(nm, 8);
    });

    const daftar = Array.from(odpMap.entries()).map(([nm, kap]) => {
      const fotoMap = {};
      fotoRows.filter(f => (f.sub_kategori || '').trim() === nm).forEach(f => {
        const m = /-P?(\d+)\.(jpg|jpeg|png|webp|pdf)$/i.exec(f.file_name || '');
        const p = m ? parseInt(m[1], 10) : null;
        if (p) fotoMap[p] = f.file_name;
      });
      const portMap = {};
      portRows.filter(p => p.odp === nm).forEach(p => {
        portMap[p.port] = p.status;
      });

      let efektifKap = kap || 8;
      const maksFoto = Object.keys(fotoMap).map(Number);
      const maksStatus = Object.keys(portMap).map(Number);
      if (maksFoto.length)   efektifKap = Math.max(efektifKap, Math.max(...maksFoto));
      if (maksStatus.length) efektifKap = Math.max(efektifKap, Math.max(...maksStatus));

      const ports = [];
      for (let p = 1; p <= efektifKap; p++) {
        const st = portMap[p] || (fotoMap[p] ? 'ukur' : 'kosong');
        ports.push({ port: p, status: st, file: fotoMap[p] || null });
      }
      return {
        nama: nm,
        kapasitas: efektifKap,
        jumlah: ports.filter(p => p.status === 'ukur').length,
        maxPort: ports.filter(p => p.status === 'ukur').reduce((mx, p) => Math.max(mx, p.port), 0),
        ports,
      };
    }).sort((a, b) => a.nama.localeCompare(b.nama));

    res.json({ daftar, totalFoto: fotoRows.length, totalIsi: portRows.filter(p => p.status === 'isi').length });
  } catch (e) {
    console.error(e);
    res.status(500).json({ message: 'Gagal memuat informasi port redaman.' });
  }
});

app.post('/api/redaman/init', checkAuth(), async (req, res) => {
  try {
    const { lokasi_id, odp, kapasitas } = req.body;
    if (!lokasi_id || !odp) return res.status(400).json({ message: 'Lokasi & nama ODP/OTB wajib diisi.' });
    const kap = [8, 16].includes(Number(kapasitas)) ? Number(kapasitas) : 8;
    await db.run(
      `INSERT INTO redaman_odp (lokasi_id, odp, kapasitas) VALUES (?, ?, ?)
       ON CONFLICT(lokasi_id, odp) DO UPDATE SET kapasitas = excluded.kapasitas`,
      [lokasi_id, odp, kap]
    );
    res.json({ success: true });
  } catch (e) {
    console.error(e);
    res.status(500).json({ message: 'Gagal menyimpan ODP/OTB.' });
  }
});

app.post('/api/redaman/isi', checkAuth(), async (req, res) => {
  try {
    const { lokasi_id, odp, port } = req.body;
    if (!lokasi_id || !odp || !port) return res.status(400).json({ message: 'Lokasi, ODP/OTB, dan port wajib diisi.' });
    await db.run(
      `INSERT INTO redaman_port (lokasi_id, odp, port, status) VALUES (?, ?, ?, 'isi')
       ON CONFLICT(lokasi_id, odp, port) DO UPDATE SET status = 'isi'`,
      [lokasi_id, odp, Number(port)]
    );
    res.json({ success: true });
  } catch (e) {
    console.error(e);
    res.status(500).json({ message: 'Gagal menyimpan status port.' });
  }
});

// ═════════════════════════════════════════════════════════════════════════════
// [API] BOQ AKTUAL — baris paste Excel (mirror input RAB, tabel terpisah)
// ═════════════════════════════════════════════════════════════════════════════
app.get('/api/boq/:lokasi_id', checkAuth(), async (req, res) => {
  try {
    const rows = await db.all(
      'SELECT * FROM boq_aktual WHERE lokasi_id = ? ORDER BY id ASC',
      [req.params.lokasi_id]
    );
    res.json(rows);
  } catch (e) {
    console.error(e);
    res.status(500).json({ message: 'Gagal memuat data BOQ aktual.' });
  }
});

app.post('/api/boq/:lokasi_id', checkAuth(), async (req, res) => {
  try {
    const lokasiId = req.params.lokasi_id;
    const { items } = req.body;
    if (!Array.isArray(items)) return res.status(400).json({ message: 'Body harus berisi "items" berupa array.' });
    await db.run('DELETE FROM boq_aktual WHERE lokasi_id = ?', [lokasiId]);
    const stmt = await db.prepare(
      'INSERT INTO boq_aktual (lokasi_id, designator, uraian, satuan, material, jasa, volume) VALUES (?, ?, ?, ?, ?, ?, ?)'
    );
    for (const it of items) {
      await stmt.run([
        lokasiId,
        String(it.designator || '').trim(),
        String(it.uraian || '').trim(),
        String(it.satuan || '').trim(),
        Number(it.material) || 0,
        Number(it.jasa) || 0,
        Number(it.volume) || 0,
      ]);
    }
    await stmt.finalize();

    await autoSetDoneIfComplete(lokasiId);

    res.json({ success: true, message: 'BOQ Aktual tersimpan.' });
  } catch (e) {
    console.error(e);
    res.status(500).json({ message: 'Gagal menyimpan BOQ aktual.' });
  }
});

app.delete('/api/boq/:id', checkAuth(), async (req, res) => {
  try {
    await db.run('DELETE FROM boq_aktual WHERE id = ?', [req.params.id]);
    res.json({ success: true });
  } catch (e) {
    console.error(e);
    res.status(500).json({ message: 'Gagal menghapus baris BOQ.' });
  }
});

// ─────────────────────────────────────────────────────────────────────────────
// [API] BOQ BARANG TIBA (input TEMATIK Node B — LACT)
// Baris: no, jenis_material, satuan, volume, keterangan
// ─────────────────────────────────────────────────────────────────────────────
app.get('/api/barang-tiba/:lokasi_id', checkAuth(), async (req, res) => {
  try {
    const rows = await db.all(
      'SELECT id, no, jenis_material, satuan, volume, keterangan, created_at FROM barang_tiba WHERE lokasi_id = ? ORDER BY no ASC, id ASC',
      [req.params.lokasi_id]
    );
    res.json(rows);
  } catch (e) {
    console.error(e);
    res.status(500).json({ message: 'Gagal memuat data barang tiba.' });
  }
});

app.post('/api/barang-tiba/:lokasi_id', checkAuth(), async (req, res) => {
  try {
    const lokasiId = req.params.lokasi_id;
    const { items } = req.body;
    if (!Array.isArray(items)) return res.status(400).json({ message: 'Body harus berisi "items" berupa array.' });
    // Transaksional: DELETE + INSERT semua baris dalam satu batch; gagal di
    // tengah → rollback, data lama tidak hilang percuma.
    await db.batch(async () => {
      await db.run('DELETE FROM barang_tiba WHERE lokasi_id = ?', [lokasiId]);
      const stmt = await db.prepare(
        'INSERT INTO barang_tiba (lokasi_id, no, jenis_material, satuan, volume, keterangan) VALUES (?, ?, ?, ?, ?, ?)'
      );
      let no = 0;
      for (const it of items) {
        no += 1;
        await stmt.run([
          lokasiId,
          no,
          String(it.jenis_material || '').trim(),
          String(it.satuan || '').trim(),
          String(it.volume || '').trim(),
          String(it.keterangan || '').trim(),
        ]);
      }
      await stmt.finalize();
    });

    await autoSetDoneIfComplete(lokasiId);

    res.json({ success: true, message: 'BOQ Barang Tiba tersimpan.' });
  } catch (e) {
    console.error(e);
    res.status(500).json({ message: 'Gagal menyimpan BOQ barang tiba.' });
  }
});

// ═════════════════════════════════════════════════════════════════════════════
// Helper: susun data Eviden Pemasangan ODP (EPODP) untuk LOOOP di template Word
// dan untuk API monitoring.
// Dipakai docxtemplater: {#epodp_list} ... {#steps}{label}{#adaFoto}{%foto_path}{/adaFoto}{/steps} {/epodp_list}
// Setiap ODP = 3 step: SEBELUM (-SBL) → PROGRES (-PRG) → SESUDAH (-SDH).
// Data lama bernomor (EPODP-{ODP}-1/2/3) dipetakan fallback ke SBL/PRG/SDH.
// ═════════════════════════════════════════════════════════════════════════════
// ═════════════════════════════════════════════════════════════════════════════
// Helper: susun data Eviden Pemasangan ODP (EPODP)
// ═════════════════════════════════════════════════════════════════════════════
const EPODP_STEP_INFO = [
  { k: 'SBL', label: 'SEBELUM' },
  { k: 'PRG', label: 'PROGRES' },
  { k: 'SDH', label: 'SESUDAH' },
];

async function susunDataEpodp(lokasi_id, gambarBasePath, extraInfo = {}) {
  const kategori = '-EVIDEN PEMASANGAN ODP';
  const rows = await db.all(
    'SELECT id, sub_kategori, file_name FROM berkas WHERE lokasi_id = ? AND kategori = ? ORDER BY id ASC',
    [lokasi_id, kategori]
  );

  const byOdp = {};
  rows.forEach(r => {
    const nm = (r.sub_kategori || '').trim();
    if (!nm) return;
    if (!byOdp[nm]) byOdp[nm] = [];
    byOdp[nm].push(r);
  });

  return await Promise.all(Object.keys(byOdp).sort((a, b) => a.localeCompare(b)).map(async nama => {
    const files = byOdp[nama];
    const steps = {};
    const numeric = [];
    
    files.forEach(f => {
      const m = /-(SBL|PRG|SDH)\.(jpg|jpeg|png|webp|pdf)$/i.exec(f.file_name || '');
      if (m) {
        const k = m[1].toUpperCase();
        if (!steps[k]) steps[k] = f;
      } else {
        const nm2 = /-P?(\d+)\.(jpg|jpeg|png|webp|pdf)$/i.exec(f.file_name || '');
        const n = nm2 ? parseInt(nm2[1], 10) : 0;
        if (n > 0) numeric[n - 1] = f;
      }
    });
    
    EPODP_STEP_INFO.forEach((s, i) => {
      if (!steps[s.k]) {
        const fb = numeric[i];
        if (fb) steps[s.k] = fb;
      }
    });

    const stepInfo = [];
    for (const s of EPODP_STEP_INFO) {
      const f = steps[s.k] || null;
      let fotoPath = null;
      if (f && gambarBasePath) {
        const cand = path.join(gambarBasePath, f.file_name);
        if (fs.existsSync(cand)) fotoPath = await seragamkanFoto(cand, 210, 160); // letterbox seragam
      }
      stepInfo.push({
        step: s.k,
        label: s.label,
        file_name: f ? f.file_name : null,
        file_id:   f ? f.id : null,
        adaFoto:   !!fotoPath,
        adaKosong: !fotoPath,
        foto_path: fotoPath,
      });
    }

    return {
      nama,
      lengkap: stepInfo.filter(x => x.adaFoto).length,
      selesai: stepInfo.every(x => x.adaFoto),
      jumlahFoto: files.length,
      steps: stepInfo,

      // ✅ HEADER info proyek — dipakai template di dalam loop {#epodp_list}
      //    (PROYEK/KONTRAK/SURAT PESANAN/DISTRICT/LOKASI per halaman ODP)
      proyek:       extraInfo.proyek       || '-',
      kontrak:      extraInfo.kontrak      || '-',
      surat_pesan:  extraInfo.surat_pesan  || '-',
      district:     extraInfo.district     || '-',
      lokasi_smile: extraInfo.lokasi_smile || '-',

      // >>> INI KUNCI AGAR FOTO MUNCUL DI DALAM TABEL 3 KOLOM <<<
      ada_sbl: stepInfo[0].adaFoto,
      foto_sbl: stepInfo[0].foto_path,
      kosong_sbl: !stepInfo[0].adaFoto,
      
      ada_prg: stepInfo[1].adaFoto,
      foto_prg: stepInfo[1].foto_path,
      kosong_prg: !stepInfo[1].adaFoto,
      
      ada_sdh: stepInfo[2].adaFoto,
      foto_sdh: stepInfo[2].foto_path,
      kosong_sdh: !stepInfo[2].adaFoto
    };
  }));
}

// ═════════════════════════════════════════════════════════════════════════════
// Helper: susun data Eviden CUT OVER (ECO) untuk LOOP di template Word.
// Dipakai docxtemplater: {#eco_list} ... {#ada_sbl}{%foto_sbl}{/ada_sbl} ...
// Setiap titik cut over = 3 step: SEBELUM (-SBL) → PROGRES (-PRG) → SESUDAH (-SDH).
// Data lama bernomor (ECO-{titik}-1/2/3 atau ECO1/ECO2/ECO3) dipetakan fallback
// ke SBL/PRG/SDH. Foto lama tanpa sub_kategori dikelompokkan sebagai "CUT OVER".
// ═════════════════════════════════════════════════════════════════════════════
async function susunDataEco(lokasi_id, gambarBasePath, extraInfo = {}) {
  const kategori = '-EVIDEN COLUSER';
  const rows = await db.all(
    'SELECT id, sub_kategori, file_name FROM berkas WHERE lokasi_id = ? AND kategori = ? ORDER BY id ASC',
    [lokasi_id, kategori]
  );

  // Nama + jenis closure (24/48/96/dll) per titik cut over
  const closures = await db.all(
    'SELECT nama, jenis FROM co_closure WHERE lokasi_id = ?',
    [lokasi_id]
  );
  const closureMap = new Map();
  closures.forEach(c => {
    const nm = (c.nama || '').trim();
    if (nm) closureMap.set(nm.toLowerCase(), (c.jenis || '').trim());
  });

  const byTitik = {};
  rows.forEach(r => {
    let nm = (r.sub_kategori || '').trim();
    if (!nm) {
      // Fallback: turunkan nama titik dari nama file lama
      // "ECO1.jpg" → "", "ECO-ODP1-1.jpg" → "ODP1"
      nm = (r.file_name || '')
        .replace(/^ECO[-_]?/i, '')
        .replace(/[-_]?\d+(\.\w+)?$/i, '')
        .replace(/\.\w+$/i, '')
        .trim();
    }
    if (!nm) nm = 'CUT OVER';
    if (!byTitik[nm]) byTitik[nm] = [];
    byTitik[nm].push(r);
  });

  return await Promise.all(Object.keys(byTitik).sort((a, b) => a.localeCompare(b)).map(async nama => {
    const files = byTitik[nama];
    const jenis = closureMap.get(nama.toLowerCase()) || '';
    const steps = {};
    const numeric = [];

    files.forEach(f => {
      const m = /-(SBL|PRG|SDH)\.(jpg|jpeg|png|webp|pdf)$/i.exec(f.file_name || '');
      if (m) {
        const k = m[1].toUpperCase();
        if (!steps[k]) steps[k] = f;
      } else {
        // Terima pola ECO1.jpg (tanpa strip), ECO-ODP1-1.jpg, dan ECO-ODP1-P1.jpg
        const nm2 = /[-P]*(\d+)\.(jpg|jpeg|png|webp|pdf)$/i.exec(f.file_name || '');
        const n = nm2 ? parseInt(nm2[1], 10) : 0;
        if (n > 0) numeric[n - 1] = f;
      }
    });

    EPODP_STEP_INFO.forEach((s, i) => {
      if (!steps[s.k]) {
        const fb = numeric[i];
        if (fb) steps[s.k] = fb;
      }
    });

    const stepInfo = [];
    for (const s of EPODP_STEP_INFO) {
      const f = steps[s.k] || null;
      let fotoPath = null;
      if (f && gambarBasePath) {
        const cand = path.join(gambarBasePath, f.file_name);
        if (fs.existsSync(cand)) fotoPath = await seragamkanFoto(cand, 210, 160); // letterbox seragam
      }
      stepInfo.push({
        step: s.k,
        label: s.label,
        file_name: f ? f.file_name : null,
        file_id:   f ? f.id : null,
        adaFoto:   !!fotoPath,
        adaKosong: !fotoPath,
        foto_path: fotoPath,
      });
    }

    return {
      nama,
      jenis,
      adaJenis: !!jenis,
      lengkap: stepInfo.filter(x => x.adaFoto).length,
      selesai: stepInfo.every(x => x.adaFoto),
      jumlahFoto: files.length,
      steps: stepInfo,

      // ✅ HEADER info proyek — dipakai template di dalam loop {#eco_list}
      //    (PROYEK/KONTRAK/SURAT PESANAN/DISTRICT/LOKASI per halaman closure)
      proyek:       extraInfo.proyek       || '-',
      kontrak:      extraInfo.kontrak      || '-',
      surat_pesan:  extraInfo.surat_pesan  || '-',
      district:     extraInfo.district     || '-',
      lokasi_smile: extraInfo.lokasi_smile || '-',

      ada_sbl: stepInfo[0].adaFoto,
      foto_sbl: stepInfo[0].foto_path,
      kosong_sbl: !stepInfo[0].adaFoto,

      ada_prg: stepInfo[1].adaFoto,
      foto_prg: stepInfo[1].foto_path,
      kosong_prg: !stepInfo[1].adaFoto,

      ada_sdh: stepInfo[2].adaFoto,
      foto_sdh: stepInfo[2].foto_path,
      kosong_sdh: !stepInfo[2].adaFoto
    };
  }));
}

// ═══ MASUKAN EVIDEN PEKERJAAN (EEP) — loop per sub_kategori ═══════════════════
// Setiap eviden (nama sub_kategori) menjadi satu item loop: mulai halaman baru,
// di-heading "EVIDEN {nomor} — {nama}", lalu tabel foto ({#foto_baris} 4 per baris).

// ✅ UBAH: tambah parameter extraInfo = {}
async function siapkanDataEep(lokasi_id, gambarBasePath, blankImage, extraInfo = {}) {
  const kategori = 'MASUKAN EVIDEN PEKERJAAN';
  const rows = await db.all(
    'SELECT id, sub_kategori, file_name FROM berkas WHERE lokasi_id = ? AND kategori = ? ORDER BY id ASC',
    [lokasi_id, kategori]
  );

  const byNama = {};
  (rows || []).forEach(r => {
    let nm = (r.sub_kategori || '').trim();
    if (!nm) nm = (r.file_name || '').replace(/\.[^.]+$/i, '').trim();
    if (!nm) nm = 'Eviden';
    if (!byNama[nm]) byNama[nm] = [];
    byNama[nm].push(r);
  });

  return await Promise.all(Object.keys(byNama)
    .sort((a, b) => a.localeCompare(b, 'id'))
    .map(async (nm, idx) => {
      const files = byNama[nm];
      const fotos = await Promise.all((files || [])
        .map(async f => {
          const cand = gambarBasePath ? path.join(gambarBasePath, f.file_name) : '';
          if (!cand || !fs.existsSync(cand)) return null;
          // letterbox seragam: semua foto EEP jadi tepat 160×120 (tanpa gepeng)
          return await seragamkanFoto(cand, 160, 120);
        }))
        .then(a => a.filter(Boolean));

      const foto_baris = [];
      const blank = blankImage ? await seragamkanFoto(blankImage, 160, 120) : blankImage;
      for (let i = 0; i < fotos.length; i += 4) {
        const row = fotos.slice(i, i + 4);
        foto_baris.push({
          f1: row[0] || blank,
          f2: row[1] || blank,
          f3: row[2] || blank,
          f4: row[3] || blank,
        });
      }
      if (foto_baris.length === 0) {
        foto_baris.push({ f1: blank, f2: blank, f3: blank, f4: blank });
      }

      return {
        nomor: idx + 1,
        nama: nm,
        jumlah: fotos.length,
        foto_baris,
        // ✅ TAMBAH: field header info proyek (dipakai template per halaman)
        proyek:       extraInfo.proyek       || '-',
        kontrak:      extraInfo.kontrak      || '-',
        surat_pesan:  extraInfo.surat_pesan  || '-',
        district:     extraInfo.district     || '-',
        lokasi_smile: extraInfo.lokasi_smile || '-',
      };
    }));
}

// ═════════════════════════════════════════════════════════════════════════════
// [API] Monitoring Eviden Pemasangan ODP (EPODP)
// ═════════════════════════════════════════════════════════════════════════════
app.get('/api/epodp/:lokasi_id', async (req, res) => {
  try {
    const lokasi_id = req.params.lokasi_id;
    const lokasi = await db.get('SELECT * FROM lokasi WHERE id = ?', [lokasi_id]);
    if (!lokasi) return res.status(404).json({ message: 'Lokasi tidak ditemukan.' });

    const ntc = (lokasi.tematik || 'Umum').replace(/[^a-zA-Z0-9_-]/g, '_');
    const nlc = lokasi.nama_lokasi.replace(/[^a-zA-Z0-9_-]/g, '_');
    const kategori = '-EVIDEN PEMASANGAN ODP';
    const kc = kategori.replace(/[^a-zA-Z0-9_-]/g, '_');

    const daftarRaw = await susunDataEpodp(lokasi_id, null);
    const daftar = daftarRaw.map(g => {
      const stepInfo = {};
      g.steps.forEach(s => {
        stepInfo[s.step] = s.file_name
          ? {
              ada: true,
              id: s.file_id,
              file: s.file_name,
              url: `/uploads/${ntc}/${nlc}/${kc}/${s.file_name}`,
            }
          : null;
      });
      return {
        nama: g.nama,
        lengkap: g.lengkap,
        selesai: g.selesai,
        steps: stepInfo,
        jumlahFoto: g.jumlahFoto,
      };
    });

    res.json({
      daftar,
      totalOdp: daftar.length,
      totalSelesai: daftar.filter(d => d.selesai).length,
      totalFoto: daftar.reduce((s, d) => s + (d.jumlahFoto || 0), 0),
    });
  } catch (e) {
    console.error(e);
    res.status(500).json({ message: 'Gagal memuat monitoring Pemasangan ODP.' });
  }
});

// ═════════════════════════════════════════════════════════════════════════════
// [API] Monitoring Eviden CUT OVER (ECO)
// ═════════════════════════════════════════════════════════════════════════════
app.get('/api/eco/:lokasi_id', async (req, res) => {
  try {
    const lokasi_id = req.params.lokasi_id;
    const lokasi = await db.get('SELECT * FROM lokasi WHERE id = ?', [lokasi_id]);
    if (!lokasi) return res.status(404).json({ message: 'Lokasi tidak ditemukan.' });

    const ntc = (lokasi.tematik || 'Umum').replace(/[^a-zA-Z0-9_-]/g, '_');
    const nlc = lokasi.nama_lokasi.replace(/[^a-zA-Z0-9_-]/g, '_');
  const kategori = '-EVIDEN COLUSER';
    const kc = kategori.replace(/[^a-zA-Z0-9_-]/g, '_');

    const daftarRaw = await susunDataEco(lokasi_id, null);
    const daftar = daftarRaw.map(g => {
      const stepInfo = {};
      g.steps.forEach(s => {
        stepInfo[s.step] = s.file_name
          ? {
              ada: true,
              id: s.file_id,
              file: s.file_name,
              url: `/uploads/${ntc}/${nlc}/${kc}/${s.file_name}`,
            }
          : null;
      });
      return {
        nama: g.nama,
        jenis: g.jenis,
        adaJenis: g.adaJenis,
        lengkap: g.lengkap,
        selesai: g.selesai,
        steps: stepInfo,
        jumlahFoto: g.jumlahFoto,
      };
    });

    res.json({
      daftar,
      totalTitik: daftar.length,
      totalSelesai: daftar.filter(d => d.selesai).length,
      totalFoto: daftar.reduce((s, d) => s + (d.jumlahFoto || 0), 0),
    });
  } catch (e) {
    console.error(e);
    res.status(500).json({ message: 'Gagal memuat monitoring CUT OVER.' });
  }
});

// Simpan / perbarui nama + jenis closure cut over (dipanggil dari wizard user)
app.post('/api/eco/closure', checkAuth(), async (req, res) => {
  try {
    const { lokasi_id, nama, jenis } = req.body;
    if (!lokasi_id || !nama) {
      return res.status(400).json({ message: 'Lokasi dan nama closure wajib diisi.' });
    }
    const namaBersih = String(nama).trim();
    if (!namaBersih) return res.status(400).json({ message: 'Nama closure wajib diisi.' });
    const jenisBersih = String(jenis || '').trim();

    await db.run(
      `INSERT INTO co_closure (lokasi_id, nama, jenis) VALUES (?, ?, ?)
       ON CONFLICT(lokasi_id, nama) DO UPDATE SET jenis = excluded.jenis`,
      [lokasi_id, namaBersih, jenisBersih]
    );
    res.json({ success: true, nama: namaBersih, jenis: jenisBersih });
  } catch (e) {
    console.error(e);
    res.status(500).json({ message: 'Gagal menyimpan data closure.', detail: String(e && e.message || e) });
  }
});

// Hapus satu ODP secara total (data port, database berkas, dan file fisik)
app.delete('/api/opm/:lokasi_id/:odp', checkAuth('admin'), async (req, res) => {
  try {
    const { lokasi_id, odp } = req.params;
    const lokasi = await db.get('SELECT * FROM lokasi WHERE id = ?', [lokasi_id]);
    if (!lokasi) return res.status(404).json({ message: 'Lokasi tidak ditemukan.' });

    const ntc = (lokasi.tematik || 'Umum').replace(/[^a-zA-Z0-9_-]/g, '_');
    const nlc = lokasi.nama_lokasi.replace(/[^a-zA-Z0-9_-]/g, '_');
    const kategori = '-EVIDENT HASIL UKUR OPM';
    const kc = kategori.replace(/[^a-zA-Z0-9_-]/g, '_');

    // 1. Ambil daftar file yang akan dihapus dari disk
    const files = await db.all(
      'SELECT file_name FROM berkas WHERE lokasi_id = ? AND kategori = ? AND sub_kategori = ?',
      [lokasi_id, kategori, odp]
    );

    // 2. Hapus file fisik
    const baseDir = path.join(__dirname, 'uploads', ntc, nlc, kc);
    files.forEach(f => {
      const fp = path.join(baseDir, f.file_name);
      if (fs.existsSync(fp)) fs.unlinkSync(fp);
    });

    // 3. Hapus dari database
    await db.run('DELETE FROM berkas   WHERE lokasi_id = ? AND kategori = ? AND sub_kategori = ?', [lokasi_id, kategori, odp]);
    await db.run('DELETE FROM opm_port WHERE lokasi_id = ? AND odp = ?', [lokasi_id, odp]);
    await db.run('DELETE FROM opm_odp  WHERE lokasi_id = ? AND odp = ?', [lokasi_id, odp]);

    res.json({ success: true, message: `ODP ${odp} berhasil dihapus.` });
  } catch (e) {
    console.error(e);
    res.status(500).json({ message: 'Gagal menghapus ODP.', detail: e.message });
  }
});

app.get('/api/get-lokasi', async (req, res) => {
  try {
    const lokasi = await db.all('SELECT id, nama_lokasi, tematik, pid FROM lokasi ORDER BY id ASC');
    res.json(lokasi);
  } catch (e) {
    console.error('Error saat ambil lokasi:', e);
    res.status(500).json({ message: 'Gagal ambil data lokasi' });
  }
});

// ═════════════════════════════════════════════════════════════════════════════
// FOTO SERAGAM (LETTERBOX) — semua foto grid dijepit ke kotak TETAP dengan
// latar putih, tanpa gepeng (rasio asli dipertahankan) dan tanpa crop.
// Hasil cache di %TEMP%/foto_seragam_webbogor (otomatis hangus saat PC reboot).
// Dipakai: grid EEP (f1..f4), 3-step ODP/closure (sbl/prg/sdh), grid OPM
// (foto1..16), foto per-port (foto_path), dan slot EDIST.
// ═════════════════════════════════════════════════════════════════════════════
const DIR_FOTO_SERAGAM = path.join(os.tmpdir(), 'foto_seragam_webbogor');
// Skala supersampling: letterbox dirender pada kelipatan ini lalu ditampilkan di
// dokumen pada ukuran kotak asli (w×h). Rasio & tampilan identik, tetapi jumlah
// piksel jauh lebih banyak → hasil cetak/zoom TIDAK PECAH (tajam). 4× ≈ 380 DPI
// pada kotak ~2,2 inci, tetap tajam saat PDF di-zoom.
const FOTO_SERAGAM_SCALE = 4;
const FOTO_SERAGAM_CACHE_V = 'v4'; // naikkan bila algoritma berubah → cache lama hangus
async function seragamkanFoto(srcPath, w, h) {
  try {
    if (!srcPath || !fs.existsSync(srcPath)) return srcPath;
    if (!/\.(jpe?g|png|webp)$/i.test(srcPath)) return srcPath; // pdf dkk dibiarkan
    fs.mkdirSync(DIR_FOTO_SERAGAM, { recursive: true });
    const st = fs.statSync(srcPath);
    const key = crypto.createHash('md5')
      .update(`${srcPath}|${st.mtimeMs}|${st.size}|${w}x${h}|${FOTO_SERAGAM_CACHE_V}`)
      .digest('hex');
    const out = path.join(DIR_FOTO_SERAGAM, `${key}.jpg`);
    if (fs.existsSync(out)) return out;
    const S = FOTO_SERAGAM_SCALE;
    await sharp(srcPath)
      // resize ke kotak supersampled; withoutEnlargement agar foto kecil tidak
      // di-upscale (tetap dijaga rasio & latar putih letterbox).
      .resize(w * S, h * S, { fit: 'contain', background: { r: 255, g: 255, b: 255 }, withoutEnlargement: false })
      .flatten({ background: '#ffffff' })
      // kualitas tinggi + warna penuh (4:4:4) supaya detail teks/label terbaca
      .jpeg({ quality: 95, chromaSubsampling: '4:4:4' })
      .toFile(out);
    return out;
  } catch (e) {
    console.warn('seragamkanFoto gagal (' + srcPath + '):', e.message);
    return srcPath; // fallback: pakai file asli
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// Helper: susun data OPM untuk LOOOP di template Word.
// Dipakai docxtemplater: {#opm_list} ... {/opm_list}
// Diubah untuk mendukung template 3 Kolom Grid (foto1, status1, foto2, dst)
// ─────────────────────────────────────────────────────────────────────────────
async function siapkanDataOpm(lokasi_id, gambarBasePath, kategori = '-EVIDENT HASIL UKUR OPM', tabel = 'opm', extraInfo = {}) {
  const odps = await db.all(`SELECT odp, kapasitas FROM ${tabel}_odp WHERE lokasi_id = ? ORDER BY odp ASC`, [lokasi_id]);
  const portRows = await db.all(`SELECT odp, port, status FROM ${tabel}_port WHERE lokasi_id = ?`, [lokasi_id]);
  const fotoRows = await db.all(
    'SELECT sub_kategori, file_name FROM berkas WHERE lokasi_id = ? AND kategori = ?',
    [lokasi_id, kategori]
  );

  // Gabungkan ODP dari tabel opm_odp DAN dari sub_kategori di berkas
  const odpMap = new Map();
  odps.forEach(g => odpMap.set(g.odp, g.kapasitas || 8));
  fotoRows.forEach(f => {
    const nm = (f.sub_kategori || '').trim();
    if (nm && !odpMap.has(nm)) odpMap.set(nm, 8);
  });

  return (await Promise.all(Array.from(odpMap.entries()).map(async ([nm, kap]) => {
    const portMap = {};
    portRows.filter(p => p.odp === nm).forEach(p => { portMap[p.port] = p.status; });
    const fotoMap = {};
    fotoRows.filter(f => (f.sub_kategori || '').trim() === nm).forEach(f => {
      const m = /-P?(\d+)\.(jpg|jpeg|png|webp|pdf)$/i.exec(f.file_name || '');
      const p = m ? parseInt(m[1], 10) : null;
      if (p) fotoMap[p] = f.file_name;
    });

    // Kapasitas efektif naik otomatis jika ada foto/status di port yang lebih tinggi
    let efektifKap = kap || 8;
    const maksFoto = Object.keys(fotoMap).map(Number);
    const maksStatus = Object.keys(portMap).map(Number);
    if (maksFoto.length)   efektifKap = Math.max(efektifKap, Math.max(...maksFoto));
    if (maksStatus.length) efektifKap = Math.max(efektifKap, Math.max(...maksStatus));

    // Objek utama untuk satu ODP
    const odpObj = { 
        nama: nm, 
        kapasitas: efektifKap, 
        ports: [], // Tetap biarkan list ports untuk backward-compatibility jika template lama dipakai
        // ✅ HEADER info proyek — dipakai template di dalam loop {#opm_list}
        //    (PROYEK/KONTRAK/SURAT PESANAN/DISTRICT/LOKASI per halaman OPM)
        proyek:       extraInfo.proyek       || '-',
        kontrak:      extraInfo.kontrak      || '-',
        surat_pesan:  extraInfo.surat_pesan  || '-',
        district:     extraInfo.district     || '-',
        lokasi_smile: extraInfo.lokasi_smile || '-',
    };
    // Alias port_list → ports supaya template loop {#port_list} (template
    // HEM/MICRODEMAND/RELOKASI) ikut merender baris per port dengan data yang sama.
    odpObj.port_list = odpObj.ports;

    for (let n = 1; n <= efektifKap; n++) {
      const status = portMap[n] || (fotoMap[n] ? 'ukur' : 'kosong');
      const fotoName = fotoMap[n] || null;
      let fotoPath = null;
      if (fotoName && gambarBasePath) {
        const cand = path.join(gambarBasePath, fotoName);
        if (fs.existsSync(cand)) fotoPath = cand;
      }

      // ── FOTO SERAGAM (letterbox) ──
      // fotoGrid : kotak tetap untuk grid 3 kolom {%foto1..16} (sel 2880 dxa ≈ 192px)
      // fotoList : kotak tetap untuk baris per-port {%foto_path}
      const fotoGrid = fotoPath ? await seragamkanFoto(fotoPath, 175, 135) : null;
      const fotoList = fotoPath ? await seragamkanFoto(fotoPath, 170, 130) : null;

      // --- LOGIKA LAMA (List memanjang ke bawah) ---
      odpObj.ports.push({
        nomor: n,
        tipe: status === 'ukur' ? 'FOTO' : (status === 'isi' ? 'ISI' : 'KOSONG'),
        status,
        adaFoto: status === 'ukur',
        adaIsi: status === 'isi',
        adaKosong: status === 'kosong',
        bukan_foto: status !== 'ukur',
        foto_path: fotoList,
      });

      // --- LOGIKA BARU UNTUK TEMPLATE 3 KOLOM (foto1, status1, dst) ---
      // Masukkan gambar (jika tidak ada foto, di-set null agar Docxtemplater memanggil blank.png)
      odpObj["foto" + n] = (status === 'ukur' && fotoGrid) ? fotoGrid : null;

      // Set teks tulisan statusnya
      if (status === 'ukur') {
          odpObj["status" + n] = ""; // Kosongkan teks jika ada fotonya agar rapi
      } else if (status === 'isi') {
          odpObj["status" + n] = "STATUS : ISI";
      } else {
          odpObj["status" + n] = "STATUS : ISI"; // Tetap ISI untuk port kosong (sesuai input)
      }
    }

    return odpObj;
  }))).sort((a, b) => a.nama.localeCompare(b.nama));
}

// ─────────────────────────────────────────────────────────────────────────────
// PILIHAN "DISMANTLING / TIDAK DISMANTLING" — Laporan Serah Terima Hasil
// Dismantling (Tematik Relokasi). Template memuat dua versi tiap kata:
// yang dipilih tampil normal, yang TIDAK dipilih tampil dicoret (w:strike
// sudah ditambal di templateRelokasi_clean.docx).
//   pilih_dismantling       + pilih_tidak_dismantling  → kata terpilih normal
//   coret_dismantling       + coret_tidak_dismantling  → kata tak terpilih coret
//   dismantling_netral                                      → belum ada pilihan
// ─────────────────────────────────────────────────────────────────────────────
function flagPilihanDismantling(pilihan) {
  const p = String(pilihan || '').trim().toUpperCase();
  const ya    = p === 'DISMANTLING';
  const tidak = p === 'TIDAK DISMANTLING';
  return {
    pilih_dismantling:       ya,
    coret_dismantling:       tidak,
    pilih_tidak_dismantling: tidak,
    coret_tidak_dismantling: ya,
    dismantling_netral:      !ya && !tidak,
  };
}

// ─────────────────────────────────────────────────────────────────────────────
// GRID EVIDEN REDAMAN ODP/OTB — 3 KOTAK KE SAMPING (template HEM & MICRODEMAND)
// Dulu 1 port per baris tabel (memanjang ke bawah); sekarang port digabung
// per 3 menjadi satu baris. Dipakai di template:
//   {#redaman_list} … {#baris} <satu baris 3 sel> {/baris} … {/redaman_list}
// Isi tiap baris (n = 1..3, slot kosong ditiadakan):
//   label{n}   : "PORT 3 · UKUR"            (hilang → paragraf kosong)
//   adaFoto{n} : true + redFoto{n} path      (tag {%redFoto1..3} di dalamnya)
//   kosong{n}  : true                        (tanpa foto → tanda "—")
// ─────────────────────────────────────────────────────────────────────────────
function barisGridRedaman(portList, perBaris) {
  const n = perBaris || 3;
  const baris = [];
  const daftar = Array.isArray(portList) ? portList : [];
  for (let i = 0; i < daftar.length; i += n) {
    const row = {};
    daftar.slice(i, i + n).forEach((pt, idx) => {
      const kol = idx + 1;
      const teks = pt.adaFoto ? 'UKUR' : (pt.adaIsi ? 'ADA USER (ISI)' : 'KOSONG');
      row['label' + kol] = 'PORT ' + pt.nomor + ' · ' + teks;
      if (pt.adaFoto && pt.foto_path) {
        row['adaFoto' + kol] = true;
        row['redFoto' + kol] = pt.foto_path;
      } else {
        row['kosong' + kol] = true;
      }
    });
    baris.push(row);
  }
  return baris;
}

// ─────────────────────────────────────────────────────────────────────────────
// Helper: daftar foto OPM yang diurutkan ODP → port (untuk mengisi slot EOPM1..N
// di template). Slot statis {%EOPM1}.. mendukung gambar pada module gratis,
// sedangkan tag %% di dalam loop TIDAK didukung. Jadi urutan slot EOPM dipakai
// sebagai pengganti "loop gambar".
// ─────────────────────────────────────────────────────────────────────────────
async function fotoEopmUrut(lokasi_id, gambarBasePath) {
  const daftar = await siapkanDataOpm(lokasi_id, gambarBasePath);
  const hasil = [];
  daftar.forEach(odp => {
    (odp.ports || []).forEach(pt => {
      if (pt.adaFoto && pt.foto_path) hasil.push(pt.foto_path);
    });
  });
  return hasil;
}

// Urutkan baris berkas EOPM dari tabel berkas menjadi urutan ODP → port.
function urutBerkasEopm(rows) {
  const parsed = rows.map(r => {
    const m = /-P?(\d+)\.(jpg|jpeg|png|webp|pdf)$/i.exec(r.file_name || '');
    return { r, odp: (r.sub_kategori || '').trim().toLowerCase(), port: m ? parseInt(m[1], 10) : 0 };
  });
  parsed.sort((a, b) => (a.odp === b.odp ? a.port - b.port : a.odp.localeCompare(b.odp)));
  return parsed.map(x => x.r);
}

// ─────────────────────────────────────────────────────────────────────────────
// [API] Progres Berkas
// ─────────────────────────────────────────────────────────────────────────────
app.get('/api/progres/:lokasi_id', async (req, res) => {
  try {
    const databaseTematikWajib = DATABASE_TEMATIK_WAJIB;

    const lokasi = await db.get('SELECT tematik FROM lokasi WHERE id = ?', [req.params.lokasi_id]);
    if (!lokasi) return res.status(404).json({ message: "Lokasi tidak ditemukan." });

    const berkasUploaded = await db.all(
      'SELECT DISTINCT kategori FROM berkas WHERE lokasi_id = ?',
      [req.params.lokasi_id]
    );
    // Kategori teks / tabel khusus yang diisi lewat form input khusus
    const inputUploaded = await db.all(
      'SELECT DISTINCT kategori FROM input_khusus WHERE lokasi_id = ?',
      [req.params.lokasi_id]
    );
    const boqAda = await db.get(
      'SELECT COUNT(*) as total FROM boq_aktual WHERE lokasi_id = ?',
      [req.params.lokasi_id]
    );
    const btAda = await db.get(
      'SELECT COUNT(*) as total FROM barang_tiba WHERE lokasi_id = ?',
      [req.params.lokasi_id]
    );
    const kategoriSudahAda = [
      ...berkasUploaded.map(b => b.kategori),
      ...inputUploaded.map(b => b.kategori),
      ...(boqAda && boqAda.total > 0 ? ['BOQ AKTUAL'] : []),
      ...(btAda && btAda.total > 0 ? ['BOQ BARANG TIBA'] : []),
    ];

    const daftarWajib = databaseTematikWajib[lokasi.tematik] || [];
    const setWajib    = new Set(daftarWajib);

    const sudahWajib = [...new Set(kategoriSudahAda.filter(k => setWajib.has(k)))];
    const belum = daftarWajib.filter(k => !kategoriSudahAda.includes(k));

    const sudahManual = [...new Set(
      kategoriSudahAda
        .filter(k => !setWajib.has(k))
        .map(k => `MANUAL:${k}`)
    )];

    const sudah = [...sudahWajib, ...sudahManual];

    // KATEGORI YANG BOLEH DIISI BERKALI-KALI (tetap muncul di dropdown user
    // walau sudah ada isinya, supaya bisa input ODP/titik berikutnya)
    const KATEGORI_BERULANG = [
      '-EVIDENT HASIL UKUR OPM',
      'INPUT REKAP OPM',
      '-EVIDEN PEMASANGAN ODP',
      '-EVIDEN COLUSER',
      '-LAMPIRAN DATA PENGUKURAN OPM',
      '-Istalasi Per FAT',
      '-Instalasi Pole',
      'EVIDEN BON MATERIAL',
      'EVIDEN INSTALASI ODP/OTB',
      'EVIDEN REDAMAN ODP/OTB',
      'MASUKAN EVIDEN PEKERJAAN'
    ];
    const berulang = [...new Set(kategoriSudahAda.filter(k => KATEGORI_BERULANG.includes(k)))];

    const inputs = await db.all(
      'SELECT id, kategori, judul, volume, isi, created_at, updated_at FROM input_khusus WHERE lokasi_id = ? ORDER BY id DESC',
      [req.params.lokasi_id]
    );

    res.json({ sudah, belum, berulang, inputs });
  } catch (e) {
    console.error(e);
    res.status(500).json({ message: "Gagal menganalisa data progres berkas." });
  }
});

// ─────────────────────────────────────────────────────────────────────────────
// [API] AUDIT SKIP — berkas "tidak berlaku" per lokasi (panel Audit Berkas)
//   GET  /api/audit-skip/:lokasi_id  -> daftar item yang ditiadakan
//   POST /api/audit-skip             -> tandai/batalkan (admin)
// Item yang ditiadakan tidak dihitung sehingga progres bisa mencapai 100%.
// ─────────────────────────────────────────────────────────────────────────────
app.get('/api/audit-skip/:lokasi_id', checkAuth(), async (req, res) => {
  try {
    const rows = await db.all(
      'SELECT item FROM audit_skip WHERE lokasi_id = ? ORDER BY id',
      [req.params.lokasi_id]
    );
    res.json({ items: rows.map(r => r.item) });
  } catch (e) {
    console.error('GET /api/audit-skip:', e.message);
    res.status(500).json({ ok: false, message: 'Gagal memuat daftar berkas tidak berlaku.' });
  }
});

app.post('/api/audit-skip', checkAuth('admin'), async (req, res) => {
  try {
    const b = req.body || {};
    const lokasiId = Number(b.lokasi_id);
    const item = String(b.item || '').trim();
    const skip = b.skip !== false;
    if (!lokasiId) return res.status(400).json({ ok: false, message: 'lokasi_id wajib diisi.' });
    if (!item) return res.status(400).json({ ok: false, message: 'item wajib diisi.' });
    const lok = await db.get('SELECT id FROM lokasi WHERE id = ?', [lokasiId]);
    if (!lok) return res.status(404).json({ ok: false, message: 'Lokasi tidak ditemukan.' });
    if (skip) {
      await db.run(
        'INSERT OR IGNORE INTO audit_skip (lokasi_id, item, skipped_by) VALUES (?, ?, ?)',
        [lokasiId, item, req.session.user.username]
      );
    } else {
      await db.run('DELETE FROM audit_skip WHERE lokasi_id = ? AND item = ?', [lokasiId, item]);
    }
    res.json({ ok: true, lokasi_id: lokasiId, item: item, skip: !!skip });
  } catch (e) {
    console.error('POST /api/audit-skip:', e.message);
    res.status(500).json({ ok: false, message: 'Gagal menyimpan status tidak berlaku.', detail: e.message });
  }
});

// ─────────────────────────────────────────────────────────────────────────────
// [API] Verifikasi Keaslian Eviden — 4 lapisan (metadata, ELA, AI, skor akhir)
// ─────────────────────────────────────────────────────────────────────────────
app.get('/api/eviden/:fotoId/verifikasi', checkAuth(), async (req, res) => {
  try {
    const fotoId = Number(req.params.fotoId);
    const d = await verif.detailVerifikasi(db, fotoId);
    if (!d) return res.status(404).json({ ok: false, message: 'Foto ini belum diverifikasi.' });
    res.json(d);
  } catch (e) {
    console.error('GET /api/eviden/:fotoId/verifikasi:', e.message);
    res.status(500).json({ ok: false, message: 'Gagal memuat detail verifikasi.', detail: e.message });
  }
});

app.get('/api/eviden/verify-progress/:batchId', checkAuth(), async (req, res) => {
  const p = verif.progresBatch(req.params.batchId);
  if (!p) return res.status(404).json({ ok: false, message: 'Batch verifikasi tidak ditemukan.' });
  res.json(p);
});

app.post('/api/eviden/:fotoId/verify', checkAuth('admin'), async (req, res) => {
  try {
    const fotoId = Number(req.params.fotoId);
    if (!fotoId) return res.status(400).json({ ok: false, message: 'ID foto tidak valid.' });
    const d = await verif.verifikasiFoto(db, fotoId, {
      user: req.session.user.username,
      force: !!(req.body || {}).force
    });
    res.json({ ok: true, verifikasi: d });
  } catch (e) {
    if (e.code === 'NOT_FOUND') return res.status(404).json({ ok: false, message: e.message });
    if (e.code === 'BUKAN_GAMBAR' || e.code === 'FILE_HILANG') {
      return res.status(422).json({ ok: false, message: e.message });
    }
    console.error('POST /api/eviden/:fotoId/verify:', e.message);
    res.status(500).json({ ok: false, message: 'Verifikasi foto gagal dijalankan.', detail: e.message });
  }
});

app.get('/api/lokasi/:lokasiId/verifikasi-summary', checkAuth(), async (req, res) => {
  try {
    const lokasiId = Number(req.params.lokasiId);
    if (!lokasiId) return res.status(400).json({ ok: false, message: 'ID lokasi tidak valid.' });
    const s = await verif.summary(db, lokasiId);
    res.json(s);
  } catch (e) {
    console.error('GET /api/lokasi/:lokasiId/verifikasi-summary:', e.message);
    res.status(500).json({ ok: false, message: 'Gagal memuat ringkasan verifikasi.', detail: e.message });
  }
});

app.post('/api/lokasi/:lokasiId/verify', checkAuth('admin'), async (req, res) => {
  try {
    const lokasiId = Number(req.params.lokasiId);
    if (!lokasiId) return res.status(400).json({ ok: false, message: 'ID lokasi tidak valid.' });
    const b = req.body || {};
    const r = await verif.verifikasiBatch(db, lokasiId, {
      checklist: b.checklist ? String(b.checklist) : null,
      force: !!b.force,
      user: req.session.user.username
    });
    res.json({ ok: true, batchId: r.batchId, total: r.total });
  } catch (e) {
    if (e.code === 'NOT_FOUND') return res.status(404).json({ ok: false, message: e.message });
    console.error('POST /api/lokasi/:lokasiId/verify:', e.message);
    res.status(500).json({ ok: false, message: 'Batch verifikasi gagal dijalankan.', detail: e.message });
  }
});

app.patch('/api/eviden/:fotoId/override', checkAuth('admin'), async (req, res) => {
  try {
    const fotoId = Number(req.params.fotoId);
    if (!fotoId) return res.status(400).json({ ok: false, message: 'ID foto tidak valid.' });
    const b = req.body || {};
    const d = await verif.override(db, fotoId, {
      status: b.status,
      catatan: b.catatan ? String(b.catatan).slice(0, 500) : null,
      user: req.session.user.username
    });
    res.json({ ok: true, verifikasi: d });
  } catch (e) {
    if (e.code === 'NOT_FOUND') return res.status(404).json({ ok: false, message: e.message });
    if (e.code === 'VALIDASI') return res.status(400).json({ ok: false, message: e.message });
    console.error('PATCH /api/eviden/:fotoId/override:', e.message);
    res.status(500).json({ ok: false, message: 'Gagal menyimpan override verifikasi.', detail: e.message });
  }
});

// ─────────────────────────────────────────────────────────────────────────────
// [API] Input Khusus (BOQ AKTUAL / INPUT RFC / EVIDEN BON MATERIAL)
// ─────────────────────────────────────────────────────────────────────────────
app.get('/api/input-khusus/:lokasi_id/:kategori', checkAuth(), async (req, res) => {
  try {
    const rows = await db.all(
      'SELECT * FROM input_khusus WHERE lokasi_id = ? AND kategori = ? ORDER BY id DESC',
      [req.params.lokasi_id, req.params.kategori]
    );
    res.json(rows);
  } catch (e) {
    console.error(e);
    res.status(500).json({ message: 'Gagal memuat data input khusus.' });
  }
});

app.post('/api/input-khusus', checkAuth(), async (req, res) => {
  try {
    const { lokasi_id, kategori, judul, volume, isi } = req.body;
    if (!lokasi_id || !kategori) {
      return res.status(400).json({ message: 'Lokasi dan kategori wajib diisi.' });
    }
    await db.run(
      'INSERT INTO input_khusus (lokasi_id, kategori, judul, volume, isi) VALUES (?, ?, ?, ?, ?)',
      [lokasi_id, kategori, judul || null, volume || null, isi || null]
    );

    await autoSetDoneIfComplete(lokasi_id);

    res.json({ success: true });
  } catch (e) {
    console.error(e);
    res.status(500).json({ message: 'Gagal menyimpan input khusus.' });
  }
});

app.delete('/api/input-khusus/:id', checkAuth(), async (req, res) => {
  try {
    await db.run('DELETE FROM input_khusus WHERE id = ?', [req.params.id]);
    res.json({ success: true });
  } catch (e) {
    console.error(e);
    res.status(500).json({ message: 'Gagal menghapus input khusus.' });
  }
});

// ─────────────────────────────────────────────────────────────────────────────
// [API] UBAH DATA INPUT KHUSUS (BOQ / RFC / BON MATERIAL) — edit & simpan ulang
// ─────────────────────────────────────────────────────────────────────────────
app.put('/api/input-khusus/:id', checkAuth(), async (req, res) => {
  try {
    const { judul, volume, isi } = req.body;
    const lama = await db.get(
      'SELECT id, lokasi_id, kategori, judul, volume, isi FROM input_khusus WHERE id = ?',
      [req.params.id]
    );
    if (!lama) {
      return res.status(404).json({ success: false, message: 'Data tidak ditemukan — mungkin sudah dihapus.' });
    }
    // Hanya kolom yang dikirim ikut diubah; sisanya dipertahankan apa adanya.
    await db.run(
      'UPDATE input_khusus SET judul = ?, volume = ?, isi = ?, updated_at = datetime(\'now\',\'localtime\') WHERE id = ?',
      [
        judul !== undefined ? (judul || null) : lama.judul,
        volume !== undefined ? (volume || null) : lama.volume,
        isi !== undefined ? (isi || null) : lama.isi,
        req.params.id
      ]
    );
    await autoSetDoneIfComplete(lama.lokasi_id);
    const baru = await db.get('SELECT * FROM input_khusus WHERE id = ?', [req.params.id]);
    res.json({ success: true, message: 'Perubahan tersimpan.', data: baru });
  } catch (e) {
    console.error(e);
    res.status(500).json({ success: false, message: 'Gagal menyimpan perubahan input khusus.' });
  }
});

// ─────────────────────────────────────────────────────────────────────────────
// [API] Galeri Foto & Hapus Foto
// ─────────────────────────────────────────────────────────────────────────────
app.get('/api/galeri/:lokasi_id', checkAuth(), async (req, res) => {
  try {
    const lokasi = await db.get('SELECT * FROM lokasi WHERE id = ?', [req.params.lokasi_id]);
    if (!lokasi) return res.status(404).json({ message: 'Lokasi tidak ditemukan' });

    const ntc = (lokasi.tematik || 'Umum').replace(/[^a-zA-Z0-9_-]/g, '_');
    const nlc = lokasi.nama_lokasi.replace(/[^a-zA-Z0-9_-]/g, '_');

    const berkas = await db.all(
      `SELECT b.*, v.status AS vrf_status, v.skor AS vrf_skor, v.override_status AS vrf_override,
              v.sha256 AS vrf_sha, v.flags_json AS vrf_flags, v.alasan AS vrf_alasan,
              v.exif_json AS vrf_exif, v.ela_json AS vrf_ela, v.ai_json AS vrf_ai,
              v.override_catatan AS vrf_override_catatan
       FROM berkas b
       LEFT JOIN verifikasi_eviden v ON v.foto_id = b.id
       WHERE b.lokasi_id = ? ORDER BY b.id DESC`,
      [req.params.lokasi_id]
    );
    
    // Format data agar frontend mendapatkan link URL gambar
    const hasil = berkas.map(b => {
      const kc = b.kategori.replace(/[^a-zA-Z0-9_-]/g, '_');
      return {
        id: b.id,
        kategori: b.kategori,
        sub_kategori: b.sub_kategori,
        file_name: b.file_name,
        url: `/uploads/${ntc}/${nlc}/${kc}/${b.file_name}`,
        kml_asli: KATEGORI_KML_FULL.has(b.kategori)
          && !!cariKmlAsli(path.join(__dirname, 'uploads', ntc, nlc, kc), baseTanpaExt(b.file_name)),
        verifikasi: b.vrf_status
          ? {
              status: b.vrf_status,
              skor: b.vrf_skor,
              efektif: b.vrf_override || b.vrf_status,
              ringkas: verif.ringkasanVerifikasi({
                sha256: b.vrf_sha,
                status: b.vrf_status,
                skor: b.vrf_skor,
                override_status: b.vrf_override,
                override_catatan: b.vrf_override_catatan,
                checklist: b.kategori,
                flags_json: b.vrf_flags,
                alasan: b.vrf_alasan,
                exif_json: b.vrf_exif,
                ela_json: b.vrf_ela,
                ai_json: b.vrf_ai
              })
            }
          : null
      };
    });
    res.json(hasil);
  } catch (e) {
    console.error(e);
    res.status(500).json({ message: 'Gagal memuat galeri.' });
  }
});

app.delete('/api/berkas/:id', checkAuth('admin'), async (req, res) => {
  try {
    const fileId = req.params.id;
    const b = await db.get('SELECT * FROM berkas WHERE id = ?', [fileId]);
    if (!b) return res.status(404).json({ message: 'File tidak ditemukan di database.' });

    const lokasi = await db.get('SELECT * FROM lokasi WHERE id = ?', [b.lokasi_id]);
    if (lokasi) {
      const ntc = (lokasi.tematik || 'Umum').replace(/[^a-zA-Z0-9_-]/g, '_');
      const nlc = lokasi.nama_lokasi.replace(/[^a-zA-Z0-9_-]/g, '_');
      const kc  = b.kategori.replace(/[^a-zA-Z0-9_-]/g, '_');
      const filePath = path.join(__dirname, 'uploads', ntc, nlc, kc, b.file_name);

      // Hapus file fisik dari penyimpanan server
      if (fs.existsSync(filePath)) {
        fs.unlinkSync(filePath);
      }
      // Pasangan KML/KMZ asli (untuk Google Earth) ikut dihapus
      const asliPath = cariKmlAsli(path.join(__dirname, 'uploads', ntc, nlc, kc), baseTanpaExt(b.file_name));
      if (asliPath) { try { fs.unlinkSync(asliPath); } catch (err) { /* abaikan */ } }
    }
    // Hapus dari database
    await db.run('DELETE FROM berkas WHERE id = ?', [fileId]);
    await db.run('DELETE FROM verifikasi_eviden WHERE foto_id = ?', [fileId]);
    verif.lupakanHash(fileId);
    res.json({ success: true, message: 'Gambar berhasil dihapus.' });
  } catch (e) {
    console.error(e);
    res.status(500).json({ message: 'Gagal menghapus gambar.' });
  }
});

// ==================== PASTE KODE INI ====================
// ─────────────────────────────────────────────────────────────────────────────
// [API] Galeri Foto & Hapus Foto
// ─────────────────────────────────────────────────────────────────────────────
app.get('/api/galeri/:lokasi_id', checkAuth(), async (req, res) => {
  try {
    const lokasi = await db.get('SELECT * FROM lokasi WHERE id = ?', [req.params.lokasi_id]);
    if (!lokasi) return res.status(404).json({ message: 'Lokasi tidak ditemukan' });

    const ntc = (lokasi.tematik || 'Umum').replace(/[^a-zA-Z0-9_-]/g, '_');
    const nlc = lokasi.nama_lokasi.replace(/[^a-zA-Z0-9_-]/g, '_');

    const berkas = await db.all(
      `SELECT b.*, v.status AS vrf_status, v.skor AS vrf_skor, v.override_status AS vrf_override,
              v.sha256 AS vrf_sha, v.flags_json AS vrf_flags, v.alasan AS vrf_alasan,
              v.exif_json AS vrf_exif, v.ela_json AS vrf_ela, v.ai_json AS vrf_ai,
              v.override_catatan AS vrf_override_catatan
       FROM berkas b
       LEFT JOIN verifikasi_eviden v ON v.foto_id = b.id
       WHERE b.lokasi_id = ? ORDER BY b.id DESC`,
      [req.params.lokasi_id]
    );
    
    // Format data agar frontend mendapatkan link URL gambar
    const hasil = berkas.map(b => {
      const kc = b.kategori.replace(/[^a-zA-Z0-9_-]/g, '_');
      return {
        id: b.id,
        kategori: b.kategori,
        sub_kategori: b.sub_kategori,
        file_name: b.file_name,
        url: `/uploads/${ntc}/${nlc}/${kc}/${b.file_name}`,
        kml_asli: KATEGORI_KML_FULL.has(b.kategori)
          && !!cariKmlAsli(path.join(__dirname, 'uploads', ntc, nlc, kc), baseTanpaExt(b.file_name)),
        verifikasi: b.vrf_status
          ? {
              status: b.vrf_status,
              skor: b.vrf_skor,
              efektif: b.vrf_override || b.vrf_status,
              ringkas: verif.ringkasanVerifikasi({
                sha256: b.vrf_sha,
                status: b.vrf_status,
                skor: b.vrf_skor,
                override_status: b.vrf_override,
                override_catatan: b.vrf_override_catatan,
                checklist: b.kategori,
                flags_json: b.vrf_flags,
                alasan: b.vrf_alasan,
                exif_json: b.vrf_exif,
                ela_json: b.vrf_ela,
                ai_json: b.vrf_ai
              })
            }
          : null
      };
    });
    res.json(hasil);
  } catch (e) {
    console.error(e);
    res.status(500).json({ message: 'Gagal memuat galeri.' });
  }
});

app.delete('/api/berkas/:id', checkAuth('admin'), async (req, res) => {
  try {
    const fileId = req.params.id;
    const b = await db.get('SELECT * FROM berkas WHERE id = ?', [fileId]);
    if (!b) return res.status(404).json({ message: 'File tidak ditemukan di database.' });

    const lokasi = await db.get('SELECT * FROM lokasi WHERE id = ?', [b.lokasi_id]);
    if (lokasi) {
      const ntc = (lokasi.tematik || 'Umum').replace(/[^a-zA-Z0-9_-]/g, '_');
      const nlc = lokasi.nama_lokasi.replace(/[^a-zA-Z0-9_-]/g, '_');
      const kc  = b.kategori.replace(/[^a-zA-Z0-9_-]/g, '_');
      const filePath = path.join(__dirname, 'uploads', ntc, nlc, kc, b.file_name);

      // Hapus file fisik dari penyimpanan server
      if (fs.existsSync(filePath)) {
        fs.unlinkSync(filePath);
      }
      // Pasangan KML/KMZ asli (untuk Google Earth) ikut dihapus
      const asliPath = cariKmlAsli(path.join(__dirname, 'uploads', ntc, nlc, kc), baseTanpaExt(b.file_name));
      if (asliPath) { try { fs.unlinkSync(asliPath); } catch (err) { /* abaikan */ } }
    }
    // Hapus dari database
    await db.run('DELETE FROM berkas WHERE id = ?', [fileId]);
    await db.run('DELETE FROM verifikasi_eviden WHERE foto_id = ?', [fileId]);
    verif.lupakanHash(fileId);
    res.json({ success: true, message: 'Gambar berhasil dihapus.' });
  } catch (e) {
    console.error(e);
    res.status(500).json({ message: 'Gagal menghapus gambar.' });
  }
});
// ==================== SAMPAI SINI ====================


// ─────────────────────────────────────────────────────────────────────────────
// [API] BA CO — GET & POST (role-agnostik: admin & user)
// ─────────────────────────────────────────────────────────────────────────────
const BA_CO_COLS = ['harico', 'tglco', 'bulanco', 'sambungan1co', 'sambungan2co', 'sambungan3co', 'sambungan4co'];

app.get('/api/ba-co/:lokasi_id', checkAuth(), async (req, res) => {
  try {
    const data = await db.get('SELECT * FROM laporan_rekon WHERE lokasi_id = ?', [req.params.lokasi_id]);
    const out = { lokasi_id: req.params.lokasi_id };
    BA_CO_COLS.forEach(c => out[c] = data ? (data[c] || '') : '');
    res.json(out);
  } catch (e) {
    console.error('Gagal ambil BA CO:', e);
    res.status(500).json({ message: 'Gagal mengambil data BA CO.' });
  }
});

app.post('/api/ba-co/:lokasi_id', checkAuth(), async (req, res) => {
  try {
    const lokasiId = req.params.lokasi_id;
    const d = req.body;
    // harico/tglco/bulanco diisi ADMIN (form berkas); jika tidak dikirim, jangan timpa.
    const get = k => (d[k] != null ? String(d[k]) : null);
    const vals = BA_CO_COLS.map(get);

    const eksis = await db.get('SELECT id FROM laporan_rekon WHERE lokasi_id = ?', [lokasiId]);
    if (eksis) {
      await db.run(
        `UPDATE laporan_rekon SET harico=COALESCE(?, harico), tglco=COALESCE(?, tglco), bulanco=COALESCE(?, bulanco), sambungan1co=?, sambungan2co=?, sambungan3co=?, sambungan4co=? WHERE lokasi_id=?`,
        [...vals, lokasiId]
      );
    } else {
      await db.run(
        `INSERT INTO laporan_rekon (lokasi_id, harico, tglco, bulanco, sambungan1co, sambungan2co, sambungan3co, sambungan4co) VALUES (?,?,?,?,?,?,?,?)`,
        [lokasiId, ...vals.map(v => (v === null ? '' : v))]
      );
    }
    res.json({ success: true, message: 'Data BA CO berhasil disimpan.' });
  } catch (e) {
    console.error('Gagal simpan BA CO:', e);
    res.status(500).json({ success: false, message: 'Gagal menyimpan data BA CO.' });
  }
});


// ─────────────────────────────────────────────────────────────────────────────
// [API] GET Laporan Rekon
// ─────────────────────────────────────────────────────────────────────────────
app.get('/api/laporan-rekon/:lokasi_id', checkAuth('admin'), async (req, res) => {
  try {
    const data = await db.get(
      'SELECT * FROM laporan_rekon WHERE lokasi_id = ?',
      [req.params.lokasi_id]
    );
    res.json(data || {});
  } catch (e) {
    console.error(e);
    res.status(500).json({ message: 'Gagal memuat data laporan rekon.' });
  }
});

app.post('/api/rekon/:lokasi_id', checkAuth('admin'), async (req, res) => {
  const lokasiId = req.params.lokasi_id;
  const { items, sumber_master } = req.body || {}; // items dari rekonResult
  const sumber = String(sumber_master || 'amd4').toLowerCase() === 'amd5' ? 'amd5' : 'amd4';
  try {
    await db.run(`DELETE FROM rab_rekon WHERE lokasi_id = ?`, [lokasiId]);
    const stmt = await db.prepare(
      `INSERT INTO rab_rekon (lokasi_id, designator, drm, rekon, tambah, kurang, satuan, harga_satuan, sumber_master)
       VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)`
    );
    for (const item of items) {
      await stmt.run(lokasiId, item.designator, item.drm, item.rekon, item.tambah, item.kurang, item.satuan, item.harga_satuan || 0, sumber);
    }
    await stmt.finalize();
    res.json({ ok: true, sumber_master: sumber });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// Simpan pemilih master VLOOKUP tanpa mengubah baris rekon.
// Dipanggil saat pengguna mengganti AMD4/AMD5 di toolbar RAB.
app.post('/api/rekon/:lokasi_id/sumber-master', checkAuth('admin'), async (req, res) => {
  try {
    const sumber = String((req.body || {}).sumber_master || 'amd4').toLowerCase() === 'amd5'
      ? 'amd5' : 'amd4';
    await db.run('UPDATE rab_rekon SET sumber_master = ? WHERE lokasi_id = ?',
      [sumber, req.params.lokasi_id]);
    res.json({ ok: true, sumber_master: sumber });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});
// ─────────────────────────────────────────────────────────────────────────────
// [API] SIMPAN DATA KHUSUS BA TAMBAH KURANG
// ─────────────────────────────────────────────────────────────────────────────
app.post('/api/simpan-batk/:lokasi_id', checkAuth('admin'), async (req, res) => {
  try {
    const { lokasi_id } = req.params;
    const d = req.body;
    const eksis = await db.get('SELECT id FROM laporan_rekon WHERE lokasi_id = ?', [lokasi_id]);
    
    if (eksis) {
      await db.run(`UPDATE laporan_rekon SET proyektk=?, kontraktk=?, sptk=?, nosptk=?, haritk=?, tgltk=?, blntk=?, thntk=?, loktk=? WHERE lokasi_id=?`,
        [d.proyektk, d.kontraktk, d.sptk, d.nosptk, d.haritk, d.tgltk, d.blntk, d.thntk, d.loktk, lokasi_id]
      );
    } else {
      await db.run(`INSERT INTO laporan_rekon (lokasi_id, proyektk, kontraktk, sptk, nosptk, haritk, tgltk, blntk, thntk, loktk) VALUES (?,?,?,?,?,?,?,?,?,?)`,
        [lokasi_id, d.proyektk, d.kontraktk, d.sptk, d.nosptk, d.haritk, d.tgltk, d.blntk, d.thntk, d.loktk]
      );
    }
    res.json({ success: true });
  } catch (e) {
    console.error(e);
    res.status(500).json({ success: false, message: 'Gagal menyimpan data.' });
  }
});


// ─────────────────────────────────────────────────────────────────────────────
// [API] Diagnostik foto
// ─────────────────────────────────────────────────────────────────────────────
app.get('/api/debug-foto/:lokasi_id', checkAuth('admin'), async (req, res) => {
  try {
    const lokasiData = await db.get('SELECT * FROM lokasi WHERE id = ?', [req.params.lokasi_id]);
    if (!lokasiData) return res.status(404).json({ message: 'Lokasi tidak ditemukan.' });

    const namaTematikClean = (lokasiData.tematik || 'Umum').replace(/[^a-zA-Z0-9_-]/g, '_');
    const namaLokasiClean  = lokasiData.nama_lokasi.replace(/[^a-zA-Z0-9_-]/g, '_');

    const KATEGORI_MAP = {
      'EPT':   '-EVIDEN PEMASANGAN TIANG',
      'EPEN':  '-EVIDEN PENARIKAN',
      'EPODP': '-EVIDEN PEMASANGAN ODP',
      'EOPM':  '-EVIDENT HASIL UKUR OPM',
      'LDP':   '-LAMPIRAN DATA PENGUKURAN OPM',
      'MCR':   '-MANCORE',
      'GR':    '-GROUNDING',
      'LKDL':  '-LAMPIRAN HASIL UJI KEDALAMAN GALIAN',
      'EGAL':  '-EVIDEN EVIDEN GALIAN DAN DIMENSI HANDHOLE',
      'EDIST': '-EVIDEN DISMANTLING',
      'ECO':   '-EVIDEN COLUSER',
      'LCO':   '-LAMPIRAN CO',
      'ROSET': '-Instalasi Roset',
      'EEP':   'MASUKAN EVIDEN PEKERJAAN',
    };

    const hasil = {};
    for (const [kode, namaKategori] of Object.entries(KATEGORI_MAP)) {
      const katClean   = namaKategori.replace(/[^a-zA-Z0-9_-]/g, '_');
      const folderPath = path.join(__dirname, 'uploads', namaTematikClean, namaLokasiClean, katClean);
      const exists     = fs.existsSync(folderPath);
      const files      = exists
        ? fs.readdirSync(folderPath).filter(f => /\.(jpg|jpeg|png|gif|webp)$/i.test(f)).sort()
        : [];
      hasil[kode] = {
        tag_template: `{%${kode}1} … {%${kode}${files.length || 'N'}}`,
        kategori_db:  namaKategori,
        folder:       folderPath,
        folder_exists: exists,
        jumlah_foto:  files.length,
        files,
      };
    }

    res.json({
      lokasi:   lokasiData.nama_lokasi,
      tematik:  lokasiData.tematik,
      base_dir: path.join(__dirname, 'uploads', namaTematikClean, namaLokasiClean),
      kategori: hasil,
    });
  } catch (e) {
    console.error(e);
    res.status(500).json({ message: 'Gagal menjalankan diagnostik.' });
  }
});

// ─────────────────────────────────────────────────────────────────────────────
// [API] POST Laporan Rekon
// ─────────────────────────────────────────────────────────────────────────────
app.post('/api/laporan-rekon/:lokasi_id', checkAuth('admin'), async (req, res) => {
  try {
    const lokasiId = req.params.lokasi_id;
    const d        = req.body;

    const proyek             = d.proyek             || '-';
    const contracts          = d.kontrak            || '-';
    const nomorSuratPesanan  = d.surat_pesanan || d.surat_pesan || '-';
    const district           = d.district           || '-';
    const pelaksana          = d.pelaksana          || '-';
    const hari_ct            = d.hari_ct            || '-';
    const tgl                = d.tgl                || '-';
    const bulan              = d.bulan              || '-';
    const nama_waspang       = d.nama_waspang       || '-';
    const nik_waspang        = d.nik_waspang        || '-';
    const lokasi_smile       = d.lokasi_smile       || '-';
    const ct_hari            = d.ct_hari            || '-';
    const ct_tgl             = d.ct_tgl             || '-';
    const ct_bulan           = d.ct_bulan           || '-';
    const dismantling_hari    = d.dismantling_hari    || '-';
    const dismantling_tanggal = d.dismantling_tanggal || '-';
    const dismantling_bulan   = d.dismantling_bulan   || '-';

    // ── LAPORAN SERAH TERIMA HASIL DISMANTLING (TEMATIK RELOKASI) ──
    // gudang_sto mengisi {STO}; pilih_dismantling menentukan mana yang dicoret.
    // Kosong dibiarkan kosong (tampil netral), jangan diganti '-'.
    const gudang_sto        = (d.gudang_sto        || '').trim();
    const pilih_dismantling = (d.pilih_dismantling || '').trim();
    const ibooster_hari       = d.ibooster_hari       || '-';
    const ibooster_tanggal    = d.ibooster_tanggal    || '-';
    const ibooster_bulan      = d.ibooster_bulan      || '-';
    const ibooster_lokasi     = d.ibooster_lokasi || d.ibooster_loksai || '-';
    const harico              = d.harico || '-';
    const tglco               = d.tglco || '-';
    const bulanco             = d.bulanco || '-';

    // ── INPUT TEMATIK NODE B (LACT) ──
    const waspang_ta               = d.waspang_ta               || '-';
    const nik_waspang_ta           = d.nik_waspang_ta           || '-';
    const tgl_bulan_thn_ct         = d.tgl_bulan_thn_ct         || '-';
    const hari_barang_tiba         = d.hari_barang_tiba         || '-';
    const tgl_barang_tiba          = d.tgl_barang_tiba          || '-';
    const bulan_barang_tiba        = d.bulan_barang_tiba        || '-';
    const tgl_bulan_thn_barang_tiba = d.tgl_bulan_thn_barang_tiba || '-';
    const tgl_bulan_thn_opm        = d.tgl_bulan_thn_opm        || '-';
    const pm_area                  = d.pm_area                  || '-';
    const site                     = d.site                     || '-';
    const ruas_pekerjaan           = d.ruas_pekerjaan           || '-';
    const izin                     = d.izin                     || '-';

    const eksis = await db.get('SELECT id FROM laporan_rekon WHERE lokasi_id = ?', [lokasiId]);

    if (eksis) {
      await db.run(`
        UPDATE laporan_rekon SET
          proyek=?, kontrak=?, surat_pesanan=?, district=?, pelaksana=?,
          hari_ct=?, tgl=?, bulan=?, nama_waspang=?, nik_waspang=?, lokasi_smile=?,
          ct_hari=?, ct_tgl=?, ct_bulan=?,
          dismantling_hari=?, dismantling_tanggal=?, dismantling_bulan=?,
          gudang_sto=?, pilih_dismantling=?,
          matrial_aset1=?, volume_aset1=?, kondisi_aset1=?,
          matrial_aset2=?, volume_aset2=?, kondisi_aset2=?,
          matrial_aset3=?, volume_aset3=?, kondisi_aset3=?,
          matrial_aset4=?, volume_aset4=?, kondisi_aset4=?,
          matrial_aset5=?, volume_aset5=?, kondisi_aset5=?,
          ibooster_hari=?, ibooster_tanggal=?, ibooster_bulan=?, ibooster_lokasi=?,
          harico=?, tglco=?, bulanco=?,
          waspang_ta=?, nik_waspang_ta=?, tgl_bulan_thn_ct=?,
          hari_barang_tiba=?, tgl_barang_tiba=?, bulan_barang_tiba=?,
          tgl_bulan_thn_barang_tiba=?, tgl_bulan_thn_opm=?,
          pm_area=?, site=?, ruas_pekerjaan=?, izin=?
        WHERE lokasi_id = ?`,
        [
          proyek, contracts, nomorSuratPesanan, district, pelaksana,
          hari_ct, tgl, bulan, nama_waspang, nik_waspang, lokasi_smile,
          ct_hari, ct_tgl, ct_bulan,
          dismantling_hari, dismantling_tanggal, dismantling_bulan,
          gudang_sto, pilih_dismantling,
          d.matrial_aset1 || '-', d.volume_aset1 || '-', d.kondisi_aset1 || '-',
          d.matrial_aset2 || '-', d.volume_aset2 || '-', d.kondisi_aset2 || '-',
          d.matrial_aset3 || '-', d.volume_aset3 || '-', d.kondisi_aset3 || '-',
          d.matrial_aset4 || '-', d.volume_aset4 || '-', d.kondisi_aset4 || '-',
          d.matrial_aset5 || '-', d.volume_aset5 || '-', d.kondisi_aset5 || '-',
          ibooster_hari, ibooster_tanggal, ibooster_bulan, ibooster_lokasi,
          harico, tglco, bulanco,
          waspang_ta, nik_waspang_ta, tgl_bulan_thn_ct,
          hari_barang_tiba, tgl_barang_tiba, bulan_barang_tiba,
          tgl_bulan_thn_barang_tiba, tgl_bulan_thn_opm,
          pm_area, site, ruas_pekerjaan, izin,
          lokasiId
        ]
      );
    } else {
      const kolomBaru = [
        'waspang_ta', 'nik_waspang_ta', 'tgl_bulan_thn_ct',
        'hari_barang_tiba', 'tgl_barang_tiba', 'bulan_barang_tiba',
        'tgl_bulan_thn_barang_tiba', 'tgl_bulan_thn_opm',
        'pm_area', 'site', 'ruas_pekerjaan', 'izin'
      ];
      await db.run(`
        INSERT INTO laporan_rekon (
          lokasi_id, proyek, kontrak, surat_pesanan, district, pelaksana,
          hari_ct, tgl, bulan, nama_waspang, nik_waspang, lokasi_smile,
          ct_hari, ct_tgl, ct_bulan,
          dismantling_hari, dismantling_tanggal, dismantling_bulan,
          gudang_sto, pilih_dismantling,
          matrial_aset1, volume_aset1, kondisi_aset1,
          matrial_aset2, volume_aset2, kondisi_aset2,
          matrial_aset3, volume_aset3, kondisi_aset3,
          matrial_aset4, volume_aset4, kondisi_aset4,
          matrial_aset5, volume_aset5, kondisi_aset5,
          ibooster_hari, ibooster_tanggal, ibooster_bulan, ibooster_lokasi,
          harico, tglco, bulanco,
          ${kolomBaru.join(', ')}
        ) VALUES (${new Array(42 + kolomBaru.length).fill('?').join(',')})`,
        [
          lokasiId, proyek, contracts, nomorSuratPesanan, district, pelaksana,
          hari_ct, tgl, bulan, nama_waspang, nik_waspang, lokasi_smile,
          ct_hari, ct_tgl, ct_bulan,
          dismantling_hari, dismantling_tanggal, dismantling_bulan,
          gudang_sto, pilih_dismantling,
          d.matrial_aset1 || '-', d.volume_aset1 || '-', d.kondisi_aset1 || '-',
          d.matrial_aset2 || '-', d.volume_aset2 || '-', d.kondisi_aset2 || '-',
          d.matrial_aset3 || '-', d.volume_aset3 || '-', d.kondisi_aset3 || '-',
          d.matrial_aset4 || '-', d.volume_aset4 || '-', d.kondisi_aset4 || '-',
          d.matrial_aset5 || '-', d.volume_aset5 || '-', d.kondisi_aset5 || '-',
          ibooster_hari, ibooster_tanggal, ibooster_bulan, ibooster_lokasi,
          harico, tglco, bulanco,
          waspang_ta, nik_waspang_ta, tgl_bulan_thn_ct,
          hari_barang_tiba, tgl_barang_tiba, bulan_barang_tiba,
          tgl_bulan_thn_barang_tiba, tgl_bulan_thn_opm,
          pm_area, site, ruas_pekerjaan, izin
        ]
      );
    }

    res.json({ success: true, message: 'Data Rekon Berhasil Disimpan.' });
  } catch (e) {
    console.error('Kesalahan internal saat menyimpan rekon:', e);
    res.status(500).json({ success: false, message: 'Gagal menyimpan entri laporan rekon ke database.' });
  }
});

// ─────────────────────────────────────────────────────────────────────────────
// [API] Download DOCX
// ─────────────────────────────────────────────────────────────────────────────
app.get('/api/download-docx/:lokasi_id', checkAuth('admin'), async (req, res) => {
    try {
        const lokasiId = req.params.lokasi_id;
        const lokasiData = await db.get('SELECT * FROM lokasi WHERE id = ?', [lokasiId]);
        if (!lokasiData) return res.status(404).json({ message: "Data lokasi tidak ditemukan." });
        await catatAktivitas("microdemand_hem", lokasiId);

        // Mode BAUT-1 Node B: ?baut=1 → pakai template TAMPLATE_BAUT_NODE_B + 5 dokumen admin
        const modeBaut1 = req.query.baut === '1' && lokasiData.tematik === 'Node B';

        const adminInput = await db.get('SELECT * FROM laporan_rekon WHERE lokasi_id = ?', [lokasiId]) || {};

// ── Fallback: ambil no_sp dari tabel lokasi jika kontrak/surat_pesanan kosong ──
if (!adminInput.kontrak || adminInput.kontrak === '-') {
  const spRow = await db.get(
    'SELECT no_sp FROM lokasi WHERE id = ?',
    [lokasiId]
  ).catch(() => null);
  if (spRow && spRow.no_sp) {
    adminInput.surat_pesanan = adminInput.surat_pesanan || spRow.no_sp;
  }
}

const semuaBerkas = await db.all('SELECT file_name, kategori, sub_kategori FROM berkas WHERE lokasi_id = ? ORDER BY id ASC', [lokasiId]);
        const namaTematikClean = (lokasiData.tematik || "Umum").replace(/[^a-zA-Z0-9_-]/g, "_");
        const namaLokasiClean = lokasiData.nama_lokasi.replace(/[^a-zA-Z0-9_-]/g, "_");
    const blankImage = path.join(__dirname, 'public', 'images', 'blank.png');

        const KATEGORI_MAP = {
            "EPT":   "-EVIDEN PEMASANGAN TIANG",
            "EPEN":  "-EVIDEN PENARIKAN",
            "EPODP": "-EVIDEN PEMASANGAN ODP",
            "EOPM":  "-EVIDENT HASIL UKUR OPM",
            "LDP":   "-LAMPIRAN DATA PENGUKURAN OPM",
            "MCR":   "-MANCORE",
            "GR":    "-GROUNDING",
            "LKDL":  "-LAMPIRAN HASIL UJI KEDALAMAN GALIAN",
            "EGAL":  "-EVIDEN EVIDEN GALIAN DAN DIMENSI HANDHOLE",
            "EDIST": "-EVIDEN DISMANTLING",
            "ECO":   "-EVIDEN COLUSER",
            "LCO":   "-LAMPIRAN CO",
            "ROSET": "-Instalasi Roset",
            "EEP":   "MASUKAN EVIDEN PEKERJAAN"
        };

        const MAX_SLOT = {
            "EPT":   9,
            "EPEN":  9,
            "EPODP": 27,
            "EOPM":  45,
            "LDP":   4,
            "MCR":   2,
            "GR":    18,
            "LKDL":  3,
            "EGAL":  18,
            "EDIST": 9,
            "ECO":   45,
            "LCO":   9,
            "ROSET": 9,
            "EEP":   18
        };

        function ambilFotoBySlot(kode, slotKe) {
            const namaKategori = KATEGORI_MAP[kode];
            if (!namaKategori) return blankImage;
            // EOPM: urutkan foto mengikuti ODP → port (agar kotak EOPM1..N sesuai eviden)
            const listFoto = (kode === 'EOPM')
                ? urutBerkasEopm(semuaBerkas.filter(b => b.kategori === namaKategori))
                : semuaBerkas.filter(b => b.kategori === namaKategori);
            const targetBerkas = listFoto[slotKe - 1];
            if (targetBerkas) {
                const katClean = targetBerkas.kategori.replace(/[^a-zA-Z0-9_-]/g, "_");
                const folderPath = path.join(__dirname, 'uploads', namaTematikClean, namaLokasiClean, katClean);
                const fullPath = path.join(folderPath, targetBerkas.file_name);
                return fs.existsSync(fullPath) ? fullPath : blankImage;
            }
            return blankImage;
        }

        const TAG_BESAR = ['LDP', 'MCR', 'LKDL'];

        const imageModule = new ImageModule({
            getImage: (tagValue) => {
                const filePath = (tagValue && fs.existsSync(tagValue)) ? tagValue : blankImage;
                return fs.readFileSync(filePath);
            },
            getSize: (img, tagValue, tagName) => {
        try {
            // 1. Ambil ukuran asli file gambar (lebar & tinggi asli)
            const dimensions = sizeOf(img);
            const srcWidth = dimensions.width;
            const srcHeight = dimensions.height;

            // 2. Tentukan KOTAK BATAS MAKSIMAL di Word Anda (Bounding Box)
            const kode = tagName.trim().replace(/\d+$/, '');
            let isBesar = TAG_BESAR.includes(kode);
            if (kode === 'LDP') { isBesar = true; }
            // Tag generik foto loop template HEM/MICRODEMAND ({%f1}..{%f4})
            const isFotoLoop = kode === 'f';
            // Foto eviden per port (redaman) di template HEM/MICRODEMAND
            const isPortFoto = kode === 'foto_path';
            // Grid redaman 3 kotak ke samping ({%redFoto1..3}) — muat 3 kolom
            const isRedFoto = kode === 'redFoto';
            // Foto KML / Peta Rute halaman penuh — hanya tag {%foto} tanpa angka
            const isKmlFoto = tagName === 'foto';
            // Grid OPM 3 kolom ({%foto1..16}) — tag "foto" + angka
            const isOpmGrid = (kode === 'foto' && tagName !== 'foto');
            // 3-step ODP / Closure ({%foto_sbl}, {%foto_prg}, {%foto_sdh})
            const isStepFoto = ['foto_sbl', 'foto_prg', 'foto_sdh'].includes(kode);
            // Grid dismantling ({%EDIST1..9}) — sel 3 kolom 3260 dxa
            const isEdist = kode === 'EDIST';
            // Foto kategori tunggal (1 foto = 1 halaman, diperbesar) — tag {%foto_penuh}
            const isFotoPenuh = tagName === 'foto_penuh';
            // Form Pengukuran OPM (Node B): gambar tabel diisi manual — tampil lebar penuh
            const isOpmForm = tagName === 'FORM PENGUKURAN OPM';
            // Grafik trace OTDR (.sor) di loop {#sor_list} — landscape lebar
            const isSorTrace = tagName === 'sor_gambar';
            // Halaman dokumen BAUT-1 Node B ({%baut_page}) — hampir penuh kertas A4
            const isBautPage = tagName === 'baut_page';

            // Silakan sesuaikan ukuran kotak maksimal ini sesuai keinginan Anda
            // Catatan: foto grid sudah di-letterbox ke ukuran kotak di payload,
            // jadi maxSize di sini diset SAMA dengan ukuran letterbox agar
            // hasil render tepat seragam (rasio = 1, tanpa upscale/downscale).
            let maxWidth  = isBautPage ? 500 : (isFotoPenuh ? 570 : (isSorTrace ? 560 : ((isKmlFoto || isOpmForm) ? 600 : (isStepFoto ? 210 : (isEdist ? 200 : (isRedFoto ? 170 : (isFotoLoop ? 160 : (isPortFoto ? 170 : (isOpmGrid ? 175 : (isBesar ? 350 : 150))))))))));
            let maxHeight = isBautPage ? 700 : (isFotoPenuh ? 600 : (isSorTrace ? 340 : ((isKmlFoto || isOpmForm) ? 720 : (isStepFoto ? 160 : (isEdist ? 150 : (isRedFoto ? 130 : (isFotoLoop ? 120 : (isPortFoto ? 130 : (isOpmGrid ? 135 : (isBesar ? 280 : 110))))))))));
            if (kode === 'LDP') { maxWidth = 600; maxHeight = 340; }

            // 3. Hitung rasio kecocokan agar pas di dalam kotak tanpa mengubah bentuk asli
            const ratioX = maxWidth / srcWidth;
            const ratioY = maxHeight / srcHeight;
            const targetRatio = Math.min(ratioX, ratioY); // Pilih rasio terkecil agar muat di dalam kotak

            // 4. Hasilkan ukuran final yang proporsional
            const finalWidth = Math.round(srcWidth * targetRatio);
            const finalHeight = Math.round(srcHeight * targetRatio);

            return [finalWidth, finalHeight];

        } catch (e) {
            console.error("Gagal menghitung ukuran gambar otomatis:", e);
            // Fallback ukuran standar jika gambar rusak / tidak terbaca
            return [160, 200]; 
        }
    }
    // ════════════════════════════════════════════════════════════════════════
});

        const baseName = namaTemplateDocx(lokasiData.tematik);
        const fixedTemplate = path.join(__dirname, 'templates', `FIXED_${baseName}`);
        const templatePath = path.join(__dirname, 'templates', baseName);
        const baut1Template = path.join(__dirname, 'templates', 'TAMPLATE_BAUT_NODE_B.docx');
        let finalTemplate;
        if (modeBaut1) {
            finalTemplate = fs.existsSync(baut1Template) ? baut1Template : templatePath;
        } else {
            finalTemplate = fs.existsSync(fixedTemplate)
                ? fixedTemplate
                : (fs.existsSync(templatePath) ? templatePath : path.join(__dirname, 'templates', 'templateRelokasi_clean.docx'));
        }

        const doc = new Docxtemplater(
            new PizZip(fs.readFileSync(finalTemplate, 'binary')),
            {
                modules: [imageModule],
                paragraphLoop: true,
                linebreaks: true,
                parser: smartParser // <--- GUNAKAN PARSER BARU DI SINI
            }
        );

        const payloadData = {
            proyek:        adminInput.proyek        || "-",
            kontrak:       adminInput.kontrak       || "-",
            surat_pesan:   adminInput.surat_pesanan || "-",
            district:      adminInput.district      || "-",
            lokasi_smile:  lokasiData.nama_lokasi   || "-",
            pelaksana:     adminInput.pelaksana     || "-",
            nama_waspang:  adminInput.nama_waspang  || "-",
            nik_waspang:   adminInput.nik_waspang   || "-",
            hari_ct:       adminInput.hari_ct       || "-",
            tgl:           adminInput.tgl           || "-",
            bulan:         adminInput.bulan         || "-",
            ct_hari:       adminInput.ct_hari       || "-",
            ct_tgl:        adminInput.ct_tgl        || "-",
            ct_bulan:      adminInput.ct_bulan      || "-",
            ibooster_hari:     adminInput.ibooster_hari     || "-",
            ibooster_tanggal:  adminInput.ibooster_tanggal  || "-",
            ibooster_bulan:    adminInput.ibooster_bulan    || "-",
            ibooster_lokasi:   adminInput.ibooster_lokasi   || "-",
            dismantling_hari:     adminInput.dismantling_hari     || "-",
            dismantling_tanggal:  adminInput.dismantling_tanggal  || "-",
            dismantling_bulan:    adminInput.dismantling_bulan    || "-",
            // Gudang STO (tag {STO}) + pilihan DISMANTLING / TIDAK DISMANTLING
            STO:               adminInput.gudang_sto || "-",
            gudang_sto:        adminInput.gudang_sto || "-",
            ...flagPilihanDismantling(adminInput.pilih_dismantling),
            matrial_aset1: adminInput.matrial_aset1 || "-",
            volume_aset1:  adminInput.volume_aset1  || "-",
            kondisi_aset1: adminInput.kondisi_aset1 || "-",
            matrial_aset2: adminInput.matrial_aset2 || "-",
            volume_aset2:  adminInput.volume_aset2  || "-",
            kondisi_aset2: adminInput.kondisi_aset2 || "-",
            matrial_aset3: adminInput.matrial_aset3 || "-",
            volume_aset3:  adminInput.volume_aset3  || "-",
            kondisi_aset3: adminInput.kondisi_aset3 || "-",
            matrial_aset4: adminInput.matrial_aset4 || "-",
            volume_aset4:  adminInput.volume_aset4  || "-",
            kondisi_aset4: adminInput.kondisi_aset4 || "-",
            matrial_aset5: adminInput.matrial_aset5 || "-",
            volume_aset5:  adminInput.volume_aset5  || "-",
            kondisi_aset5: adminInput.kondisi_aset5 || "-",
            harico:        adminInput.harico        || "-",
            tglco:         adminInput.tglco         || "-",
            bulanco:       adminInput.bulanco       || "-",
            sambungan1co:  adminInput.sambungan1co  || "",
            sambungan2co:  adminInput.sambungan2co  || "",
            sambungan3co:  adminInput.sambungan3co  || "",
            sambungan4co:  adminInput.sambungan4co  || "",
        };

        Object.entries(MAX_SLOT).forEach(([kode, maxSlot]) => {
            for (let slot = 1; slot <= maxSlot; slot++) {
                payloadData[`${kode}${slot}`] = ambilFotoBySlot(kode, slot);
            }
        });
        // EDIST = grid 3 kolom (sel 3260 dxa ≈ 217px) → letterbox seragam 200×150
        for (let s = 1; s <= MAX_SLOT.EDIST; s++) {
            payloadData[`EDIST${s}`] = await seragamkanFoto(payloadData[`EDIST${s}`], 200, 150);
        }

        // Data LOOP hasil ukur OPM per ODP + per port (foto / ISI / kosong)
        // Template: {#opm_list}{nama} ... {#port_list}{nomor}{tipe}{/port_list}{/opm_list}
        const folderOpm = path.join(__dirname, 'uploads', namaTematikClean, namaLokasiClean, '-EVIDENT_HASIL_UKUR_OPM');
        payloadData.opm_list = await siapkanDataOpm(lokasiId, folderOpm, '-EVIDENT HASIL UKUR OPM', 'opm', {
          proyek:       payloadData.proyek        || '-',
          kontrak:      payloadData.kontrak       || '-',
          surat_pesan:  payloadData.surat_pesan   || '-',
          district:     payloadData.district      || '-',
          lokasi_smile: payloadData.lokasi_smile  || '-',
        });
        const folderEpodp = path.join(__dirname, 'uploads', namaTematikClean, namaLokasiClean, '-EVIDEN_PEMASANGAN_ODP');
        payloadData.epodp_list = await susunDataEpodp(lokasiId, folderEpodp, {
          proyek:       payloadData.proyek        || '-',
          kontrak:      payloadData.kontrak       || '-',
          surat_pesan:  payloadData.surat_pesan   || '-',
          district:     payloadData.district      || '-',
          lokasi_smile: payloadData.lokasi_smile  || '-',
        });
        const folderEco = path.join(__dirname, 'uploads', namaTematikClean, namaLokasiClean, '-EVIDEN_COLUSER');
        payloadData.eco_list = await susunDataEco(lokasiId, folderEco, {
          proyek:       payloadData.proyek        || '-',
          kontrak:      payloadData.kontrak       || '-',
          surat_pesan:  payloadData.surat_pesan   || '-',
          district:     payloadData.district      || '-',
          lokasi_smile: payloadData.lokasi_smile  || '-',
        });
        // MASUKAN EVIDEN PEKERJAAN — loop per eviden (sub_kategori), tiap eviden di halaman sendiri
        const folderEep = path.join(__dirname, 'uploads', namaTematikClean, namaLokasiClean, 'MASUKAN_EVIDEN_PEKERJAAN');
        payloadData.eep_list = await siapkanDataEep(lokasiId, folderEep, blankImage, {
  proyek:       payloadData.proyek        || '-',
  kontrak:      payloadData.kontrak       || '-',
  surat_pesan:  payloadData.surat_pesan   || '-',   // ← bukan nomorSuratPesanan!
  district:     payloadData.district      || '-',
  lokasi_smile: payloadData.lokasi_smile  || '-',
});
        console.log('[EEP] lokasiId=' + lokasiId + ' folder=' + folderEep + ' exists=' + fs.existsSync(folderEep) + ' groups=' + (payloadData.eep_list||[]).length);
        (payloadData.eep_list||[]).forEach(function(g){ console.log('[EEP]   ' + g.nomor + '. ' + g.nama + ' (' + g.jumlah + ' foto, ' + g.foto_baris.length + ' baris)'); });
        const folderRedaman = path.join(__dirname, 'uploads', namaTematikClean, namaLokasiClean, 'EVIDEN_REDAMAN_ODP_OTB');
        payloadData.redaman_list = await siapkanDataOpm(lokasiId, folderRedaman, KATEGORI_REDAMAN, 'redaman', {
          proyek:       payloadData.proyek        || '-',
          kontrak:      payloadData.kontrak       || '-',
          surat_pesan:  payloadData.surat_pesan   || '-',
          district:     payloadData.district      || '-',
          lokasi_smile: payloadData.lokasi_smile  || '-',
        });
        // GRID BARU: tiap ODP dibagi menjadi baris 3 port (ke samping)
        payloadData.redaman_list = payloadData.redaman_list.map(o =>
          Object.assign({}, o, { baris: barisGridRedaman(o.port_list, 3) })
        );
        payloadData.redaman_ada    = payloadData.redaman_list.length > 0;
        payloadData.redaman_kosong = payloadData.redaman_list.length === 0;

        // ── REKAP OPM RELOKASI (INPUT REKAP OPM) ─────────────────────────────
        try {
          const folderRekap = path.join(__dirname, 'uploads', namaTematikClean, namaLokasiClean, 'INPUT_REKAP_OPM');
          pastikanFolderEksis(folderRekap);
          const berkasRekap = await db.all(
            'SELECT sub_kategori, file_name FROM berkas WHERE lokasi_id = ? AND kategori = ? AND (file_name LIKE "%.png" OR file_name LIKE "%.jpg" OR file_name LIKE "%.jpeg" OR file_name LIKE "%.webp" OR file_name LIKE "%.pdf")',
            [lokasiId, 'INPUT REKAP OPM']
          );
          const petaRekap = {};
          berkasRekap.forEach(b => { const k = (b.sub_kategori || '').trim(); if (k) { petaRekap[k] = path.join(folderRekap, b.file_name); } });
          payloadData.rekap_opm_list = (payloadData.opm_list || []).map(o => {
            const p = petaRekap[o.nama];
            return { nama: o.nama, LDP1: (p && fs.existsSync(p)) ? p : blankImage };
          });
        } catch (e) {
          console.warn('rekap_opm_list error:', e.message);
          payloadData.rekap_opm_list = [];
        }

        // ── TAMPLATE_NODE_B.docx (Tematik Node B) — bagian "EVIDEN REDAMAN ODP/OTB" ──
        // Template Node B dibuat ke samping: loop {#opm_list} dengan sel datar
        // {%foto1..16}{status1..16} per ODP (tag {nama}/{kapasitas} tetap dipakai).
        // Field datar foto{n}/status{n} sudah disiapkan siapkanDataOpm, jadi cukup
        // arahkan opm_list ke data redaman — HANYA bila template Node B dipakai,
        // karena template lain (termasuk TAMPLATE_BAUT_NODE_B) masih memakai opm_list
        // untuk eviden OPM asli.
        if (path.basename(finalTemplate) === 'TAMPLATE_NODE_B.docx') {
          payloadData.opm_list = payloadData.redaman_list.map(o => {
            const datar = Object.assign({}, o);
            // Port melebihi kapasitas ODP (sel statis 1..16 di template): foto kosong
            // (blank.png) + tanda "—" supaya tidak ada tulisan status yang menggantung.
            for (let n = (Number(o.kapasitas) || 0) + 1; n <= 16; n++) {
              if (datar['foto' + n]   === undefined) datar['foto' + n]   = null;
              if (datar['status' + n] === undefined) datar['status' + n] = '—';
            }
            return datar;
          });
        }

        // ── LOOP GENERIK SEMUA KATEGORI (template_HEM / template_MICRODEMAND) ──
        // Template: {#kategori_list} KATEGORI {nomor}. {nama} ({jumlah} foto)
        //           {#boq}…{/boq}  {#bukan_boq}{#kosong}…{/kosong}{#terlampir}…{/terlampir}{#ada}{#foto_baris}{%f1}…{/foto_baris}{/ada}{/bukan_boq}
        //           {/kategori_list}
        const urutWajib = DATABASE_TEMATIK_WAJIB[lokasiData.tematik] || [];
        const kategoriUrut = [...urutWajib];
        semuaBerkas.forEach(b => { if (!kategoriUrut.includes(b.kategori)) kategoriUrut.push(b.kategori); });
        payloadData.tematik = lokasiData.tematik || 'Umum';
        const fmtAngka = (n) => (Number(n) || 0).toLocaleString('id-ID', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
        const boqRowsAll = await db.all(
          'SELECT designator, uraian, satuan, material, jasa, volume FROM boq_aktual WHERE lokasi_id = ? ORDER BY id ASC',
          [lokasiId]
        );

        // ── Expand EVIDEN PEKERJAAN per sub_kategori ──
        const NAMA_EEP = 'MASUKAN EVIDEN PEKERJAAN';
        const IMG_EXT = /\.(jpe?g|png|webp|gif|bmp)$/i;
        const berkasEepAll = semuaBerkas.filter(b => b.kategori === NAMA_EEP);
        const eepBySub = {};
        berkasEepAll.forEach(b => {
          const sub = (b.sub_kategori || '').trim() || (b.file_name || '').replace(/\.[^.]+$/i, '').trim() || 'Eviden';
          if (!eepBySub[sub]) eepBySub[sub] = [];
          eepBySub[sub].push(b);
        });
        const eepSubKeys = Object.keys(eepBySub).sort((a, b) => a.localeCompare(b, 'id'));

        // Build final list: non-EEP categories first, then expanded EEP sub-categories
        const kategoriFinal = [];
        kategoriUrut.forEach(k => {
          if (k === NAMA_EEP) {
            eepSubKeys.forEach(sub => {
              kategoriFinal.push({ nama: `${NAMA_EEP} — ${sub}`, files: eepBySub[sub] });
            });
          } else {
            kategoriFinal.push({ nama: k, files: semuaBerkas.filter(b => b.kategori === k) });
          }
        });

        // ── "TIDAK BERLAKU" (audit skip, ditandai admin di panel Audit Berkas) ──
        // Kategori yang ditandai tidak berlaku ATAU yang berkasnya kosong
        // tidak ditampilkan sama sekali di laporan hasil generate.
        const skipRows  = await db.all('SELECT item FROM audit_skip WHERE lokasi_id = ?', [lokasiId]);
        const skipSet   = new Set(skipRows.map(r => (r.item || '').trim()).filter(Boolean));
        const diSkip    = (nama) => skipSet.has(nama) || (nama.includes(' — ') && skipSet.has(nama.split(' — ')[0]));

        const kategoriItems = kategoriFinal
          .filter(item => item.nama !== KATEGORI_REDAMAN && !KATEGORI_KML_FULL.has(item.nama)
            // Node B: tiga inputan ini punya bagian template sendiri (tabel barang
            // tiba, gambar form OPM, loop file .SOR) — jangan digandakan di sini.
            && item.nama !== 'BOQ BARANG TIBA'
            && item.nama !== 'FORM PENGUKURAN OPM'
            && item.nama !== 'OTDR (.SOR)')
          // Ditandai "Tidak Berlaku" → jangan tampil di laporan.
          .filter(item => !diSkip(item.nama))
          // Tanpa berkas → jangan tampil (kecuali BOQ AKTUAL yang baris tabelnya ada).
          .filter(item => item.files.length > 0 || (item.nama === 'BOQ AKTUAL' && boqRowsAll.length > 0));
        // for-of (bukan map) karena berkas PDF harus di-await dulu → jadi gambar halaman
        payloadData.kategori_list = [];
        for (let i = 0; i < kategoriItems.length; i++) {
          const item = kategoriItems[i];
          const nama = item.nama;
          const berkasKategori = item.files;
          const fotoFiles = berkasKategori.filter(b => IMG_EXT.test(b.file_name)).sort((a, b) => (a.file_name || '').localeCompare(b.file_name || ''));
          const pdfFiles  = berkasKategori.filter(b => /\.pdf$/i.test(b.file_name));
          const dwgFiles  = berkasKategori.filter(b => /\.(dwg|dxf)$/i.test(b.file_name));
          const sorFiles  = berkasKategori.filter(b => /\.sor$/i.test(b.file_name));
          const fotos = fotoFiles
            .map(b => {
              const katClean = b.kategori.replace(/[^a-zA-Z0-9_-]/g, '_');
              const fullPath = path.join(__dirname, 'uploads', namaTematikClean, namaLokasiClean, katClean, b.file_name);
              return fs.existsSync(fullPath) ? fullPath : null;
            })
            .filter(Boolean);
          // PDF (mis. INPUT RFC .pdf / ABD .pdf) → render tiap halaman jadi gambar,
          // ikut mengisi grid {#foto_baris} / {#penuh}{%foto_penuh}. DWG/DXF tetap
          // masuk fileLain → "Berkas terlampir: nama.dwg" (butuh CAD untuk render).
          for (const b of pdfFiles) {
            const katClean = b.kategori.replace(/[^a-zA-Z0-9_-]/g, '_');
            const fullPath = path.join(__dirname, 'uploads', namaTematikClean, namaLokasiClean, katClean, b.file_name);
            if (!fs.existsSync(fullPath)) continue;
            (await renderPdfHalaman(fullPath)).forEach(p => fotos.push(p));
          }
          // ── DWG/DXF → konversi via AutoCAD → gambar; .SOR → grafik trace ──
          // Berhasil terkonversi → tampil sebagai gambar (bukan sekadar nama).
          const konversiOk = new Set();
          for (const b of dwgFiles) {
            const katClean = b.kategori.replace(/[^a-zA-Z0-9_-]/g, '_');
            const fullPath = path.join(__dirname, 'uploads', namaTematikClean, namaLokasiClean, katClean, b.file_name);
            if (!fs.existsSync(fullPath)) continue;
            const pdfDwg = await konversiDwgKePdf(fullPath);
            if (!pdfDwg) continue;
            const pngs = await renderPdfDwg(pdfDwg);
            if (pngs.length) {
              pngs.forEach(p => fotos.push(p));
              konversiOk.add(b.file_name);
            }
          }
          for (const b of sorFiles) {
            const katClean = b.kategori.replace(/[^a-zA-Z0-9_-]/g, '_');
            const fullPath = path.join(__dirname, 'uploads', namaTematikClean, namaLokasiClean, katClean, b.file_name);
            if (!fs.existsSync(fullPath)) continue;
            const png = await renderSorGambar(fullPath);
            if (png) {
              fotos.push(png);
              konversiOk.add(b.file_name);
            }
          }
          const fileLain = berkasKategori.filter(b =>
            !IMG_EXT.test(b.file_name) && !/\.pdf$/i.test(b.file_name) && !konversiOk.has(b.file_name));
          const foto_baris = [];
          for (let k = 0; k < fotos.length; k += 4) {
            const r4 = fotos.slice(k, k + 4);
            foto_baris.push({ f1: r4[0] || blankImage, f2: r4[1] || blankImage, f3: r4[2] || blankImage, f4: r4[3] || blankImage });
          }
          const item_out = {
            nomor: i + 1,
            nama,
            jumlah: fotos.length + fileLain.length,
            ada: fotos.length > 0,
            kosong: berkasKategori.length === 0,
            terlampir: fileLain.length > 0,
            nama_file: fileLain.map(b => b.file_name).join(', '),
            foto_baris,
            // 1 foto = 1 halaman penuh, foto diperbesar ({#penuh}{%foto_penuh}
            // dipakai template FIXED_template_HEM / FIXED_template_MICRODEMAND)
            penuh: fotos.length === 1,
            foto_penuh: fotos[0] || blankImage,
            bukan_boq: nama !== 'BOQ AKTUAL'
          };
          if (nama === 'BOQ AKTUAL') {
            const boqRows = boqRowsAll;
            const boq_list = boqRows.map((b, bi) => ({
              nomor: bi + 1,
              designator: b.designator || '-',
              uraian: b.uraian || '-',
              satuan: b.satuan || '-',
              material: fmtAngka(b.material),
              jasa: fmtAngka(b.jasa),
              volume: fmtAngka(b.volume),
            }));
            const totMat = boqRows.reduce((s, b) => s + (Number(b.material) || 0), 0);
            const totJas = boqRows.reduce((s, b) => s + (Number(b.jasa) || 0), 0);
            Object.assign(item_out, {
              boq: boq_list.length > 0,
              boq_kosong: boq_list.length === 0,
              boq_list,
              boq_total_material: fmtAngka(totMat),
              boq_total_jasa: fmtAngka(totJas),
              kosong: false, ada: false, terlampir: false,
              penuh: false,
            });
          }
          payloadData.kategori_list.push(item_out);
        }

        // ── PETA / KML — GAMBAR FULL PAGE (1 gambar = 1 halaman) ──
        // Template: {#kml_ada}{#kml_list} {nama} {%foto} {/kml_list}{/kml_ada}
        //           {#kml_kosong}— Belum ada peta/KML —{/kml_kosong}
        const kmlBerkas = semuaBerkas.filter(b => KATEGORI_KML_FULL.has(b.kategori));
        const IMG_EXT2 = /\.(jpe?g|png|webp|gif|bmp)$/i;
        const kmlFotoFiles = kmlBerkas.filter(b => IMG_EXT2.test(b.file_name)).sort((a, b) => (a.file_name || '').localeCompare(b.file_name || ''));
        // Gambar IDENTIK (upload rute yang sama ke 2 kategori, atau render sama persis)
        // cukup tampil sekali — kalau tidak, dokumen jadi dobel/kembar.
        const kmlHashTerlihat = new Set();
        payloadData.kml_list = [];
        for (const b of kmlFotoFiles) {
          const katClean = b.kategori.replace(/[^a-zA-Z0-9_-]/g, '_');
          const fullPath = path.join(__dirname, 'uploads', namaTematikClean, namaLokasiClean, katClean, b.file_name);
          if (!fs.existsSync(fullPath)) {
            payloadData.kml_list.push({ nama: b.sub_kategori || 'Peta Rute', foto: blankImage });
            continue;
          }
          let h = null;
          try {
            h = require('crypto').createHash('sha1').update(fs.readFileSync(fullPath)).digest('hex');
          } catch (e) {}
          if (h && kmlHashTerlihat.has(h)) {
            console.log('[KML] gambar identik dilewati: ' + b.kategori + '/' + b.file_name);
            continue;
          }
          if (h) kmlHashTerlihat.add(h);
          payloadData.kml_list.push({ nama: b.sub_kategori || 'Peta Rute', foto: fullPath });
        }
        payloadData.kml_ada    = payloadData.kml_list.length > 0;
        payloadData.kml_kosong = false; // placeholder "Belum ada peta/KML" tidak dicetak

        // Data input khusus (BOQ AKTUAL / INPUT RFC / BON MATERIAL)
        const semuaInput = await db.all(
          'SELECT kategori, judul, volume, isi, created_at FROM input_khusus WHERE lokasi_id = ? ORDER BY id ASC',
          [lokasiId]
        );
        payloadData.input_list = semuaInput.filter(r => !diSkip(r.kategori));
        payloadData.ada_input    = payloadData.input_list.length > 0;
        payloadData.kosong_input = false; // placeholder "belum ada input" tidak dicetak

        // ── NODE B (LACT) — tag khusus TAMPLATE_NODE_B.docx ──────────────────
        // Catatan: smartParser memangkas spasi di dalam tag, jadi kunci payload
        // ditulis TANPA spasi ganda/ekor ("tgl ct nodeb", bukan "tgl ct nodeb ").
        payloadData['hari ct nodeb']   = adminInput.ct_hari  || '-';
        payloadData['tgl ct nodeb']    = adminInput.ct_tgl   || '-';
        payloadData['bulan ct nodeb']  = adminInput.ct_bulan || '-';
        payloadData['tgl-bulan-thn ct'] = adminInput.tgl_bulan_thn_ct || '-';

        // Wasapang: template punya varian kapital & huruf kecil — isi keduanya.
        payloadData['waspang Telkom']    = adminInput.nama_waspang || '-';
        payloadData['nik waspang Telkom'] = adminInput.nik_waspang || '-';
        payloadData['waspang telkom']    = adminInput.nama_waspang || '-';
        payloadData['nik waspang telkom'] = adminInput.nik_waspang || '-';
        payloadData['waspang TA']        = adminInput.waspang_ta || '-';
        payloadData['nik waspang TA']    = adminInput.nik_waspang_ta || '-';

        payloadData['hari barang tiba']  = adminInput.hari_barang_tiba  || '-';
        payloadData['tgl barang tiba']   = adminInput.tgl_barang_tiba   || '-';
        payloadData['bulan barang tiba'] = adminInput.bulan_barang_tiba || '-';
        payloadData['tgl-bulan-thn barang tiba'] = adminInput.tgl_bulan_thn_barang_tiba || '-';
        payloadData['tgl-bulan-thn opm'] = adminInput.tgl_bulan_thn_opm || '-';

        payloadData['pm area']        = adminInput.pm_area || '-';
        payloadData['site']           = adminInput.site || '-';
        payloadData['Ruas Pekerjaan'] = adminInput.ruas_pekerjaan || '-';
        payloadData['izin']           = adminInput.izin || '-';
        // ({lokasi_smile} sudah diisi payload utama dari lokasiData.nama_lokasi)

        // BOQ BARANG TIBA — loop tabel di template ({#barang_tiba_list}…)
        const btRows = await db.all(
          'SELECT no, jenis_material, satuan, volume, keterangan FROM barang_tiba WHERE lokasi_id = ? ORDER BY no ASC, id ASC',
          [lokasiId]
        );
        payloadData.barang_tiba_list = (btRows.length
          ? btRows.map((r, i) => ({
              no: r.no || (i + 1),
              jenis_material: r.jenis_material || '-',
              satuan: r.satuan || '-',
              volume: r.volume || '-',
              keterangan: r.keterangan || '-'
            }))
          : [{ no: '-', jenis_material: 'Belum ada data barang tiba', satuan: '-', volume: '-', keterangan: '-' }]
        ).map(r => {
          // Tag template memakai kunci pendek — isi dua-duanya.
          r.nama = r.jenis_material;
          return r;
        });
        payloadData['BOQ BARANG TIBA'] = payloadData.barang_tiba_list.map(r =>
          `${r.no}. ${r.jenis_material} — ${r.satuan} — ${r.volume} — ${r.keterangan}`).join('\n');

        // FORM PENGUKURAN OPM — gambar terakhir yang diupload (tag {%…})
        const opmForm = await db.get(
          "SELECT file_name, kategori FROM berkas WHERE lokasi_id = ? AND kategori = 'FORM PENGUKURAN OPM' ORDER BY id DESC LIMIT 1",
          [lokasiId]
        );
        let opmPath = null;
        if (opmForm) {
          const katClean = opmForm.kategori.replace(/[^a-zA-Z0-9_-]/g, '_');
          const p = path.join(__dirname, 'uploads', namaTematikClean, namaLokasiClean, katClean, opmForm.file_name);
          if (fs.existsSync(p)) opmPath = p;
        }
        payloadData.opm_form_ada = !!opmPath && !diSkip('FORM PENGUKURAN OPM');
        payloadData.opm_form_kosong = false; // placeholder "belum ada form" tidak dicetak
        payloadData['FORM PENGUKURAN OPM'] = opmPath || blankImage;

        // OTDR (.SOR) — loop daftar file (diupload apa adanya, tanpa konversi)
        const sorBerkas = await db.all(
          "SELECT file_name FROM berkas WHERE lokasi_id = ? AND kategori = 'OTDR (.SOR)' ORDER BY id ASC",
          [lokasiId]
        );
        const sorTampil = sorBerkas.length > 0 && !diSkip('OTDR (.SOR)');
        payloadData.sor_list = [];
        if (sorTampil) {
          const katFolderSor = 'OTDR (.SOR)'.replace(/[^a-zA-Z0-9_-]/g, '_');
          for (let i = 0; i < sorBerkas.length; i++) {
            const fp = path.join(__dirname, 'uploads', namaTematikClean, namaLokasiClean, katFolderSor, sorBerkas[i].file_name);
            const png = fs.existsSync(fp) ? await renderSorGambar(fp) : null;
            payloadData.sor_list.push({ no: i + 1, nama_file: sorBerkas[i].file_name, sor_gambar: png });
          }
        }
        payloadData.sor_ada = sorTampil;
        payloadData.sor_kosong = false; // placeholder "belum ada .SOR" tidak dicetak

        // ── BAUT-1 Node B: 5 dokumen hasil upload admin (folder BAUT1_*) ──
        if (modeBaut1) {
            const daftarBaut1 = [
                ['nota_dinas',        'nota_dinas_hal',     'nota_dinas_kosong'],
                ['permohonan_tim_ut', 'permohonan_tim_hal', 'permohonan_tim_kosong'],
                ['penunjukan_tim_ut', 'penunjukan_tim_hal', 'penunjukan_tim_kosong'],
                ['form_ceklis',       'form_ceklis_hal',    'form_ceklis_kosong'],
                ['ba_prelim',         'ba_prelim_hal',      'ba_prelim_kosong'],
            ];
            for (const [folderKey, arrKey, kosongKey] of daftarBaut1) {
                const dir = path.join(__dirname, 'uploads', namaTematikClean, namaLokasiClean,
                    'BAUT1_' + folderKey.toUpperCase());
                const files = fs.existsSync(dir)
                    ? fs.readdirSync(dir).filter(f => /^page-\d+\.(jpg|jpeg|png)$/i.test(f)).sort()
                    : [];
                payloadData[arrKey] = files.map(f => ({ baut_page: path.join(dir, f) }));
                payloadData[kosongKey] = files.length === 0;
                console.log('[BAUT1] ' + folderKey + ': ' + files.length + ' halaman');
            }
        }

        console.log('[RENDER] eep_list length:', (payloadData.eep_list || []).length);
        console.log('[RENDER] payload keys:', Object.keys(payloadData).join(', '));
        console.log('[RENDER] eep_list[0]?', JSON.stringify((payloadData.eep_list || [])[0] || null));
        doc.render(payloadData);

        const buf0 = doc.getZip().generate({
    type: 'nodebuffer',
    compression: 'DEFLATE'
});

// Fix: pastikan gif terdaftar di Content_Types
const zipFix = new PizZip(buf0);
const ctFile = zipFix.file('[Content_Types].xml');
if (ctFile) {
    let ctText = ctFile.asText();
    if (!ctText.includes('Extension="gif"')) {
        ctText = ctText.replace('</Types>', '<Default Extension="gif" ContentType="image/gif"/></Types>');
        zipFix.file('[Content_Types].xml', ctText);
    }
}
const buf0fixed = zipFix.generate({ type: 'nodebuffer', compression: 'DEFLATE' });
        // PETA/KML sudah tampil di badan dokumen lewat {#kml_list} → jangan tambah
        // lampiran peta lagi di akhir (dulu gambarnya dobel/kembar).
        let templateAdaKml = false;
        try {
          templateAdaKml = new PizZip(fs.readFileSync(finalTemplate)).file('word/document.xml')
            .asText().includes('{#kml_list}');
        } catch (e) {}
        const lewatiLampiranPeta = templateAdaKml && (payloadData.kml_list || []).length > 0;
        const buf = await lampirkanPetaKeDocx(buf0fixed, lokasiId, { sisipPeta: !lewatiLampiranPeta });
        const namaUnduh = modeBaut1 ? `BAUT1_${namaLokasiClean}.docx` : `Laporan_${namaLokasiClean}.docx`;
        await kirimDokumen(req, res, buf, namaUnduh);

    } catch (error) {
        console.error("Error cetak laporan:", error);
        
        // >>> TAMBAHKAN KODE INI DI SINI <<<
        if (error.properties && error.properties.errors) {
            console.log("=== DETAIL MULTI ERROR DOCXTEMPLATER ===");
            error.properties.errors.forEach((e, index) => {
                console.log(`Error ${index + 1}:`, e.properties);
            });
        }

        res.status(500).json({ message: "Gagal menyusun dokumen.", detail: error.message });
    }
});

// ─────────────────────────────────────────────────────────────────────────────
// [API] RAB — Rencana Anggaran Biaya (VERSI LENGKAP + SUMMARY CACHE)
// ─────────────────────────────────────────────────────────────────────────────

// GET semua item RAB untuk satu lokasi
app.get('/api/rab/:lokasi_id', checkAuth(), async (req, res) => {
  try {
    const items = await db.all(
      `SELECT * FROM rab WHERE lokasi_id = ? ORDER BY no ASC, id ASC`,
      [req.params.lokasi_id]
    );
    res.json(items);
  } catch (e) { res.status(500).json({ error: e.message }); }
});

// POST — simpan/replace semua item + perbarui cache rab_summary
app.post('/api/rab/:lokasi_id', checkAuth('admin'), async (req, res) => {
  const lokasiId = req.params.lokasi_id;
  const { items, totalMaterial, totalJasa } = req.body;
  if (!Array.isArray(items))
    return res.status(400).json({ error: 'Body harus berisi field "items" berupa array.' });
  try {
    await db.run(`DELETE FROM rab WHERE lokasi_id = ?`, [lokasiId]);
    const stmt = await db.prepare(
      `INSERT INTO rab (lokasi_id, no, designator, des_material, des_jasa, kategori, satuan, material, jasa, volume)
       VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`
    );
    for (const item of items) {
      await stmt.run(
        lokasiId,
        item.no           ?? null,
        item.designator   ?? '',
        item.des_material ?? null,
        item.des_jasa     ?? null,
        item.kategori     ?? 'Gabungan',
        item.satuan       ?? null,
        parseFloat(item.material) || 0,
        parseFloat(item.jasa)     || 0,
        parseFloat(item.volume)   || 0,
      );
    }
    await stmt.finalize();

    // Hitung summary (dari body jika tersedia, atau hitung ulang)
    const tMat  = typeof totalMaterial === 'number'
      ? totalMaterial
      : items.reduce((s,r) => s + ((parseFloat(r.material)||0)*(parseFloat(r.volume)||0)), 0);
    const tJasa = typeof totalJasa === 'number'
      ? totalJasa
      : items.reduce((s,r) => s + ((parseFloat(r.jasa)||0)*(parseFloat(r.volume)||0)), 0);
    const tGrand = tMat + tJasa;
    const tCount = items.length;

    // Upsert rab_summary (cache)
    const exist = await db.get(`SELECT lokasi_id FROM rab_summary WHERE lokasi_id = ?`, [lokasiId]);
    if (exist) {
      await db.run(
        `UPDATE rab_summary SET total_material=?, total_jasa=?, grand_total=?, rab_count=?, updated_at=datetime('now','localtime') WHERE lokasi_id=?`,
        [tMat, tJasa, tGrand, tCount, lokasiId]
      );
    } else {
      await db.run(
        `INSERT INTO rab_summary (lokasi_id, total_material, total_jasa, grand_total, rab_count) VALUES (?,?,?,?,?)`,
        [lokasiId, tMat, tJasa, tGrand, tCount]
      );
    }
    console.log(` RAB saved: lokasi=${lokasiId}, item=${tCount}, total=Rp${fmtLog(tGrand)}`);
    res.json({ ok: true, saved: tCount, total_material: tMat, total_jasa: tJasa, grand_total: tGrand });
  } catch (e) {
    console.error('Error POST /api/rab:', e);
    res.status(500).json({ error: e.message });
  }
});

// DELETE — hapus semua RAB + hapus cache summary
app.delete('/api/rab/:lokasi_id', checkAuth('admin'), async (req, res) => {
  try {
    await db.run(`DELETE FROM rab         WHERE lokasi_id = ?`, [req.params.lokasi_id]);
    await db.run(`DELETE FROM rab_summary WHERE lokasi_id = ?`, [req.params.lokasi_id]);
    res.json({ ok: true });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

// GET summary — dari cache rab_summary (fallback: hitung langsung dari tabel rab)
// checkAuth() tanpa role agar bisa diakses user maupun admin (untuk audit checker)
app.get('/api/rab/:lokasi_id/summary', checkAuth(), async (req, res) => {
  try {
    // Coba dari cache dulu
    const cached = await db.get(
      `SELECT total_material, total_jasa, grand_total, rab_count, updated_at FROM rab_summary WHERE lokasi_id = ?`,
      [req.params.lokasi_id]
    );
    if (cached && cached.rab_count > 0) {
      const desRows = await db.all(
        `SELECT DISTINCT designator FROM rab WHERE lokasi_id = ? ORDER BY designator ASC`,
        [req.params.lokasi_id]
      );
      return res.json({
        count:          cached.rab_count,
        total_material: cached.total_material,
        total_jasa:     cached.total_jasa,
        total:          cached.grand_total,
        updated_at:     cached.updated_at,
        designators:    desRows.map(r => r.designator).filter(Boolean),
      });
    }
    // Fallback: hitung langsung
    const items = await db.all(
      `SELECT designator, material, jasa, volume FROM rab WHERE lokasi_id = ? ORDER BY no ASC`,
      [req.params.lokasi_id]
    );
    const tMat  = items.reduce((s,r) => s + ((r.material||0)*(r.volume||0)), 0);
    const tJasa = items.reduce((s,r) => s + ((r.jasa||0)*(r.volume||0)), 0);
    const designators = [...new Set(items.map(r => r.designator).filter(Boolean))].sort();
    res.json({ count: items.length, total_material: tMat, total_jasa: tJasa, total: tMat + tJasa, designators });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

// GET daftar designator unik untuk lokasi (dipakai audit checker admin.html)
app.get('/api/rab/:lokasi_id/designators', checkAuth(), async (req, res) => {
  try {
    const rows = await db.all(
      `SELECT DISTINCT designator FROM rab WHERE lokasi_id = ? ORDER BY designator ASC`,
      [req.params.lokasi_id]
    );
    res.json(rows.map(r => r.designator).filter(Boolean));
  } catch (e) { res.status(500).json({ error: e.message }); }
});

// ─────────────────────────────────────────────────────────────────────────────
// [API] Auth
// ─────────────────────────────────────────────────────────────────────────────
// Dua "kunci" tambahan (bisa diganti lewat environment variable):
//   AKSES_PENGGUNA / akses123 → password kedua untuk membuka menu Pengguna
//                               pada admin biasa (login dengan admin123).
//   MASTER_PASSWORD / master123 → password kedua pada akun 'admin' yang sama;
//                               login dengan ini menandakan peran "master".
const AKSES_PENGGUNA = String(process.env.AKSES_PENGGUNA || 'akses123');
const PASS_MASTER = String(process.env.MASTER_PASSWORD || 'master123');

// Penjaga menu Pengguna: master PWEN dibebaskan; admin biasa wajib membuka
// kunci dulu lewat POST /api/pengguna/unlock (flag tersimpan di session).
function checkPengguna(req, res, next) {
  const u = req.session && req.session.user;
  if (!u) return res.status(401).json({ ok: false, message: 'Sesi Anda telah berakhir. Silakan login kembali.' });
  if (u.role !== 'admin') return res.status(403).json({ ok: false, message: 'Hak akses ditolak: bukan administrator.' });
  if (u.master || u.penggunaUnlock) return next();
  return res.status(423).json({ ok: false, terkunci: true, message: 'Menu Pengguna terkunci. Masukkan password akses.' });
}

app.post('/api/login', async (req, res) => {
  try {
    const { username, password } = req.body;
    const user = await db.get('SELECT * FROM users WHERE username = ?', [username]);
    if (!user) return res.status(400).json({ message: 'Username tidak terdaftar. Hubungi administrator.' });

    // Akun lapangan (role 'user') cukup menulis username — tanpa password.
    // Akun administrator tetap memerlukan password. master123 adalah password
    // kedua di akun admin yang SAMA → login sebagai master (bisa putus sesi).
    let master = false;
    if (user.role === 'admin') {
      if (!password) {
        return res.status(401).json({ message: 'Akun administrator memerlukan password.', need_password: true });
      }
      const match = await bcrypt.compare(password, user.password);
      if (!match) {
        if (password === PASS_MASTER) master = true;
        else return res.status(400).json({ message: 'Password salah. Coba lagi.' });
      }
    } else if (password) {
      // password opsional untuk role user; kalau diisi harus benar
      const match = await bcrypt.compare(password, user.password);
      if (!match) return res.status(400).json({ message: 'Password salah. Coba lagi.' });
    }

    // Daftar tematik yang boleh dibuka akun ini ikut tersimpan di sesi
    // (LOGIN TEMATIK) sehingga penjaga akses tidak perlu membaca DB tiap request.
    const tematik = user.role === 'admin' ? [] : bacaTematikUser(user);
    req.session.user = {
      id: user.id, username: user.username, role: user.role, tematik: tematik,
      master: master,                      // true hanya bila login dengan master123
      penggunaUnlock: master || false,     // master langsung boleh buka Pengguna
      // Info perangkat untuk panel "Sesi login aktif" (admin)
      dev: { ua: String(req.headers['user-agent'] || ''), ip: ipKlien(req), masuk: Date.now() },
      akhir: Date.now()
    };
    res.json({
      success: true, role: user.role, tematik: tematik,
      master: master, pengguna_unlock: !!req.session.user.penggunaUnlock
    });
  } catch (e) {
    console.error(e);
    res.status(500).json({ message: 'Terjadi gangguan sistem pada layanan login.' });
  }
});

// Buka / kunci kembali menu Pengguna untuk admin biasa.
app.post('/api/pengguna/unlock', checkAuth('admin'), (req, res) => {
  if (req.session.user.master) return res.json({ ok: true, sudah_master: true });
  const pass = String((req.body && req.body.password) || '').trim();
  if (pass !== AKSES_PENGGUNA) {
    return res.status(403).json({ ok: false, message: 'Password akses salah.' });
  }
  req.session.user.penggunaUnlock = true;
  res.json({ ok: true, sudah_master: false });
});

app.post('/api/pengguna/lock', checkAuth('admin'), (req, res) => {
  req.session.user.penggunaUnlock = false;
  res.json({ ok: true });
});

app.get('/api/logout', (req, res) => {
  req.session.destroy();
  res.json({ success: true });
});

// ─────────────────────────────────────────────────────────────────────────────
// [API] PENGELOLAAN PENGGUNA — khusus admin
// Hanya akun lapangan (role 'user') yang bisa ditambah/ubah/hapus; akun admin
// bawaan sistem (admin:admin123) dijaga agar tidak hilang.
// ─────────────────────────────────────────────────────────────────────────────
const USER_RX = /^[A-Za-z0-9._-]{3,32}$/;

function cekFormatUsername(username) {
  if (!username) return 'Username wajib diisi.';
  if (!USER_RX.test(username)) {
    return 'Username 3-32 karakter, hanya huruf, angka, titik, garis bawah, atau strip (tanpa spasi).';
  }
  return null;
}

// Buang sesi aktif milik pengguna tertentu (dipakai saat ganti/hapus akun)
async function buangSesiUser(cocok) {
  try {
    const rows = await db.all('SELECT id, data FROM sessions');
    for (const r of rows) {
      try {
        const sesi = JSON.parse(r.data || '{}');
        if (sesi && sesi.user && cocok(sesi.user)) {
          await db.run('DELETE FROM sessions WHERE id = ?', [r.id]);
        }
      } catch (err) { /* sesi rusak — abaikan */ }
    }
  } catch (err) { /* gagal membersihkan sesi tidak menghalangi aksi */ }
}

// ─────────────────────────────────────────────────────────────────────────────
// [API] SESI LOGIN AKTIF — khusus admin
// Menampilkan berapa sesi (login) yang sedang aktif, plus detail perangkat
// tiap sesi: OS, peramban, IP, waktu masuk & terakhir aktif.
// Sumber: tabel sessions (persisten, tidak hilang saat server restart).
// ─────────────────────────────────────────────────────────────────────────────
function deteksiPerangkat(ua) {
  const s = String(ua || '');
  let os = 'Tidak diketahui';
  if (/windows/i.test(s)) os = 'Windows';
  else if (/android/i.test(s)) os = 'Android';
  else if (/iphone|ipad|ipod|\bios\b/i.test(s)) os = 'iOS';
  else if (/mac os|macintosh/i.test(s)) os = 'macOS';
  else if (/linux/i.test(s)) os = 'Linux';
  if (/bot|crawler|spider|curl|wget|python-requests/i.test(s)) return { os: 'Otomatis', peramban: 'Bot/script' };

  let peramban = 'Lainnya';
  if (/edg\//i.test(s)) peramban = 'Edge';
  else if (/opr\/|opera/i.test(s)) peramban = 'Opera';
  else if (/firefox\//i.test(s)) peramban = 'Firefox';
  else if (/chrome\/|crios/i.test(s)) peramban = 'Chrome';
  else if (/safari\//i.test(s)) peramban = 'Safari';
  return { os, peramban };
}

// Ambang "masih online": aktivitas terakhir < 5 menit. Sesi yang lebih lama
// ditandai Diam dan boleh dihapus oleh master lewat panel Pengguna.
const SESI_ONLINE_MS = 5 * 60 * 1000;
// Masa berlaku cookie sesi (sama dengan cookie.maxAge di app.use(session)).
const SESI_USIA_MAKS_MS = 24 * 60 * 60 * 1000;

// "5 menit lalu" / "2 jam lalu" — ringkas untuk tampilan
function relatifWaktu(ms) {
  if (!ms) return '—';
  const d = Math.max(0, Date.now() - ms);
  const m = Math.floor(d / 60000);
  if (m < 1) return 'baru saja';
  if (m < 60) return m + ' mnt lalu';
  const j = Math.floor(m / 60);
  if (j < 24) return j + ' jam lalu';
  return Math.floor(j / 24) + ' hari lalu';
}

app.get('/api/sesi-aktif', checkAuth('admin'), async (req, res) => {
  try {
    const rows = await db.all('SELECT id, data, updated_at FROM sessions');
    const now = Date.now();
    const unlock = !!(req.session.user.master || req.session.user.penggunaUnlock);
    const daftar = [];
    for (const r of rows) {
      let sesi;
      try { sesi = JSON.parse(r.data || '{}'); } catch (e) { continue; }
      if (!sesi || !sesi.user) continue;                       // sesi tanpa login — lewati
      if (sesi.cookie && sesi.cookie.expires && new Date(sesi.cookie.expires).getTime() < now) continue; // kedaluwarsa

      // updated_at tersimpan sebagai waktu lokal ('YYYY-MM-DD HH:MM:SS')
      let updLocal = 0;
      if (r.updated_at) {
        const t = Date.parse(String(r.updated_at).replace(' ', 'T'));
        if (!isNaN(t)) updLocal = t;
      }
      const dev = sesi.user.dev || {};
      const masuk = dev.masuk || updLocal || 0;
      const akhir = Math.max(sesi.user.akhir || 0, updLocal);
      const { os, peramban } = deteksiPerangkat(dev.ua);
      daftar.push({
        id: String(r.id).slice(0, 8),
        // sid lengkap + User-Agent hanya diberikan setelah menu Pengguna dibuka
        ...(unlock ? { sid: String(r.id), ua: String(dev.ua || '') } : {}),
        username: sesi.user.username,
        role: sesi.user.role,
        master: !!sesi.user.master,
        os, peramban,
        ip: dev.ip || '—',
        masuk, akhir,
        umur_masuk: relatifWaktu(masuk),
        terakhir: relatifWaktu(akhir),
        online: (now - akhir) < SESI_ONLINE_MS, // aktif < 5 menit = online
        sendiri: String(r.id) === String(req.sessionID)
      });
    }
    daftar.sort((a, b) => (b.online - a.online) || (b.akhir - a.akhir));
    const unikUser = new Set(daftar.map(d => d.username));
    res.json({
      ok: true,
      terkunci: !unlock,          // jumlah untuk badge topbar tetap bisa dibaca
      master: !!req.session.user.master,
      total: daftar.length,
      online: daftar.filter(d => d.online).length,
      diam: daftar.filter(d => !d.online).length,
      pengguna_unik: unikUser.size,
      daftar: unlock ? daftar : []   // admin yang belum unlock hanya dapat angka
    });
  } catch (e) {
    console.error('[sesi-aktif]', e);
    res.status(500).json({ ok: false, message: 'Gagal membaca daftar sesi.' });
  }
});

// PUTUS SESI — khusus admin master. Sesi sendiri sengaja dikecualikan agar
// master tidak mengunci dirinya sendiri. Bisa satu per satu atau semua sesi lain.
app.post('/api/sesi-aktif/putus', checkAuth('admin'), async (req, res) => {
  if (!req.session.user.master) {
    return res.status(403).json({ ok: false, message: 'Hanya admin master yang bisa memutus sesi.' });
  }
  try {
    const semua = !!(req.body && req.body.semua);
    const minta = String((req.body && req.body.sid) || '');
    const rows = await db.all('SELECT id, data FROM sessions');
    const target = [];
    for (const r of rows) {
      if (String(r.id) === String(req.sessionID)) continue;   // jangan putus sesi sendiri
      if (!semua && String(r.id) !== minta) continue;
      let s;
      try { s = JSON.parse(r.data || '{}'); } catch (e) { continue; }
      if (!s || !s.user) continue;
      target.push(String(r.id));
    }
    if (!target.length) {
      return res.status(404).json({ ok: false, message: 'Sesi tidak ditemukan (mungkin sudah berakhir).' });
    }
    await Promise.all(target.map(sid => new Promise((resolve, reject) => {
      req.sessionStore.destroy(sid, err => err ? reject(err) : resolve());
    })));
    res.json({ ok: true, diputus: target.length, message: target.length + ' sesi berhasil diputus.' });
  } catch (e) {
    console.error('[sesi-aktif/putus]', e);
    res.status(500).json({ ok: false, message: 'Gagal memutus sesi.' });
  }
});

// Ambil waktu aktivitas terakhir sesi (dari data login atau updated_at tabel).
function akhirSesi(r, dataSesi) {
  let upd = 0;
  if (r.updated_at) {
    const t = Date.parse(String(r.updated_at).replace(' ', 'T'));
    if (!isNaN(t)) upd = t;
  }
  return Math.max((dataSesi.user && dataSesi.user.akhir) || 0, upd);
}

function sesiKedaluwarsa(dataSesi, now) {
  if (!dataSesi.cookie || !dataSesi.cookie.expires) return false;
  const exp = new Date(dataSesi.cookie.expires).getTime();
  return !!exp && exp < now;
}

// HAPUS SESI DIAM — khusus admin master. Berbeda dengan "putus": sesi yang
// masih aktif tidak boleh dihapus lewat sini (pakai tombol Putus), dan sesi
// sendiri dikecualikan agar master tidak mengeluarkan dirinya sendiri.
app.post('/api/sesi-aktif/hapus', checkAuth('admin'), async (req, res) => {
  if (!req.session.user.master) {
    return res.status(403).json({ ok: false, message: 'Hanya admin master yang bisa menghapus sesi.' });
  }
  try {
    const semua = !!(req.body && (req.body.semua || req.body.diam));
    const minta = String((req.body && req.body.sid) || '');
    const now = Date.now();
    const rows = await db.all('SELECT id, data, updated_at FROM sessions');
    const target = [];
    for (const r of rows) {
      if (String(r.id) === String(req.sessionID)) continue;   // jangan hapus sesi sendiri
      let s;
      try { s = JSON.parse(r.data || '{}'); } catch (e) { s = null; }
      if (!s || !s.user) continue;
      const aktif = (now - akhirSesi(r, s)) < SESI_ONLINE_MS && !sesiKedaluwarsa(s, now);
      if (semua) {
        if (aktif) continue;                                  // hanya bersihkan yang Diam
      } else {
        if (String(r.id) !== minta) continue;
        if (aktif) {
          return res.status(409).json({ ok: false, message: 'Sesi ini masih aktif. Gunakan tombol Putus.' });
        }
      }
      target.push(String(r.id));
    }
    if (!target.length) {
      return res.status(404).json({ ok: false, message: 'Tidak ada sesi Diam yang bisa dihapus.' });
    }
    await Promise.all(target.map(sid => new Promise((resolve, reject) => {
      req.sessionStore.destroy(sid, err => err ? reject(err) : resolve());
    })));
    res.json({
      ok: true, dihapus: target.length,
      message: target.length + ' sesi Diam berhasil dihapus dari daftar.'
    });
  } catch (e) {
    console.error('[sesi-aktif/hapus]', e);
    res.status(500).json({ ok: false, message: 'Gagal menghapus sesi.' });
  }
});

// Pembersih otomatis: buang baris sesi yang cookie-nya sudah kedaluwarsa atau
// tidak ada aktivitasnya selama > 24 jam. Dipanggil saat start dan tiap jam
// supaya tabel sessions tidak menumpuk walaupun tidak ada yang menekan tombol.
async function bersihkanSesiLama() {
  const now = Date.now();
  const rows = await db.all('SELECT id, data, updated_at FROM sessions');
  const buang = [];
  for (const r of rows) {
    let s;
    try { s = JSON.parse(r.data || '{}'); } catch (e) { buang.push(String(r.id)); continue; }
    const kadaluwarsa = sesiKedaluwarsa(s, now);
    const akhir = akhirSesi(r, s || {});
    const terlaluLama = akhir && (now - akhir) > SESI_USIA_MAKS_MS;
    if (kadaluwarsa || terlaluLama || !s || !s.user) buang.push(String(r.id));
  }
  for (const sid of buang) await db.run('DELETE FROM sessions WHERE id = ?', [sid]);
  return buang.length;
}

app.get('/api/users', checkAuth('admin'), checkPengguna, async (req, res) => {
  try {
    const users = await db.all('SELECT id, username, role, tematik FROM users ORDER BY role DESC, username COLLATE NOCASE');

    // Lengkapi informasi tiap akun dengan kondisi sesi login aktifnya.
    // Password/tidak pernah dikirim ke klien.
    const sesiRows = await db.all('SELECT data FROM sessions');
    const now = Date.now();
    const agregat = new Map();
    for (const r of sesiRows) {
      let s;
      try { s = JSON.parse(r.data || '{}'); } catch (e) { continue; }
      const su = s && s.user;
      if (!su) continue;
      if (s.cookie && s.cookie.expires && new Date(s.cookie.expires).getTime() < now) continue;
      const kunci = String(su.username || '').toLowerCase();
      const a = agregat.get(kunci) || { sesi: 0, online: 0, terakhir: 0, master: false };
      a.sesi++;
      if ((now - (su.akhir || 0)) < 5 * 60000) a.online++;
      a.terakhir = Math.max(a.terakhir, su.akhir || 0);
      if (su.master) a.master = true;
      agregat.set(kunci, a);
    }

    res.json({
      success: true,
      users: users.map(function (u) {
        const a = agregat.get(String(u.username || '').toLowerCase()) ||
          { sesi: 0, online: 0, terakhir: 0, master: false };
        return {
          id: u.id, username: u.username, role: u.role, tematik: bacaTematikUser(u),
          sesi: a.sesi, online: a.online,
          terakhir: a.terakhir, terakhir_label: relatifWaktu(a.terakhir),
          master: a.master
        };
      })
    });
  } catch (e) {
    console.error('[users] gagal memuat:', e);
    res.status(500).json({ success: false, message: 'Gagal memuat daftar pengguna.' });
  }
});

// Daftar tematik yang bisa dipilih admin saat mengaitkan akun ke tematik
app.get('/api/tematik', checkAuth('admin'), checkPengguna, async (req, res) => {
  try {
    res.json({ success: true, tematik: await daftarTematikSistem() });
  } catch (e) {
    console.error('[tematik] gagal memuat:', e);
    res.status(500).json({ success: false, message: 'Gagal memuat daftar tematik.' });
  }
});

// Rapikan pilihan tematik dari klien: terima array atau "A,B" → array nama valid.
// Nilai yang tidak ada di sistem (typo/tematik lama) ditolak agar akses tidak
// diam-diam menjadi kosong karena salah tulis.
async function normalisasiPilihanTematik(masukan) {
  let arr = [];
  if (Array.isArray(masukan)) arr = masukan;
  else if (typeof masukan === 'string' && masukan.trim()) {
    const s = masukan.trim();
    if (s.charAt(0) === '[') {
      try { const a = JSON.parse(s); if (Array.isArray(a)) arr = a; } catch (e) { arr = []; }
    }
    if (!arr.length) arr = s.split(',');
  }
  arr = Array.from(new Set(arr.map(x => String(x).trim()).filter(Boolean)));
  if (!arr.length) return { ok: true, nilai: [] };

  const sistem = await daftarTematikSistem();
  const peta = new Map(sistem.map(t => [kunciTematik(t), t]));
  const hasil = [];
  const salah = [];
  for (const t of arr) {
    const cocok = peta.get(kunciTematik(t));
    if (cocok) { if (hasil.indexOf(cocok) < 0) hasil.push(cocok); } else salah.push(t);
  }
  if (salah.length) {
    return {
      ok: false,
      message: 'Tematik tidak dikenal: ' + salah.join(', ') +
               '. Pilih dari daftar tematik yang tersedia.'
    };
  }
  return { ok: true, nilai: hasil };
}

// Daftarkan akun lapangan baru — tanpa password, cukup username + tematik tugas
app.post('/api/users', checkAuth('admin'), checkPengguna, async (req, res) => {
  try {
    const username = String(req.body && req.body.username || '').trim();
    const salah = cekFormatUsername(username);
    if (salah) return res.status(400).json({ success: false, message: salah });

    const pilihan = await normalisasiPilihanTematik(req.body && req.body.tematik);
    if (!pilihan.ok) return res.status(400).json({ success: false, message: pilihan.message });

    const ada = await db.get('SELECT id FROM users WHERE username = ?', [username]);
    if (ada) return res.status(400).json({ success: false, message: 'Username sudah terdaftar.' });

    // hash acak: kolom password wajib terisi, tetapi tidak pernah dipakai saat login
    const hash = await bcrypt.hash(username + ':' + Date.now(), 10);
    await db.run('INSERT INTO users (username, password, role, tematik) VALUES (?, ?, ?, ?)',
      [username, hash, 'user', JSON.stringify(pilihan.nilai)]);

    const baris = await db.get('SELECT id, username, role, tematik FROM users WHERE username = ?', [username]);
    const pesan = pilihan.nilai.length
      ? 'Akun lapangan "' + username + '" terdaftar (tematik: ' + pilihan.nilai.join(', ') + ').'
      : 'Akun lapangan "' + username + '" terdaftar — BELUM ada tematik, jadi belum bisa mengakses data. Atur tematiknya terlebih dahulu.';
    res.json({
      success: true,
      user: { id: baris.id, username: baris.username, role: baris.role, tematik: pilihan.nilai },
      message: pesan
    });
  } catch (e) {
    console.error('[users] gagal menambah:', e);
    res.status(500).json({ success: false, message: 'Gagal mendaftarkan pengguna. ' + (e && e.message || '') });
  }
});

// Ubah username &/atau tematik tugas akun lapangan
app.put('/api/users/:id', checkAuth('admin'), checkPengguna, async (req, res) => {
  try {
    const id = Number(req.params.id);
    if (!Number.isInteger(id) || id <= 0) return res.status(400).json({ success: false, message: 'ID pengguna tidak valid.' });

    const lama = await db.get('SELECT id, username, role, tematik FROM users WHERE id = ?', [id]);
    if (!lama) return res.status(404).json({ success: false, message: 'Pengguna tidak ditemukan.' });
    if (lama.role === 'admin') return res.status(400).json({ success: false, message: 'Akun administrator tidak dapat diubah.' });

    const username = String(req.body && req.body.username || '').trim();
    const salah = cekFormatUsername(username);
    if (salah) return res.status(400).json({ success: false, message: salah });

    // Tematik hanya diubah bila klien benar-benar mengirimkannya
    const kirimTematik = !!(req.body && req.body.tematik !== undefined);
    let tematikBaru = bacaTematikUser(lama);
    if (kirimTematik) {
      const pilihan = await normalisasiPilihanTematik(req.body.tematik);
      if (!pilihan.ok) return res.status(400).json({ success: false, message: pilihan.message });
      tematikBaru = pilihan.nilai;
    }

    const bentrok = await db.get('SELECT id FROM users WHERE username = ? AND id <> ?', [username, id]);
    if (bentrok) return res.status(400).json({ success: false, message: 'Username sudah dipakai pengguna lain.' });

    await db.run('UPDATE users SET username = ?, tematik = ? WHERE id = ?',
      [username, JSON.stringify(tematikBaru), id]);

    // Sesi lama membawa tematik versi lama → buang supaya pembatasan langsung berlaku
    const berubah = username !== lama.username ||
      JSON.stringify(tematikBaru) !== JSON.stringify(bacaTematikUser(lama));
    if (berubah) await buangSesiUser(function (u) { return Number(u.id) === id; });

    const pesan = kirimTematik
      ? 'Akses "' + username + '" diperbarui → ' +
        (tematikBaru.length ? tematikBaru.join(', ') : 'tanpa tematik (belum bisa mengakses data)') + '.'
      : 'Username diperbarui menjadi "' + username + '".';

    res.json({
      success: true,
      user: { id: id, username: username, role: lama.role, tematik: tematikBaru },
      message: pesan
    });
  } catch (e) {
    console.error('[users] gagal mengubah:', e);
    res.status(500).json({ success: false, message: 'Gagal mengubah pengguna. ' + (e && e.message || '') });
  }
});

// Hapus akun lapangan
app.delete('/api/users/:id', checkAuth('admin'), checkPengguna, async (req, res) => {
  try {
    const id = Number(req.params.id);
    if (!Number.isInteger(id) || id <= 0) return res.status(400).json({ success: false, message: 'ID pengguna tidak valid.' });

    const lama = await db.get('SELECT id, username, role FROM users WHERE id = ?', [id]);
    if (!lama) return res.status(404).json({ success: false, message: 'Pengguna tidak ditemukan.' });
    if (lama.role === 'admin') return res.status(400).json({ success: false, message: 'Akun administrator tidak dapat dihapus.' });
    if (req.session.user && Number(req.session.user.id) === id) {
      return res.status(400).json({ success: false, message: 'Tidak bisa menghapus akun yang sedang dipakai untuk login.' });
    }

    await db.run('DELETE FROM users WHERE id = ?', [id]);
    await buangSesiUser(function (u) { return Number(u.id) === id; });

    res.json({ success: true, message: 'Akun "' + lama.username + '" dihapus.' });
  } catch (e) {
    console.error('[users] gagal menghapus:', e);
    res.status(500).json({ success: false, message: 'Gagal menghapus pengguna. ' + (e && e.message || '') });
  }
});



// Tentukan master harga untuk generate BA/BOQ.
// Prioritas: query ?sumber_master= → sumber tersimpan di rab_rekon → amd4.
// Ini menyamakan angka "hitung realisasi" dengan angka saat download,
// termasuk download dari menu Admin yang tidak mengirim query.
async function sumberMasterUntukRekon(lokasi_id, dariQuery) {
  const q = String(dariQuery == null ? '' : dariQuery).toLowerCase();
  if (q === 'amd5' || q === 'amd4') return q;
  try {
    const row = await db.get(
      'SELECT sumber_master FROM rab_rekon WHERE lokasi_id = ? LIMIT 1', [lokasi_id]
    );
    if (row && String(row.sumber_master).toLowerCase() === 'amd5') return 'amd5';
  } catch (e) { /* kolom belum ada → pakai default */ }
  return 'amd4';
}

// ─────────────────────────────────────────────────────────────────────────────
// [API] HARGA SATUAN DARI BOQ TAMKUR
// Dipakai DRM/REKON untuk mengalikan volume (M- → G material, J- → H jasa).
// ─────────────────────────────────────────────────────────────────────────────
app.get('/api/boq-tamkur/harga', checkAuth('admin'), (req, res) => {
  try {
    // DEFAULT 'amd4' = alur lama. 'amd5' → BOQ TAMKUR AMD 5 PAKET 5.xlsx.
    const sumberMaster = String(req.query.sumber_master || 'amd4').toLowerCase() === 'amd5'
      ? 'amd5' : 'amd4';
    const harga = bacaHargaBoqTamkur(sumberMaster);
    res.json({
      success: true, harga, jumlah: Object.keys(harga).length,
      sumber_master: sumberMaster,
      sheet_master: sumberMaster === 'amd5' ? 'AMD 5 Paket5' : 'BOQ TAMKUR'
    });
  } catch (e) {
    res.status(500).json({ message: 'Gagal membaca harga BOQ TAMKUR.', detail: e.message });
  }
});

// ─────────────────────────────────────────────────────────────────────────────
// [API] CEK DATA MASTER BOQ DRM ALL (baris 4 = nama lokasi)
// Endpoint BARU — tidak mengubah upload Excel / hitungan / endpoint lama.
// Master hanya dibaca; bila lokasi tidak ada, frontend tetap mengizinkan upload.
// ─────────────────────────────────────────────────────────────────────────────
app.get('/api/boq-drm-all/cek', checkAuth('admin'), async (req, res) => {
  try {
    const nama = String(req.query.nama_lokasi || '').trim();
    if (!nama) return res.status(400).json({ message: 'Nama lokasi belum dipilih.' });
    // Google Sheet / BOQ DRM ALL = SUMBER VOLUME (untuk ditampilkan).
    // Harga jasa/material = BOQ TAMKUR AMD4 / AMD5 (database harga).
    // Keduanya dipakai sesuai pemilih VLOOKUP di toolbar.
    const sumberMaster = String(req.query.sumber_master || 'amd4').toLowerCase() === 'amd5'
      ? 'amd5' : 'amd4';
    const hasil = await cariLokasiDiMasterDrm(nama);
    hasil.sumber_master = sumberMaster;

    const harga = bacaHargaBoqTamkur(sumberMaster);
    // Kunci cadangan: spasi/tanda baca di designator master (mis.
    // "J-PRELIMINARY PROJECT" vs "J-PRELIMINARY-PROJECT") tidak boleh
    // membuat harga gagal ditemukan. Tidak ada tabrakan harga di kedua master.
    const hargaKunci = {};
    Object.keys(harga).forEach(k => {
      const kunci = String(k).toUpperCase().replace(/[^A-Z0-9]/g, '');
      if (kunci && hargaKunci[kunci] == null) hargaKunci[kunci] = harga[k];
    });

    ['drm', 'rekon'].forEach(kunci => {
      const bagian = hasil[kunci];
      if (!bagian || !Array.isArray(bagian.items)) return;
      bagian.items = bagian.items.map(it => {
        const asli = String(it.designator || '').toUpperCase().trim();
        const kunciDes = asli.replace(/[^A-Z0-9]/g, '');
        const h = typeof harga[asli] === 'number' ? harga[asli] : hargaKunci[kunciDes];
        const salinan = Object.assign({}, it);
        if (/^J-/.test(salinan.designator)) salinan.jasa  = typeof h === 'number' ? h : 0;
        else if (/^M-/.test(salinan.designator)) salinan.material = typeof h === 'number' ? h : 0;
        return salinan;
      });
    });
    hasil.sumber_harga = 'file';
    hasil.sheet_harga = sumberMaster === 'amd5' ? 'AMD 5 Paket5' : 'BOQ AMD4 PAKET 5';
    res.json({ success: true, ...hasil });
  } catch (e) {
    res.status(500).json({ message: 'Gagal membaca data master BOQ DRM ALL.', detail: e.message });
  }
});

// ─────────────────────────────────────────────────────────────────────────────
// [API] UPLOAD EXCEL REKON
// Hanya mengambil designator material (M-), designator jasa (J-), dan nilai
// REKON. Harga TIDAK diambil — tetap memakai BOQ TAMKUR saat generate BA.
// ─────────────────────────────────────────────────────────────────────────────
app.post('/api/upload-rekon-excel', checkAuth('admin'), upload.single('file'), (req, res) => {
  if (!req.file) return res.status(400).json({ message: 'Pilih file Excel REKON terlebih dahulu.' });
  try {
    const hasil = bacaRekonDariExcel(req.file.buffer);
    if (!hasil.items.length) {
      return res.status(400).json({
        message: 'Tidak ada designator material/jasa yang ditemukan.',
        sheet: hasil.sheet, laporan: hasil.laporan
      });
    }
    res.json({ success: true, ...hasil });
  } catch (e) {
    res.status(500).json({ message: 'Gagal membaca file Excel.', detail: e.message });
  }
});

// ─────────────────────────────────────────────────────────────────────────────
// [API] UPLOAD EXCEL DRM
// Sama seperti REKON, hanya kolom nilainya "TOTAL VOLUME" (kolom I pada
// BOQ RELOKASI METLAND 2026). Harga tidak diambil.
// ─────────────────────────────────────────────────────────────────────────────
app.post('/api/upload-drm-excel', checkAuth('admin'), upload.single('file'), (req, res) => {
  if (!req.file) return res.status(400).json({ message: 'Pilih file Excel DRM terlebih dahulu.' });
  try {
    const hasil = bacaDrmDariExcel(req.file.buffer);
    if (!hasil.items.length) {
      return res.status(400).json({
        message: 'Tidak ada designator material/jasa yang ditemukan.',
        sheet: hasil.sheet, laporan: hasil.laporan
      });
    }
    res.json({ success: true, ...hasil });
  } catch (e) {
    res.status(500).json({ message: 'Gagal membaca file Excel.', detail: e.message });
  }
});

// Endpoint: GET /api/generate/ba_tambahkurang/:lokasi_id
// Template : templates/template_BA_TambahKurang.docx
//
// Placeholder yang tersedia di template Word:
//  Identitas   : {lokasi} {pid} {tematik} {proyek} {kontrak} {surat_pesan}
//                {pelaksana} {nama_waspang} {nik_waspang} {tgl_generate}
//  Tabel baris : {tk_no1}..{tk_no50}  {tk_des1}..{tk_des50}
//                {tk_sat1}..{tk_sat50} {tk_drm1}..{tk_drm50}
//                {tk_rek1}..{tk_rek50} {tk_plus1}..{tk_plus50}
//                {tk_min1}..{tk_min50} {tk_ket1}..{tk_ket50}
//  Summary     : {tk_total_plus} {tk_total_min} {tk_jumlah_item}
//                {tk_jumlah_tambah} {tk_jumlah_kurang}
// ─────────────────────────────────────────────────────────────────────────────
// Identitas untuk placeholder {{AREA}} / {{LOKASI}} di BOQ TAMKUR.xlsx.
// AREA diambil dari branch di detail_tambahan (contoh: BOGOR), bukan tematik.
// ─────────────────────────────────────────────────────────────────────────────
function areaDariLokasi(lokasiData, laporan) {
  let detail = lokasiData && lokasiData.detail_tambahan;
  if (typeof detail === 'string') {
    try { detail = JSON.parse(detail); } catch (e) { detail = {}; }
  }
  detail = detail || {};
  const l = laporan || {};
  return String(
    detail.branch || detail.area || detail.wilayah || detail.pm_area ||
    l.pm_area || l.district ||
    (lokasiData && lokasiData.tematik) || '-'
  ).trim();
}
function namaDariLokasi(lokasiData) {
  return String((lokasiData && (lokasiData.nama_lokasi || lokasiData.pid)) || '-').trim();
}

// ─────────────────────────────────────────────────────────────────────────────
// [API] GENERATE BA TAMBAH KURANG — string-replace langsung di XML
// ─────────────────────────────────────────────────────────────────────────────
app.get('/api/generate/ba_tambahkurang/:lokasi_id', checkAuth('admin'), async (req, res) => {
  try {
    const { lokasi_id } = req.params;
    const templatePath  = path.join(__dirname, 'templates', 'templateTambahKurang.docx');
    // Harga BA ikut pemilih VLOOKUP: amd4 = BOQ TAMKUR, amd5 = BOQ AMD 5 PAKET 5.
    // Bila download dari menu Admin (tanpa query), pakai sumber tersimpan.
    const sumberMaster = await sumberMasterUntukRekon(lokasi_id, req.query.sumber_master);

    if (!fs.existsSync(templatePath)) {
      return res.status(404).json({ message: 'Template templateTambahKurang.docx tidak ditemukan di folder /templates/' });
    }

    const lokasiData = await db.get('SELECT * FROM lokasi WHERE id = ?', [lokasi_id]);
    if (!lokasiData) return res.status(404).json({ message: 'Data lokasi tidak ditemukan.' });

    const adminInput = await db.get('SELECT * FROM laporan_rekon WHERE lokasi_id = ?', [lokasi_id]) || {};
    
    // ── AMBIL DATA REKON DAN HARGA SATUAN DARI RAB (JOIN TABEL) ──
   // GANTI query rekonItems dengan ini:
const rekonItems = await db.all(`
  SELECT 
    rk.designator,
    rk.satuan,
    rk.drm       AS vol_drm,
    rk.rekon     AS vol_rekon,
    rk.tambah    AS vol_tambah,
    rk.kurang    AS vol_kurang,
    rk.harga_satuan AS harga_tersimpan,
    COALESCE(
      (
        SELECT (r2.material + r2.jasa)
        FROM rab r2
        WHERE UPPER(TRIM(r2.designator)) = UPPER(TRIM(rk.designator))
          AND r2.lokasi_id = rk.lokasi_id
        LIMIT 1
      ),
      (
        SELECT (r3.material + r3.jasa)
        FROM rab r3
        WHERE UPPER(TRIM(r3.designator)) = UPPER(TRIM(rk.designator))
        LIMIT 1
      ),
      0
    ) AS harga_dari_rab
  FROM rab_rekon rk
  WHERE rk.lokasi_id = ?
`, [lokasi_id]);

    // Harga satuan WAJIB dari BOQ TAMKUR sesuai master yang dipilih.
    // Kunci kanonik menyatukan ejaan "J-PRELIMINARY PROJECT" /
    // "J-PRELIMINARY-PROJECT" tanpa mengubah teks designator asli.
    const hargaBoq = bacaHargaBoqTamkur(sumberMaster);
    const hargaBoqKunci = {};
    Object.keys(hargaBoq).forEach(k => {
      const kunci = String(k).toUpperCase().replace(/[^A-Z0-9]/g, '');
      if (kunci && hargaBoqKunci[kunci] == null) hargaBoqKunci[kunci] = hargaBoq[k];
    });

console.log('DEBUG rekonItems:', JSON.stringify(rekonItems.slice(0,3), null, 2));
    let tkRows = [];
    let totalDrmRp = 0, totalRekRp = 0, totalPlusRp = 0, totalMinRp = 0;

    // GANTI SELURUH BAGIAN forEach ini:
rekonItems.forEach((item, i) => {
  const hSatRAB       = parseFloat(item.harga_dari_rab)  || 0;
  const hSatTersimpan = parseFloat(item.harga_tersimpan) || 0;
  const kunciDes = String(item.designator || '').toUpperCase().replace(/[^A-Z0-9]/g, '');
  const hSatBoq  = hargaBoqKunci[kunciDes];
  // Prioritas: BOQ TAMKUR (master terpilih) → RAB → harga tersimpan.
  // Google Doc / BOQ DRM ALL hanya memberi volume, bukan harga.
  const hSat = typeof hSatBoq === 'number' && hSatBoq > 0
    ? hSatBoq
    : (hSatRAB > 0 ? hSatRAB : hSatTersimpan);

  console.log(`[HSAT] ${item.designator} | boq=${typeof hSatBoq === 'number' ? hSatBoq : 0} | rab=${hSatRAB} | tersimpan=${hSatTersimpan} | pakai=${hSat} | sumber=${sumberMaster}`);

  const volDrm   = parseFloat(item.vol_drm)   || 0;
  const volRekon = parseFloat(item.vol_rekon)  || 0;
  const volPlus  = parseFloat(item.vol_tambah) || 0;
  const volMin   = parseFloat(item.vol_kurang) || 0;

  const nilaiDrm   = volDrm   * hSat;
  const nilaiRekon = volRekon * hSat;
  const nilaiPlus  = volPlus  * hSat;
  const nilaiMin   = volMin   * hSat;

  totalDrmRp  += nilaiDrm;
  totalRekRp  += nilaiRekon;
  totalPlusRp += nilaiPlus;
  totalMinRp  += nilaiMin;

  const ket    = volPlus > 0 ? 'TAMBAH' : (volMin > 0 ? 'KURANG' : 'TETAP');
  const fmtVol = n => n > 0 ? n.toLocaleString('id-ID', { maximumFractionDigits: 2 }) : '0';
  const fmtRp  = n => n > 0 ? n.toLocaleString('id-ID') : '0';

  tkRows.push({
    no:      String(i + 1),
    des:     item.designator || '-',
    drm:     fmtVol(volDrm),
    rek:     fmtVol(volRekon),
    plus:    fmtVol(volPlus),
    min:     fmtVol(volMin),
    drm_rp:  fmtRp(nilaiDrm),
    rek_rp:  fmtRp(nilaiRekon),
    plus_rp: fmtRp(nilaiPlus),
    min_rp:  fmtRp(nilaiMin),
    ket,
  });
});

    // ── Tulis DRM/REKON ke BOQ TAMKUR.xlsx ──────────────────────────────
    // Hanya untuk AMD4. File AMD5 adalah MASTER HARGA (read-only):
    // penulis lama tulisBoqTamkur() bisa merusak sel/price di file itu.
    // Template AMD5 tetap dipakai read-only oleh download BOQ.
    // Kegagalan sinkronisasi TIDAK mengubah hitungan / menggagalkan BA.
    try {
      if (sumberMaster === 'amd5') {
        console.log('[BOQ TAMKUR] sinkronisasi dilewati (AMD5 = master harga read-only)');
      } else {
        const sinkron = tulisBoqTamkur(rekonItems.map(it => ({
          designator: it.designator,
          drm:  parseFloat(it.vol_drm)   || 0,
          rekon: parseFloat(it.vol_rekon) || 0
        })), {
          lokasi: namaDariLokasi(lokasiData),
          area:   areaDariLokasi(lokasiData, adminInput),
          sumberMaster
        });
        console.log('[BOQ TAMKUR] sinkronisasi:', JSON.stringify(sinkron));
      }
    } catch (e) {
      console.warn('[BOQ TAMKUR] gagal sinkronisasi (diabaikan):', e.message);
    }

    const formatRupiah = (n) => typeof n === 'number' ? n.toLocaleString('id-ID') : '0';
// Nilai Deviasi = Rekon - DRM (negatif jika ada pekerjaan kurang)
const nilaiDeviasi = totalRekRp - totalDrmRp;

const values = {
  // ── Identitas Dokumen ──────────────────────────────────────
  proyektk:  adminInput.proyek        || '-',
  kontraktk: adminInput.kontrak       || '-',
  sptk:      adminInput.surat_pesanan || '-',
  nosptk:    adminInput.surat_pesanan || '-',
  loktk:     lokasiData.nama_lokasi   || '-',
  pelaksana: adminInput.pelaksana     || 'PT TELKOM AKSES',

  // ── Tanggal BA ─────────────────────────────────────────────
  haritk: adminInput.hari_ct          || '-',
  tgltk:  adminInput.tgl              || '-',
  blntk:  adminInput.bulan            || '-',
  thntk:  new Date().getFullYear().toString(),

  // ── Nilai Rekapitulasi (Rupiah) ────────────────────────────
  tk_total_drm:  formatRupiah(totalDrmRp),
  tk_total_rek:  formatRupiah(totalRekRp),
  tk_total_plus: formatRupiah(totalPlusRp),
  tk_total_min:  formatRupiah(totalMinRp),
  // Deviasi: tampil negatif jika kurang (pakai tanda minus eksplisit)
  tk_total_dev:  (nilaiDeviasi < 0 ? '-' : '') + formatRupiah(Math.abs(nilaiDeviasi)),

  // ── Alias Tambahan untuk Template Word ────────────────────
  nilai_drm:     formatRupiah(totalDrmRp),
  nilai_rekon:   formatRupiah(totalRekRp),
  nilai_tambah:  formatRupiah(totalPlusRp),
  nilai_kurang:  formatRupiah(totalMinRp),
  nilai_deviasi: (nilaiDeviasi < 0 ? '-' : '') + formatRupiah(Math.abs(nilaiDeviasi)),
};

    // Isi placeholder per baris tabel (max 19 baris sesuai template)
   const MAX_ROWS = 13;
for (let i = 1; i <= MAX_ROWS; i++) {
  const row = tkRows[i - 1];
  values['tk_no'     + i] = row ? String(i)    : '';
  values['tk_des'    + i] = row?.des            ?? '';
  values['tk_drm'    + i] = row?.drm            ?? '';
  values['tk_rek'    + i] = row?.rek            ?? '';
  values['tk_plus'   + i] = row?.plus           ?? '';
  values['tk_min'    + i] = row?.min            ?? '';
  values['tk_ket'    + i] = row?.ket            ?? '';
  // === NILAI RP PER BARIS (INI YANG HILANG) ===
  values['tk_drm_rp' + i] = row?.drm_rp        ?? '';
  values['tk_rek_rp' + i] = row?.rek_rp        ?? '';
  values['tk_plus_rp'+ i] = row?.plus_rp       ?? '';
  values['tk_min_rp' + i] = row?.min_rp        ?? '';
}
// Log verifikasi summary
console.log('[BA-TK] Summary →',
  '| DRM  :', formatRupiah(totalDrmRp),
  '| REK  :', formatRupiah(totalRekRp),
  '| PLUS :', formatRupiah(totalPlusRp),
  '| MIN  :', formatRupiah(totalMinRp),
);

    // ── Gunakan Docxtemplater agar XML Word tidak corrupt ──
    const zipBuf = fs.readFileSync(templatePath, 'binary');
    const zip    = new PizZip(zipBuf);
    
    const doc = new Docxtemplater(zip, {
      paragraphLoop: true,
      linebreaks: true,
      parser: (tag) => ({
        get: (scope) => {
          const cleanTag = tag.replace(/[\u200B-\u200D\uFEFF]/g, "").trim();
          return scope[cleanTag] !== undefined ? scope[cleanTag] : "";
        }
      })
    });

    doc.render(values);
    const outBuf0 = doc.getZip().generate({ type: 'nodebuffer', compression: 'DEFLATE' });
    const outBuf = await lampirkanPetaKeDocx(outBuf0, lokasi_id);

    const nlc      = lokasiData.nama_lokasi.replace(/[^a-zA-Z0-9_-]/g, '_');
    const namaFile = 'BA_TAMBAH_KURANG_' + nlc + '.docx';

    await catatAktivitas('ba_tambahkurang', lokasi_id);
    await kirimDokumen(req, res, outBuf, namaFile);

    console.log('BA TK generated: lokasi=' + lokasi_id + ', baris=' + tkRows.length);

  } catch (err) {
    console.error('Error generate BA Tambah Kurang:', err);
    res.status(500).json({ message: 'Gagal generate BA Tambah Kurang.', detail: err.message });
  }
});
// ─────────────────────────────────────────────────────────────────────────────
// [API] GENERATE EXCEL BOQ TAMBAH KURANG — HANYA BOQ
// Sumber : templates/BOQ TAMKUR.xlsx (amd4) atau
//          templates/BOQ TAMKUR AMD 5 PAKET 5.xlsx (amd5) — template tidak ditimpa.
// Isi    : K = DRM, L = REKON dari rab_rekon; I/J/M/N dihitung ulang.
// ─────────────────────────────────────────────────────────────────────────────
const { bangunBoqTamkur } = require('./services/boqTamkur');

app.get('/api/generate/boq_tambahkurang/:lokasi_id', checkAuth('admin'), async (req, res) => {
  try {
    const { lokasi_id } = req.params;
    const sumberMaster = await sumberMasterUntukRekon(lokasi_id, req.query.sumber_master);
    const lokasiData = await db.get(
      'SELECT nama_lokasi, pid, tematik, detail_tambahan FROM lokasi WHERE id = ?', [lokasi_id]
    );
    if (!lokasiData) return res.status(404).json({ message: 'Data lokasi tidak ditemukan.' });

    const laporan = await db.get(
      'SELECT pm_area, district FROM laporan_rekon WHERE lokasi_id = ?', [lokasi_id]
    ) || {};

    const rekonItems = await db.all(
      'SELECT designator, drm, rekon FROM rab_rekon WHERE lokasi_id = ?', [lokasi_id]
    );
    if (!rekonItems.length) {
      return res.status(400).json({ message: 'Belum ada data rekon tersimpan. Hitung & Simpan Rekon dulu.' });
    }

    const hasil = bangunBoqTamkur(rekonItems.map(it => ({
      designator: it.designator,
      drm:  parseFloat(it.drm)   || 0,
      rekon: parseFloat(it.rekon) || 0,
    })), {
      lokasi: namaDariLokasi(lokasiData),
      area:   areaDariLokasi(lokasiData, laporan),
      sumberMaster
    });

    if (!hasil.ok || !hasil.buffer) {
      return res.status(500).json({ message: 'Gagal menyusun BOQ: ' + (hasil.alasan || 'tidak diketahui') });
    }

    const nlc = String(lokasiData.nama_lokasi || 'lokasi').replace(/[^a-zA-Z0-9_-]/g, '_');
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename="BOQ_TAMBAH_KURANG_' + nlc + '.xlsx"');
    res.send(hasil.buffer);

    console.log('[BOQ TAMKUR] download Excel: lokasi=' + lokasi_id +
      ', sumber=' + sumberMaster + ', sheet=' + (hasil.sheet || '-') +
      ', ditulis=' + hasil.ditulis + ', dilewati=' + hasil.dilewati.length +
      ', identitas=' + JSON.stringify(hasil.identitas) +
      (hasil.dilewati.length ? ' (' + hasil.dilewati.slice(0, 3).join(', ') + ')' : ''));
  } catch (err) {
    console.error('Error generate BOQ Tambah Kurang:', err);
    res.status(500).json({ message: 'Gagal generate Excel BOQ Tambah Kurang.', detail: err.message });
  }
});
// ─────────────────────────────────────────────────────────────────────────────
// [API] Generate Dokumen — mapping jenis → template
// ─────────────────────────────────────────────────────────────────────────────
const TEMPLATE_MAP = {
  'lact':            'templateRelokasi_clean.docx',
  'ba_lapangan':     'template_BA_Lapangan.docx',
  'ba_perubahan':    'template_BA_Perubahan.docx',
  'ba_prelim':       'template_BA_Prelim.docx',
  'bact':            'template_BACT.docx',
  'baut':            'template_BAUT.docx',
  'sph':             'template_SPH.docx',
  'ba_tambahkurang': 'templateTambahKurang.docx',
  'microdemand':     'template_MICRODEMAND.docx',
  'hem':             'template_HEM.docx',
  'mom_redesign':    'template_MOM_Redesign.docx',
};

// ─────────────────────────────────────────────────────────────────────────────
// [API] GENERATE BAUT — POST dengan upload gambar SP & SSBOQ
// ─────────────────────────────────────────────────────────────────────────────
app.post('/api/generate/baut/:lokasi_id', checkAuth('admin'), upload.fields([
  { name: 'sp_file',    maxCount: 1 },
  { name: 'ssboq_file', maxCount: 1 },
]), async (req, res) => {
  try {
    const lokasi_id = req.params.lokasi_id;

    // ── Template ──────────────────────────────────────────────────────────────
    const templatePath = path.join(__dirname, 'templates', 'template_BAUT.docx');
    if (!fs.existsSync(templatePath)) {
      return res.status(404).json({
        message: 'Template "template_BAUT.docx" belum tersedia.',
        info: 'Letakkan file "template_BAUT.docx" di folder /templates/ pada server.',
        template: 'template_BAUT.docx',
      });
    }

    // ── Data lokasi & berkas ─────────────────────────────────────────────────
    const lokasiData  = await db.get('SELECT * FROM lokasi WHERE id = ?', [lokasi_id]);
    if (!lokasiData)  return res.status(404).json({ message: 'Data lokasi tidak ditemukan.' });

    const semuaBerkas = await db.all(
      'SELECT file_name, kategori, sub_kategori FROM berkas WHERE lokasi_id = ? ORDER BY id ASC',
      [lokasi_id]
    );

    const ntc = (lokasiData.tematik || 'Umum').replace(/[^a-zA-Z0-9_-]/g, '_');
    const nlc = lokasiData.nama_lokasi.replace(/[^a-zA-Z0-9_-]/g, '_');
    const blankImage = path.join(__dirname, 'public', 'images', 'blank.png');

    // ── Input dari form ──────────────────────────────────────────────────────
    const proyekut       = (req.body.proyekut       || '-').trim();
const kontrakut      = (req.body.kontrakut      || '-').trim();
const surat_pesanut  = (req.body.surat_pesanut  || '-').trim();
const districtut     = (req.body.districtut     || '-').trim();
const lokasi_smileut = (req.body.lokasi_smileut || lokasiData.nama_lokasi || '-').trim();
// Field ABD
const hariabd    = (req.body.hariabd    || '-').trim();
const tglabd     = (req.body.tglabd     || '-').trim();
const districtabd= (req.body.districtabd|| '-').trim();
const stoabd     = (req.body.stoabd     || '-').trim();
const lokasiabd  = (req.body.lokasiabd  || lokasiData.nama_lokasi || '-').trim();
    // ── Simpan & proses gambar SP (Surat Penunjukan) ─────────────────────────
    // Folder: uploads/{tematik}/{lokasi}/-SURAT_PENUNJUKAN/
    const folderSP = path.join(__dirname, 'uploads', ntc, nlc, '-SURAT_PENUNJUKAN');
    pastikanFolderEksis(folderSP);

    let pathSP = blankImage;
    if (req.files && req.files['sp_file'] && req.files['sp_file'][0]) {
      const spBuf = req.files['sp_file'][0].buffer;
      const spDest = path.join(folderSP, `SURATPENUNJUKAN_${nlc}.jpg`);
      await sharp(spBuf)
        .resize({ width: 2480, height: 3508, fit: 'inside', withoutEnlargement: true })
        .jpeg({ quality: 92, chromaSubsampling: '4:4:4' })
        .toFile(spDest);
      pathSP = spDest;
    }

    // ── Simpan & proses gambar SSBOQ ─────────────────────────────────────────
    // Folder: uploads/{tematik}/{lokasi}/-SSBOQ/
    const folderSSBOQ = path.join(__dirname, 'uploads', ntc, nlc, '-SSBOQ');
    pastikanFolderEksis(folderSSBOQ);

    let pathSSBOQ = blankImage;
    if (req.files && req.files['ssboq_file'] && req.files['ssboq_file'][0]) {
      const ssboqBuf = req.files['ssboq_file'][0].buffer;
      const ssboqDest = path.join(folderSSBOQ, `SSBOQ_${nlc}.jpg`);
      await sharp(ssboqBuf)
        .resize({ width: 2480, height: 1754, fit: 'inside', withoutEnlargement: true })
        .jpeg({ quality: 92, chromaSubsampling: '4:4:4' })
        .toFile(ssboqDest);
      pathSSBOQ = ssboqDest;
    }

    // ── Ambil EOPM1 & LDP1 dari berkas yang sudah ada (sama seperti LACT) ────
    function ambilFotoBAUT(kode, slotKe) {
      const KM_BAUT = {
        'EOPM': '-EVIDENT HASIL UKUR OPM',
        'LDP':  '-LAMPIRAN DATA PENGUKURAN OPM',
      };
      const namaKat = KM_BAUT[kode];
      if (!namaKat) return blankImage;
      const list   = semuaBerkas.filter(b => b.kategori === namaKat);
      const target = list[slotKe - 1];
      if (target) {
        const fp = path.join(__dirname, 'uploads', ntc, nlc,
          target.kategori.replace(/[^a-zA-Z0-9_-]/g, '_'), target.file_name);
        return fs.existsSync(fp) ? fp : blankImage;
      }
      return blankImage;
    }

    // ── ImageModule ───────────────────────────────────────────────────────────
    // SURATPENUNJUKAN: hampir penuh kertas A4 → ~490×680 px (EMU ~= 490*9144/100)
    // SSBOQ: setengah kertas → ~490×337 px
    const TAG_BESAR_BAUT = ['SURATPENUNJUKAN', 'SSBOQ'];
    const imageModuleBaut = new ImageModule({
      getImage: tv => fs.readFileSync((tv && fs.existsSync(tv)) ? tv : blankImage),
      getSize: (img, tv, tn) => {
        const base = tn.replace(/\d+$/, '');
        if (base === 'SURATPENUNJUKAN') return [490, 680]; // hampir penuh kertas
        if (base === 'SSBOQ')           return [490, 337]; // ±setengah kertas
        if (['LDP','MCR','LKDL'].includes(base)) return [450, 320];
        if (base === 'foto') {
          // Foto OPM: sesuaikan dengan kotak grid — rasio asli terjaga,
          // muat dalam batas maks 250×180 (sama seperti LACT).
          try {
            const d = sizeOf(img);
            if (!d || !d.width || !d.height) return [250, 150];
            const maxW = 250, maxH = 180;
            const ratio = Math.min(maxW / d.width, maxH / d.height);
            return [Math.round(d.width * ratio), Math.round(d.height * ratio)];
          } catch { return [250, 150]; }
        }
        return [250, 150];
      },
    });

    // ── Payload ───────────────────────────────────────────────────────────────
    const payload = {
  // Data form utama
  proyekut,
  kontrakut,
  surat_pesanut,
  districtut,
  lokasi_smileut,
  // Alias LACT-style untuk blok loop {#opm_list} (header PROYEK/KONTRAK/DLL)
  proyek:       proyekut,
  kontrak:      kontrakut,
  surat_pesan:  surat_pesanut,
  district:     districtut,
  lokasi_smile: lokasi_smileut,
  // Data ABD
  hariabd,
  tglabd,
  districtabd,
  stoabd,
  lokasiabd,
  // Info umum
  lokasi:      lokasiData.nama_lokasi || '-',
  tematik:     lokasiData.tematik    || '-',
  pid:         lokasiData.pid        || '-',
  tgl_generate: new Date().toLocaleDateString('id-ID', {
    weekday: 'long', day: '2-digit', month: 'long', year: 'numeric'
  }),
  // Gambar upload
  SURATPENUNJUKAN: pathSP,
  SSBOQ:           pathSSBOQ,
  // LDP semua slot (Lampiran Data Pengukuran OPM)
  LDP1:   ambilFotoBAUT('LDP',  1),
  LDP2:   ambilFotoBAUT('LDP',  2),
  LDP3:   ambilFotoBAUT('LDP',  3),
  LDP4:   ambilFotoBAUT('LDP',  4),
  // Loop OPM (Hasil Ukur) — grid per ODP/port otomatis & unlimited,
  // persis seperti LACT. Data disiapkan oleh siapkanDataOpm().
  opm_list: await siapkanDataOpm(lokasi_id, path.join(__dirname, 'uploads', ntc, nlc, '-EVIDENT_HASIL_UKUR_OPM'), '-EVIDENT HASIL UKUR OPM', 'opm', {
    proyek:       payload.proyek        || '-',
    kontrak:      payload.kontrak       || '-',
    surat_pesan:  payload.surat_pesan   || '-',
    district:     payload.district      || '-',
    lokasi_smile: payload.lokasi_smile  || '-',
  }),
};
    // ── Render dokumen ────────────────────────────────────────────────────────
    const doc = new Docxtemplater(
      new PizZip(fs.readFileSync(templatePath, 'binary')),
      {
        modules: [imageModuleBaut],
        paragraphLoop: true,
        linebreaks: true,
        parser: tag => ({ get: scope => {
          const ct = tag.replace(/[\u200B-\u200D\uFEFF]/g, '').trim();
          return scope[ct] !== undefined ? scope[ct] : '';
        }}),
      }
    );

    doc.render(payload);
    const buf0 = doc.getZip().generate({ type: 'nodebuffer', compression: 'DEFLATE' });
    const buf = await lampirkanPetaKeDocx(buf0, lokasi_id);

    const namaFile = 'BAUT_' + nlc + '.docx';
    await catatAktivitas('baut', lokasi_id);
    await kirimDokumen(req, res, buf, namaFile);

    console.log('BAUT generated: lokasi=' + lokasi_id + ', nama=' + lokasiData.nama_lokasi);

  } catch (err) {
    console.error('Error generate BAUT:', err);
    res.status(500).json({ message: 'Gagal generate BAUT.', detail: err.message });
  }
});

// ─────────────────────────────────────────────────────────────────────────────
// [API] UPLOAD 5 DOKUMEN BAUT-1 NODE B — dipanggil modal SEBELUM generate.
// Gambar → JPG; PDF multi-halaman → PNG per halaman (renderPdfHalaman).
// Hasil dibaca ulang oleh GET /api/download-docx/:id?baut=1 (mode BAUT-1).
// ─────────────────────────────────────────────────────────────────────────────
const BAUT1_DOKUMEN = [
  ['nota_dinas',        'BAUT1_NOTA_DINAS'],
  ['permohonan_tim_ut', 'BAUT1_PERMOHONAN_TIM_UT'],
  ['penunjukan_tim_ut', 'BAUT1_PENUNJUKAN_TIM_UT'],
  ['form_ceklis',       'BAUT1_FORM_CEKLIS'],
  ['ba_prelim',         'BAUT1_BA_PRELIM'],
];
app.post('/api/baut-nodeb/upload/:lokasi_id', checkAuth('admin'),
  upload.fields(BAUT1_DOKUMEN.map(([k]) => ({ name: k, maxCount: 20 }))),
  async (req, res) => {
    try {
      const lokasi_id = req.params.lokasi_id;
      const lok = await db.get('SELECT * FROM lokasi WHERE id = ?', [lokasi_id]);
      if (!lok) return res.status(404).json({ message: 'Data lokasi tidak ditemukan.' });
      if (lok.tematik !== 'Node B') {
        return res.status(400).json({ message: 'Upload dokumen BAUT-1 hanya untuk tematik Node B.' });
      }
      const ntc = (lok.tematik || 'Umum').replace(/[^a-zA-Z0-9_-]/g, '_');
      const nlc = lok.nama_lokasi.replace(/[^a-zA-Z0-9_-]/g, '_');
      const hasil = {};
      let urut = 0;
      for (const [field, folder] of BAUT1_DOKUMEN) {
        const files = (req.files && req.files[field]) || [];
        if (files.length === 0) continue; // tanpa file baru → dokumen lama dipertahankan
        const dir = path.join(__dirname, 'uploads', ntc, nlc, folder);
        fs.rmSync(dir, { recursive: true, force: true });
        pastikanFolderEksis(dir);
        let no = 0;
        for (const f of files) {
          if (f.mimetype === 'application/pdf') {
            const tmp = path.join(PDF_CACHE_DIR,
              'baut1_src_' + Date.now() + '_' + (++urut) + '.pdf');
            fs.writeFileSync(tmp, f.buffer);
            const hal = await renderPdfHalaman(tmp);
            try { fs.unlinkSync(tmp); } catch (e) {}
            for (const png of hal) {
              no++;
              fs.copyFileSync(png, path.join(dir,
                'page-' + String(no).padStart(3, '0') + '.png'));
            }
          } else {
            no++;
            await sharp(f.buffer)
              .flatten({ background: '#ffffff' })
              .resize({ width: 2480, height: 3508, fit: 'inside', withoutEnlargement: true })
              .jpeg({ quality: 92, chromaSubsampling: '4:4:4' })
              .toFile(path.join(dir, 'page-' + String(no).padStart(3, '0') + '.jpg'));
          }
        }
        hasil[field] = no;
        console.log('[BAUT1 upload] ' + field + ' → ' + no + ' halaman');
      }
      res.json({ ok: true, tersimpan: hasil });
    } catch (err) {
      console.error('Error upload BAUT-1:', err);
      res.status(500).json({ message: 'Gagal menyimpan dokumen BAUT-1.', detail: err.message });
    }
  });

// ─────────────────────────────────────────────────────────────────────────────
// [API] GENERATE BA PRELIM — POST dengan upload banyak Lampiran
// ─────────────────────────────────────────────────────────────────────────────
const prelimUploadFields = [
  { name: 'kesepakatan_file', maxCount: 1 },
  { name: 'penunjukan_file', maxCount: 1 },
  { name: 'boq_file', maxCount: 1 },
  { name: 'invoice_file', maxCount: 1 },
  { name: 'kwitansi_file', maxCount: 1 },
  { name: 'ba_penerimaan_file', maxCount: 1 },
  { name: 'ba_klarifikasi_file', maxCount: 1 },
  { name: 'mom_file', maxCount: 1 },
  { name: 'sph_file', maxCount: 1 },
  { name: 'spph_file', maxCount: 1 },
  { name: 'undangan_file', maxCount: 1 },
  { name: 'ktp_npwp_file', maxCount: 1 },
  { name: 'surat_pemkot_file', maxCount: 1 },
  { name: 'nota_kesepahaman_file', maxCount: 1 },
  { name: 'dokumentasi_file', maxCount: 1 }
];

app.post('/api/generate/ba_prelim/:lokasi_id', checkAuth('admin'), upload.fields(prelimUploadFields), async (req, res) => {
  try {
    const lokasi_id = req.params.lokasi_id;

    // Cek ketersediaan template BA Prelim
    const templatePath = path.join(__dirname, 'templates', 'template_BA_Prelim.docx');
    if (!fs.existsSync(templatePath)) {
      return res.status(404).json({ message: 'Template "template_BA_Prelim.docx" belum tersedia di server.' });
    }

    const lokasiData = await db.get('SELECT * FROM lokasi WHERE id = ?', [lokasi_id]);
    if (!lokasiData) return res.status(404).json({ message: 'Data lokasi tidak ditemukan.' });

    const ntc = (lokasiData.tematik || 'Umum').replace(/[^a-zA-Z0-9_-]/g, '_');
    const nlc = lokasiData.nama_lokasi.replace(/[^a-zA-Z0-9_-]/g, '_');
    const blankImage = path.join(__dirname, 'public', 'images', 'blank.png');

    // Folder tempat nyimpan file prelim
    const folderPrelim = path.join(__dirname, 'uploads', ntc, nlc, '-BA_PRELIM');
    pastikanFolderEksis(folderPrelim);

    const payload = {
      npprelim: req.body.npprelim || '-',
      nkprelim: req.body.nkprelim || '-',
      hariprelim: req.body.hariprelim || '-',
      tglprelim: req.body.tglprelim || '-',
      blmprelim: req.body.blmprelim || '-',
      lokasi1: req.body.lokasi1 || '-',
      sto1: req.body.sto1 || '-',
      nilaiprelim1: req.body.nilaiprelim1 || '-',
      totalprelim: req.body.totalprelim || '-',
      namadrprelim: req.body.namadrprelim || '-',
      noktpprelim: req.body.noktpprelim || '-',
      penyebutannilai: req.body.penyebutannilai || '-'
    };

    // Array nama slot sesuai urutan
    const fileSlots = ['kesepakatan', 'penunjukan', 'boq', 'invoice', 'kwitansi', 'ba_penerimaan', 'ba_klarifikasi', 'mom', 'sph', 'spph', 'undangan', 'ktp_npwp', 'surat_pemkot', 'nota_kesepahaman', 'dokumentasi'];

    // Proses penyimpanan foto/PDF yang sudah dikonversi
    for (const slot of fileSlots) {
      const fieldName = slot + '_file';
      if (req.files && req.files[fieldName] && req.files[fieldName][0]) {
        const fileBuf = req.files[fieldName][0].buffer;
        const destPath = path.join(folderPrelim, `PRELIM_${slot}_${nlc}.jpg`);
        // Standarkan ukuran A4 (bisa disesuaikan ukurannya bila tidak ingin seukuran A4 penuh)
        await sharp(fileBuf)
          .resize({ width: 2480, height: 3508, fit: 'inside', withoutEnlargement: true })
          .jpeg({ quality: 92, chromaSubsampling: '4:4:4' })
          .toFile(destPath);
        
        payload[slot] = destPath; // Misal {%kesepakatan}
      } else {
        payload[slot] = blankImage;
      }
    }

    // Set Image Module agar gambar ter-*inject* di DOCX
    const imageModulePrelim = new ImageModule({
      getImage: tv => fs.readFileSync((tv && fs.existsSync(tv)) ? tv : blankImage),
      getSize: (img, tv, tn) => {
        // Asumsi format di dalam template word memakai 1 kertas full -> 490x680
        return [490, 680]; 
      },
    });

    const doc = new Docxtemplater(
      new PizZip(fs.readFileSync(templatePath, 'binary')),
      {
        modules: [imageModulePrelim],
        paragraphLoop: true,
        linebreaks: true,
        parser: tag => ({ get: scope => {
          const ct = tag.replace(/[\u200B-\u200D\uFEFF]/g, '').trim();
          return scope[ct] !== undefined ? scope[ct] : '';
        }}),
      }
    );

    doc.render(payload);
    const buf0 = doc.getZip().generate({ type: 'nodebuffer', compression: 'DEFLATE' });
    const buf = await lampirkanPetaKeDocx(buf0, lokasi_id);

    const namaFile = 'BA_PRELIM_' + nlc + '.docx';
    await catatAktivitas('ba_prelim', lokasi_id);
    await kirimDokumen(req, res, buf, namaFile);

  } catch (err) {
    console.error('Error generate BA Prelim:', err);
    res.status(500).json({ message: 'Gagal generate BA Prelim.', detail: err.message });
  }
});
// ─────────────────────────────────────────────────────────────────────────────
// [API] HEAD + GET GENERATE UMUM (lact, ba_lapangan, dll)
// Dipakai frontend untuk cek template dulu (HEAD), lalu download (GET)
// ─────────────────────────────────────────────────────────────────────────────
async function handleGenerate(jenis, lokasi_id, req, res, db, path, fs, PizZip, Docxtemplater, ImageModule, sizeOf) {
  const templateNama = TEMPLATE_MAP[jenis];
  if (!templateNama) {
    return res.status(400).json({ message: 'Jenis dokumen tidak dikenal: ' + jenis });
  }

  // Pilihan template: FIXED_* (jika ada) → TEMPLATE_MAP → fallback templateRelokasi_clean.docx
  const fixedTemplate = path.join(__dirname, 'templates', `FIXED_${templateNama}`);
  const templatePath = fs.existsSync(fixedTemplate)
    ? fixedTemplate
    : path.join(__dirname, 'templates', templateNama);
  const finalTemplate = fs.existsSync(templatePath)
    ? templatePath
    : path.join(__dirname, 'templates', 'templateRelokasi_clean.docx');

  if (!fs.existsSync(finalTemplate)) {
    return res.status(404).json({
      message: 'Template belum tersedia. Harap siapkan file template terlebih dahulu.',
      template: templateNama,
      jenis,
    });
  }

const lokasiData  = await db.get('SELECT * FROM lokasi WHERE id = ?', [lokasi_id]);
    if (!lokasiData)  return res.status(404).json({ message: 'Data lokasi tidak ditemukan.' });

  const adminInput  = await db.get('SELECT * FROM laporan_rekon WHERE lokasi_id = ?', [lokasi_id]) || {};

  // ── Fallback: ambil no_sp dari tabel lokasi jika kontrak & surat_pesanan kosong ──
  // (sama seperti /api/download-docx — agar SP & Kontrak ikut terisi di LACT)
  if ((!adminInput.kontrak || adminInput.kontrak === '-') &&
      (!adminInput.surat_pesanan || adminInput.surat_pesanan === '-')) {
    const spRow = await db.get('SELECT no_sp FROM lokasi WHERE id = ?', [lokasi_id]).catch(() => null);
    if (spRow && spRow.no_sp) adminInput.surat_pesanan = spRow.no_sp;
  }

  const semuaBerkas = await db.all(
    'SELECT file_name, kategori, sub_kategori FROM berkas WHERE lokasi_id = ? ORDER BY id ASC',
    [lokasi_id]
  );

  const ntc = (lokasiData.tematik || 'Umum').replace(/[^a-zA-Z0-9_-]/g, '_');
  const nlc = lokasiData.nama_lokasi.replace(/[^a-zA-Z0-9_-]/g, '_');
  const blankImage = path.join(__dirname, 'public', 'images', 'blank.png');

  const KATEGORI_MAP = {
    "EPT":   "-EVIDEN PEMASANGAN TIANG",
    "EPEN":  "-EVIDEN PENARIKAN",
    "EPODP": "-EVIDEN PEMASANGAN ODP",
    "EOPM":  "-EVIDENT HASIL UKUR OPM",
    "LDP":   "-LAMPIRAN DATA PENGUKURAN OPM",
    "MCR":   "-MANCORE",
    "GR":    "-GROUNDING",
    "LKDL":  "-LAMPIRAN HASIL UJI KEDALAMAN GALIAN",
    "EGAL":  "-EVIDEN EVIDEN GALIAN DAN DIMENSI HANDHOLE",
    "EDIST": "-EVIDEN DISMANTLING",
    "ECO":   "-EVIDEN COLUSER",
    "LCO":   "-LAMPIRAN CO",
    "ROSET": "-Instalasi Roset",
    "EEP":   "MASUKAN EVIDEN PEKERJAAN",
  };

  const MAX_SLOT = {
    "EPT":9,"EPEN":9,"EPODP":27,"EOPM":45,"LDP":4,
    "MCR":2,"GR":18,"LKDL":3,"EGAL":18,"EDIST":9,
    "ECO":45,"LCO":9,"ROSET":9,"EEP":18
  };

  function ambilFoto(kode, slotKe) {
    const namaKat = KATEGORI_MAP[kode];
    if (!namaKat) return blankImage;
    // EOPM: urutkan foto mengikuti ODP → port agar slot EOPM1..N sesuai eviden
    const list = (kode === 'EOPM')
      ? urutBerkasEopm(semuaBerkas.filter(b => b.kategori === namaKat))
      : semuaBerkas.filter(b => b.kategori === namaKat);
    const target = list[slotKe - 1];
    if (target) {
      const fp = path.join(__dirname, 'uploads', ntc, nlc,
        target.kategori.replace(/[^a-zA-Z0-9_-]/g, '_'), target.file_name);
      return fs.existsSync(fp) ? fp : blankImage;
    }
    return blankImage;
  }

  const TAG_BESAR = ['LDP', 'MCR', 'LKDL'];
  const imageModule = new ImageModule({
    getImage: tv => fs.readFileSync((tv && fs.existsSync(tv)) ? tv : blankImage),
    getSize: (img, tv, tn) => {
      try {
        const dimensions = sizeOf(img);
        // Ukuran per-jenis tag (samakan dengan /api/download-docx):
        // sel tabel eep/redaman hanya ±2420 dxa (≈161px), foto lebih besar
        // dari itu akan memaksa tabel melebar → layout berantakan.
        const kode = (tn || '').trim().replace(/\d+$/, '');
        const isBesar    = TAG_BESAR.includes(kode);
        const isFotoLoop = kode === 'f';        // {%f1}..{%f4} — grid 4 kolom
        const isRedFoto  = kode === 'redFoto';  // {%redFoto1..3} — grid 3 kolom
        const isPortFoto = kode === 'foto_path';
        const isOpmGrid  = (kode === 'foto' && tn !== 'foto'); // {%foto1..16} — grid OPM
        const isStepFoto = ['foto_sbl', 'foto_prg', 'foto_sdh'].includes(kode); // ODP/closure
        const isEdist    = kode === 'EDIST';    // {%EDIST1..9} — grid dismantling 3 kolom
        // Ukuran = ukuran letterbox di payload → rasio 1, hasil seragam persis.
        const maxW = isBesar ? 350 : (isStepFoto ? 210 : (isEdist ? 200 : (isFotoLoop ? 160 : (isRedFoto ? 170 : (isPortFoto ? 170 : (isOpmGrid ? 175 : 150))))));
        const maxH = isBesar ? 280 : (isStepFoto ? 160 : (isEdist ? 150 : (isFotoLoop ? 120 : (isRedFoto ? 130 : (isPortFoto ? 130 : (isOpmGrid ? 135 : 110))))));
        const ratio = Math.min(maxW / dimensions.width, maxH / dimensions.height);
        return [Math.round(dimensions.width * ratio), Math.round(dimensions.height * ratio)];
      } catch (e) { return [160, 200]; }
    },
  });

  const payload = {
    proyek:       adminInput.proyek        || '-',
    kontrak:      adminInput.kontrak       || '-',
    surat_pesan:  adminInput.surat_pesanan || '-',
    district:     adminInput.district      || '-',
    lokasi:       lokasiData.nama_lokasi   || '-',
    tematik:      lokasiData.tematik       || '-',
    pid:          lokasiData.pid           || '-',
    lokasi_smile: lokasiData.nama_lokasi   || '-',
    pelaksana:    adminInput.pelaksana     || '-',
    nama_waspang: adminInput.nama_waspang  || '-',
    nik_waspang:  adminInput.nik_waspang   || '-',
    hari_ct:      adminInput.hari_ct       || '-',
    tgl:          adminInput.tgl           || '-',
    bulan:        adminInput.bulan         || '-',
    ct_hari:      adminInput.ct_hari       || '-',
    ct_tgl:       adminInput.ct_tgl        || '-',
    ct_bulan:     adminInput.ct_bulan      || '-',
    ibooster_hari:     adminInput.ibooster_hari     || '-',
    ibooster_tanggal:  adminInput.ibooster_tanggal  || '-',
    ibooster_bulan:    adminInput.ibooster_bulan    || '-',
    ibooster_lokasi:   adminInput.ibooster_lokasi   || '-',
    dismantling_hari:     adminInput.dismantling_hari     || '-',
    dismantling_tanggal:  adminInput.dismantling_tanggal  || '-',
    dismantling_bulan:    adminInput.dismantling_bulan    || '-',
    // Gudang STO (tag {STO}) + pilihan DISMANTLING / TIDAK DISMANTLING
    STO:            adminInput.gudang_sto || '-',
    gudang_sto:     adminInput.gudang_sto || '-',
    ...flagPilihanDismantling(adminInput.pilih_dismantling),
    harico:        adminInput.harico        || '-',
    tglco:         adminInput.tglco         || '-',
    bulanco:       adminInput.bulanco       || '-',
    sambungan1co:  adminInput.sambungan1co  || '',
    sambungan2co:  adminInput.sambungan2co  || '',
    sambungan3co:  adminInput.sambungan3co  || '',
    sambungan4co:  adminInput.sambungan4co  || '',
    matrial_aset1: adminInput.matrial_aset1 || '-', volume_aset1: adminInput.volume_aset1 || '-', kondisi_aset1: adminInput.kondisi_aset1 || '-',
    matrial_aset2: adminInput.matrial_aset2 || '-', volume_aset2: adminInput.volume_aset2 || '-', kondisi_aset2: adminInput.kondisi_aset2 || '-',
    matrial_aset3: adminInput.matrial_aset3 || '-', volume_aset3: adminInput.volume_aset3 || '-', kondisi_aset3: adminInput.kondisi_aset3 || '-',
    matrial_aset4: adminInput.matrial_aset4 || '-', volume_aset4: adminInput.volume_aset4 || '-', kondisi_aset4: adminInput.kondisi_aset4 || '-',
    matrial_aset5: adminInput.matrial_aset5 || '-', volume_aset5: adminInput.volume_aset5 || '-', kondisi_aset5: adminInput.kondisi_aset5 || '-',
    tgl_generate: new Date().toLocaleDateString('id-ID', {
      weekday:'long', day:'2-digit', month:'long', year:'numeric'
    }),
  };

  Object.entries(MAX_SLOT).forEach(([kode, maxSlot]) => {
    for (let s = 1; s <= maxSlot; s++) payload[`${kode}${s}`] = ambilFoto(kode, s);
  });
  // EDIST = grid 3 kolom (sel 3260 dxa ≈ 217px) → letterbox seragam 200×150
  for (let s = 1; s <= MAX_SLOT.EDIST; s++) {
    payload[`EDIST${s}`] = await seragamkanFoto(payload[`EDIST${s}`], 200, 150);
  }

  // Data LOOP hasil ukur OPM per ODP + per port (foto / ISI / kosong)
  const folderOpmGen = path.join(__dirname, 'uploads', ntc, nlc, '-EVIDENT_HASIL_UKUR_OPM');
  payload.opm_list = await siapkanDataOpm(lokasi_id, folderOpmGen, '-EVIDENT HASIL UKUR OPM', 'opm', {
    proyek:       payload.proyek        || '-',
    kontrak:      payload.kontrak       || '-',
    surat_pesan:  payload.surat_pesan   || '-',
    district:     payload.district      || '-',
    lokasi_smile: payload.lokasi_smile  || '-',
  });
  const folderEpodpGen = path.join(__dirname, 'uploads', ntc, nlc, '-EVIDEN_PEMASANGAN_ODP');
  payload.epodp_list = await susunDataEpodp(lokasi_id, folderEpodpGen, {
    proyek:       payload.proyek        || '-',
    kontrak:      payload.kontrak       || '-',
    surat_pesan:  payload.surat_pesan   || '-',
    district:     payload.district      || '-',
    lokasi_smile: payload.lokasi_smile  || '-',
  });
  const folderEcoGen = path.join(__dirname, 'uploads', ntc, nlc, '-EVIDEN_COLUSER');
  payload.eco_list = await susunDataEco(lokasi_id, folderEcoGen, {
    proyek:       payload.proyek        || '-',
    kontrak:      payload.kontrak       || '-',
    surat_pesan:  payload.surat_pesan   || '-',
    district:     payload.district      || '-',
    lokasi_smile: payload.lokasi_smile  || '-',
  });
  // MASUKAN EVIDEN PEKERJAAN — loop per sub_kategori
  const folderEepGen = path.join(__dirname, 'uploads', ntc, nlc, 'MASUKAN_EVIDEN_PEKERJAAN');
  payload.eep_list = await siapkanDataEep(
  lokasi_id,
  folderEepGen,
  blankImage,
  {
    proyek:       payload.proyek        || '-',
    kontrak:      payload.kontrak       || '-',   // ← bukan payload.contracts!
    surat_pesan:  payload.surat_pesan   || '-',   // ← bukan payload.nomorSuratPesanan!
    district:     payload.district      || '-',
    lokasi_smile: payload.lokasi_smile  || '-',
  }
);
  console.log('[HG-EEP] lokasi_id=' + lokasi_id + ' folder=' + folderEepGen + ' exists=' + fs.existsSync(folderEepGen));
  console.log('[HG-EEP] groups=' + (payload.eep_list || []).length);
  (payload.eep_list || []).forEach(function(g){ console.log('[HG-EEP]   ' + g.nomor + '. ' + g.nama + ' (' + g.jumlah + ' foto, ' + g.foto_baris.length + ' baris)'); });

  const doc = new Docxtemplater(
    new PizZip(fs.readFileSync(finalTemplate, 'binary')),
    {
      modules: [imageModule],
      paragraphLoop: true,
      linebreaks: true,
      parser: tag => ({ get: scope => {
        const ct = tag.replace(/[\u200B-\u200D\uFEFF]/g, '').trim();
        return scope[ct] !== undefined ? scope[ct] : '';
      }}),
    }
  );

  doc.render(payload);
  const buf0 = doc.getZip().generate({ type: 'nodebuffer', compression: 'DEFLATE' });
  const buf = await lampirkanPetaKeDocx(buf0, lokasi_id);

  const namaFile = `${jenis.toUpperCase()}_${nlc}.docx`;
  await kirimDokumen(req, res, buf, namaFile);
  console.log(`Generate ${jenis} OK: lokasi=${lokasi_id}`);
}

// HEAD — cek template tersedia
app.head('/api/generate/:jenis/:lokasi_id', checkAuth('admin'), (req, res) => {
  const { jenis } = req.params;
  if (jenis === 'baut') return res.sendStatus(200);
  const templateNama = TEMPLATE_MAP[jenis];
  if (!templateNama) return res.sendStatus(400);
  const fixedTp = path.join(__dirname, 'templates', `FIXED_${templateNama}`);
  const tp = fs.existsSync(fixedTp) ? fixedTp : path.join(__dirname, 'templates', templateNama);
  return res.sendStatus(fs.existsSync(tp) ? 200 : 404);
});

// GET — download dokumen LACT (bisa diakses user & admin — untuk BA CO)
app.get('/api/generate/lact/:lokasi_id', checkAuth(), async (req, res) => {
  try {
    await handleGenerate('lact', req.params.lokasi_id, req, res, db, path, fs, PizZip, Docxtemplater, ImageModule, sizeOf);
    if (res.statusCode < 400) await catatAktivitas('lact', req.params.lokasi_id);
  } catch (err) {
    console.error('Error generate lact (user):', err);
    if (!res.headersSent) res.status(500).json({ message: 'Gagal generate LACT.', detail: err.message });
  }
});

// GET — download dokumen (semua jenis kecuali BAUT & BA_TAMBAHKURANG yang sudah punya route sendiri)
app.get('/api/generate/:jenis/:lokasi_id', checkAuth('admin'), async (req, res) => {
  const { jenis, lokasi_id } = req.params;
  if (jenis === 'baut')            return res.status(405).json({ message: 'BAUT menggunakan metode POST.' });
  if (jenis === 'ba_tambahkurang') {
    const qs = req.url.includes('?') ? req.url.slice(req.url.indexOf('?')) : '';
    return res.redirect(`/api/generate/ba_tambahkurang/${lokasi_id}${qs}`);
  }
  try {
    await handleGenerate(jenis, lokasi_id, req, res, db, path, fs, PizZip, Docxtemplater, ImageModule, sizeOf);
    if (res.statusCode < 400) await catatAktivitas(jenis, lokasi_id);
  } catch (err) {
    console.error('Error generate', jenis, err);
    res.status(500).json({ message: 'Gagal generate dokumen.', detail: err.message });
  }
});
// ═══════════════════════════════════════════════════════════════
// [API] HEAD — Cek ketersediaan template (dipakai frontend sebelum download)
// ═══════════════════════════════════════════════════════════════
app.head('/api/generate/:jenis/:lokasi_id', checkAuth('admin'), (req, res) => {
  const { jenis } = req.params;
  // BAUT pakai POST bukan GET, tapi HEAD tetap 200 agar frontend bisa lanjut ke modal
  if (jenis === 'baut') return res.sendStatus(200);
  const templateNama = TEMPLATE_MAP[jenis];
  if (!templateNama) return res.sendStatus(400);
  const tp = path.join(__dirname, 'templates', templateNama);
  return res.sendStatus(fs.existsSync(tp) ? 200 : 404);
});

// ═══════════════════════════════════════════════════════════════
// [API] GET — Download dokumen (lact, ba_lapangan, dll — BUKAN baut & ba_tambahkurang)
// ═══════════════════════════════════════════════════════════════
app.get('/api/generate/:jenis/:lokasi_id', checkAuth('admin'), async (req, res) => {
  const { jenis, lokasi_id } = req.params;

  // BAUT hanya via POST
  if (jenis === 'baut') {
    return res.status(405).json({ message: 'BAUT menggunakan metode POST dengan upload file.' });
  }
  // BA Tambah Kurang sudah punya route sendiri di atas
  if (jenis === 'ba_tambahkurang') {
    const qs = req.url.includes('?') ? req.url.slice(req.url.indexOf('?')) : '';
    return res.redirect(`/api/generate/ba_tambahkurang/${lokasi_id}${qs}`);
  }

  const templateNama = TEMPLATE_MAP[jenis];
  if (!templateNama) {
    return res.status(400).json({ message: 'Jenis dokumen tidak dikenal: ' + jenis });
  }

  const templatePath = path.join(__dirname, 'templates', templateNama);
  if (!fs.existsSync(templatePath)) {
    return res.status(404).json({
      message: `Template "${templateNama}" belum tersedia di folder /templates/`,
      template: templateNama,
      jenis,
    });
  }

  try {
    const lokasiData  = await db.get('SELECT * FROM lokasi WHERE id = ?', [lokasi_id]);
    if (!lokasiData)  return res.status(404).json({ message: 'Data lokasi tidak ditemukan.' });

    const adminInput  = await db.get('SELECT * FROM laporan_rekon WHERE lokasi_id = ?', [lokasi_id]) || {};
    const semuaBerkas = await db.all(
      'SELECT file_name, kategori, sub_kategori FROM berkas WHERE lokasi_id = ? ORDER BY id ASC',
      [lokasi_id]
    );

    const ntc = (lokasiData.tematik || 'Umum').replace(/[^a-zA-Z0-9_-]/g, '_');
    const nlc = lokasiData.nama_lokasi.replace(/[^a-zA-Z0-9_-]/g, '_');
    const blankImage = path.join(__dirname, 'public', 'images', 'blank.png');

    const KATEGORI_MAP_GEN = {
      'EPT':   '-EVIDEN PEMASANGAN TIANG',
      'EPEN':  '-EVIDEN PENARIKAN',
      'EPODP': '-EVIDEN PEMASANGAN ODP',
      'EOPM':  '-EVIDENT HASIL UKUR OPM',
      'LDP':   '-LAMPIRAN DATA PENGUKURAN OPM',
      'MCR':   '-MANCORE',
      'GR':    '-GROUNDING',
      'LKDL':  '-LAMPIRAN HASIL UJI KEDALAMAN GALIAN',
      'EGAL':  '-EVIDEN EVIDEN GALIAN DAN DIMENSI HANDHOLE',
      'EDIST': '-EVIDEN DISMANTLING',
      'ECO':   '-EVIDEN COLUSER',
      'LCO':   '-LAMPIRAN CO',
      'ROSET': '-Instalasi Roset',
      'EEP':   'MASUKAN EVIDEN PEKERJAAN',
    };

    const MAX_SLOT_GEN = {
      'EPT':9,'EPEN':9,'EPODP':27,'EOPM':45,'LDP':4,
      'MCR':2,'GR':18,'LKDL':3,'EGAL':18,'EDIST':9,
      'ECO':45,'LCO':9,'ROSET':9,'EEP':18
    };

    function ambilFotoGen(kode, slotKe) {
      const namaKat = KATEGORI_MAP_GEN[kode];
      if (!namaKat) return blankImage;
      // EOPM: urutkan foto mengikuti ODP → port agar slot EOPM1..N sesuai eviden
      const list = (kode === 'EOPM')
        ? urutBerkasEopm(semuaBerkas.filter(b => b.kategori === namaKat))
        : semuaBerkas.filter(b => b.kategori === namaKat);
      const target = list[slotKe - 1];
      if (target) {
        const fp = path.join(__dirname, 'uploads', ntc, nlc,
          target.kategori.replace(/[^a-zA-Z0-9_-]/g, '_'), target.file_name);
        return fs.existsSync(fp) ? fp : blankImage;
      }
      return blankImage;
    }

    const TAG_BESAR_GEN = ['LDP', 'MCR', 'LKDL'];
    const imageModuleGen = new ImageModule({
      getImage: tv => fs.readFileSync((tv && fs.existsSync(tv)) ? tv : blankImage),
      getSize: (img, tv, tn) => {
        try {
          const dim = sizeOf(img);
          // Ukuran per-jenis tag (samakan dengan /api/download-docx) —
          // sel tabel eep/redaman hanya ±2420 dxa (≈161px).
          const kode = (tn || '').trim().replace(/\d+$/, '');
          const isBesar    = TAG_BESAR_GEN.includes(kode);
          const isFotoLoop = kode === 'f';
          const isRedFoto  = kode === 'redFoto';
          const isPortFoto = kode === 'foto_path';
          const isOpmGrid  = (kode === 'foto' && tn !== 'foto');
          const isStepFoto = ['foto_sbl', 'foto_prg', 'foto_sdh'].includes(kode); // ODP/closure
          const isEdist    = kode === 'EDIST';    // {%EDIST1..9} — grid dismantling
          // Ukuran = ukuran letterbox di payload → rasio 1, hasil seragam persis.
          const maxW = isBesar ? 350 : (isStepFoto ? 210 : (isEdist ? 200 : (isFotoLoop ? 160 : (isRedFoto ? 170 : (isPortFoto ? 170 : (isOpmGrid ? 175 : 150))))));
          const maxH = isBesar ? 280 : (isStepFoto ? 160 : (isEdist ? 150 : (isFotoLoop ? 120 : (isRedFoto ? 130 : (isPortFoto ? 130 : (isOpmGrid ? 135 : 110))))));
          const ratio = Math.min(maxW / dim.width, maxH / dim.height);
          return [Math.round(dim.width * ratio), Math.round(dim.height * ratio)];
        } catch (e) { return [160, 200]; }
      },
    });

    const payload = {
      proyek:       adminInput.proyek        || '-',
      kontrak:      adminInput.kontrak       || '-',
      surat_pesan:  adminInput.surat_pesanan || '-',
      district:     adminInput.district      || '-',
      lokasi_smile: lokasiData.nama_lokasi   || '-',
      pelaksana:    adminInput.pelaksana     || '-',
      nama_waspang: adminInput.nama_waspang  || '-',
      nik_waspang:  adminInput.nik_waspang   || '-',
      hari_ct:      adminInput.hari_ct       || '-',
      tgl:          adminInput.tgl           || '-',
      bulan:        adminInput.bulan         || '-',
      ct_hari:      adminInput.ct_hari       || '-',
      ct_tgl:       adminInput.ct_tgl        || '-',
      ct_bulan:     adminInput.ct_bulan      || '-',
      ibooster_hari:    adminInput.ibooster_hari    || '-',
      ibooster_tanggal: adminInput.ibooster_tanggal || '-',
      ibooster_bulan:   adminInput.ibooster_bulan   || '-',
      ibooster_lokasi:  adminInput.ibooster_lokasi  || '-',
      dismantling_hari:     adminInput.dismantling_hari     || '-',
      dismantling_tanggal:  adminInput.dismantling_tanggal  || '-',
      dismantling_bulan:    adminInput.dismantling_bulan    || '-',
      // Gudang STO (tag {STO}) + pilihan DISMANTLING / TIDAK DISMANTLING
      STO:            adminInput.gudang_sto || '-',
      gudang_sto:     adminInput.gudang_sto || '-',
      ...flagPilihanDismantling(adminInput.pilih_dismantling),
      harico:        adminInput.harico        || '-',
      tglco:         adminInput.tglco         || '-',
      bulanco:       adminInput.bulanco       || '-',
      sambungan1co:  adminInput.sambungan1co  || '',
      sambungan2co:  adminInput.sambungan2co  || '',
      sambungan3co:  adminInput.sambungan3co  || '',
      sambungan4co:  adminInput.sambungan4co  || '',
      matrial_aset1: adminInput.matrial_aset1 || '-', volume_aset1: adminInput.volume_aset1 || '-', kondisi_aset1: adminInput.kondisi_aset1 || '-',
      matrial_aset2: adminInput.matrial_aset2 || '-', volume_aset2: adminInput.volume_aset2 || '-', kondisi_aset2: adminInput.kondisi_aset2 || '-',
      matrial_aset3: adminInput.matrial_aset3 || '-', volume_aset3: adminInput.volume_aset3 || '-', kondisi_aset3: adminInput.kondisi_aset3 || '-',
      matrial_aset4: adminInput.matrial_aset4 || '-', volume_aset4: adminInput.volume_aset4 || '-', kondisi_aset4: adminInput.kondisi_aset4 || '-',
      matrial_aset5: adminInput.matrial_aset5 || '-', volume_aset5: adminInput.volume_aset5 || '-', kondisi_aset5: adminInput.kondisi_aset5 || '-',
      tgl_generate: new Date().toLocaleDateString('id-ID', {
        weekday: 'long', day: '2-digit', month: 'long', year: 'numeric'
      }),
    };

    // Isi semua slot foto
    Object.entries(MAX_SLOT_GEN).forEach(([kode, maxSlot]) => {
      for (let s = 1; s <= maxSlot; s++) {
        payload[`${kode}${s}`] = ambilFotoGen(kode, s);
      }
    });
    // EDIST = grid 3 kolom (sel 3260 dxa ≈ 217px) → letterbox seragam 200×150
    for (let s = 1; s <= MAX_SLOT_GEN.EDIST; s++) {
      payload[`EDIST${s}`] = await seragamkanFoto(payload[`EDIST${s}`], 200, 150);
    }

    // Data LOOP hasil ukur OPM per ODP + per port (foto / ISI / kosong)
    const folderOpmGen2 = path.join(__dirname, 'uploads', ntc, nlc, '-EVIDENT_HASIL_UKUR_OPM');
    payload.opm_list = await siapkanDataOpm(lokasi_id, folderOpmGen2, '-EVIDENT HASIL UKUR OPM', 'opm', {
      proyek:       payload.proyek        || '-',
      kontrak:      payload.kontrak       || '-',
      surat_pesan:  payload.surat_pesan   || '-',
      district:     payload.district      || '-',
      lokasi_smile: payload.lokasi_smile  || '-',
    });
    const folderEpodpGen2 = path.join(__dirname, 'uploads', ntc, nlc, '-EVIDEN_PEMASANGAN_ODP');
    payload.epodp_list = await susunDataEpodp(lokasi_id, folderEpodpGen2, {
      proyek:       payload.proyek        || '-',
      kontrak:      payload.kontrak       || '-',
      surat_pesan:  payload.surat_pesan   || '-',
      district:     payload.district      || '-',
      lokasi_smile: payload.lokasi_smile  || '-',
    });
    const folderEcoGen2 = path.join(__dirname, 'uploads', ntc, nlc, '-EVIDEN_COLUSER');
    payload.eco_list = await susunDataEco(lokasi_id, folderEcoGen2, {
      proyek:       payload.proyek        || '-',
      kontrak:      payload.kontrak       || '-',
      surat_pesan:  payload.surat_pesan   || '-',
      district:     payload.district      || '-',
      lokasi_smile: payload.lokasi_smile  || '-',
    });
    // MASUKAN EVIDEN PEKERJAAN — loop per sub_kategori
    const folderEepGen2 = path.join(__dirname, 'uploads', ntc, nlc, 'MASUKAN_EVIDEN_PEKERJAAN');
    payload.eep_list = await siapkanDataEep(
  lokasi_id,
  folderEepGen2,
  blankImage,
  {
    proyek:       payload.proyek        || '-',
    kontrak:      payload.kontrak       || '-',   // ← bukan payload.contracts!
    surat_pesan:  payload.surat_pesan   || '-',   // ← bukan payload.nomorSuratPesanan!
    district:     payload.district      || '-',
    lokasi_smile: payload.lokasi_smile  || '-',
  }
);

    const doc = new Docxtemplater(
      new PizZip(fs.readFileSync(templatePath, 'binary')),
      {
        modules: [imageModuleGen],
        paragraphLoop: true,
        linebreaks: true,
        parser: tag => ({ get: scope => {
          const ct = tag.replace(/[\u200B-\u200D\uFEFF]/g, '').trim();
          return scope[ct] !== undefined ? scope[ct] : '';
        }}),
      }
    );

    doc.render(payload);
    const buf0 = doc.getZip().generate({ type: 'nodebuffer', compression: 'DEFLATE' });
    const buf = await lampirkanPetaKeDocx(buf0, lokasi_id);

    const namaFile = `${jenis.toUpperCase()}_${nlc}.docx`;
    await kirimDokumen(req, res, buf, namaFile);
    console.log(`Generate ${jenis} OK: lokasi=${lokasi_id}, template=${templateNama}`);

  } catch (err) {
    console.error(`Error generate ${jenis}:`, err);
    res.status(500).json({ message: 'Gagal generate dokumen.', detail: err.message });
  }
});
// ─────────────────────────────────────────────────────────────────────────────
// [HALAMAN] Routes
// ─────────────────────────────────────────────────────────────────────────────
// Halaman RAB — hanya admin
app.get('/rab', checkAuth('admin'), (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'rab.html'));
});

// ═══ DEBUG: test EEP rendering ═══
app.get('/api/debug-eep/:lokasi_id', checkAuth('admin'), async (req, res) => {
  try {
    const lokasiId = req.params.lokasi_id;
    const lokasiData = await db.get('SELECT * FROM lokasi WHERE id = ?', [lokasiId]);
    if (!lokasiData) return res.status(404).json({ message: 'Lokasi tidak ditemukan' });

    const namaTematikClean = (lokasiData.tematik || 'Umum').replace(/[^a-zA-Z0-9_-]/g, '_');
    const namaLokasiClean = lokasiData.nama_lokasi.replace(/[^a-zA-Z0-9_-]/g, '_');
    const blankImage = path.join(__dirname, 'public', 'images', 'blank.png');

    const folderEep = path.join(__dirname, 'uploads', namaTematikClean, namaLokasiClean, 'MASUKAN_EVIDEN_PEKERJAAN');
    const eepList = await siapkanDataEep(lokasiId, folderEep, blankImage);

    // Check template
    const baseName = namaTemplateDocx(lokasiData.tematik);
    const fixedTemplate = path.join(__dirname, 'templates', `FIXED_${baseName}`);
    const templatePath = path.join(__dirname, 'templates', baseName);
    const finalTemplate = fs.existsSync(fixedTemplate)
      ? fixedTemplate
      : (fs.existsSync(templatePath) ? templatePath : path.join(__dirname, 'templates', 'templateRelokasi_clean.docx'));

    const tzip = new PizZip(fs.readFileSync(finalTemplate, 'binary'));
    const txml = tzip.files['word/document.xml'].asText();
    const hasEepOpen = txml.includes('{#eep_list}');
    const hasEepClose = txml.includes('{/eep_list}');
    const hasFoto1 = txml.includes('{%f1}');
    const hasFotoBaris = txml.includes('{#foto_baris}');

    res.json({
      lokasi: lokasiData.nama_lokasi,
      tematik: lokasiData.tematik,
      template: path.basename(finalTemplate),
      template_has_eep_open: hasEepOpen,
      template_has_eep_close: hasEepClose,
      template_has_foto1: hasFoto1,
      template_has_foto_baris: hasFotoBaris,
      folderEep: folderEep,
      folderExists: fs.existsSync(folderEep),
      eep_groups: eepList.length,
      eep_details: eepList.map(g => ({
        nomor: g.nomor, nama: g.nama, jumlah: g.jumlah,
        baris: g.foto_baris.length,
        sample_f1_exists: g.foto_baris[0] ? fs.existsSync(g.foto_baris[0].f1) : false
      }))
    });
  } catch (e) {
    res.status(500).json({ error: e.message, stack: e.stack });
  }
});

const PORT = process.env.PORT || 3000;
app.get('/', (req, res) => {
    res.redirect('/login.html');
});

// Endpoint debug: cek path foto yang akan dipakai generate
// GET /api/debug-generate/:lokasi_id
app.get('/api/debug-generate/:lokasi_id', checkAuth('admin'), async (req, res) => {
  try {
    const lokasiData = await db.get('SELECT * FROM lokasi WHERE id = ?', [req.params.lokasi_id]);
    if (!lokasiData) return res.status(404).json({ message: 'Lokasi tidak ditemukan' });

    const ntc = (lokasiData.tematik || 'Umum').replace(/[^a-zA-Z0-9_-]/g, '_');
    const nlc = lokasiData.nama_lokasi.replace(/[^a-zA-Z0-9_-]/g, '_');
    const baseDir = path.join(__dirname, 'uploads', ntc, nlc);

    const KM = {
      "EPT":"-EVIDEN PEMASANGAN TIANG","EPEN":"-EVIDEN PENARIKAN",
      "EPODP":"-EVIDEN PEMASANGAN ODP","EOPM":"-EVIDENT HASIL UKUR OPM",
      "LDP":"-LAMPIRAN DATA PENGUKURAN OPM","MCR":"-MANCORE",
      "GR":"-GROUNDING","LKDL":"-LAMPIRAN HASIL UJI KEDALAMAN GALIAN",
      "EGAL":"-EVIDEN EVIDEN GALIAN DAN DIMENSI HANDHOLE",
      "EDIST":"-EVIDEN DISMANTLING","ECO":"-EVIDEN COLUSER",
      "LCO":"-LAMPIRAN CO","ROSET":"-Instalasi Roset",
      "EEP":"MASUKAN EVIDEN PEKERJAAN"
    };

    const hasil = {};
    for (const [kode, kat] of Object.entries(KM)) {
      const katClean = kat.replace(/[^a-zA-Z0-9_-]/g, '_');
      const folderPath = path.join(baseDir, katClean);
      const exists = fs.existsSync(folderPath);
      const files = exists ? fs.readdirSync(folderPath).filter(f => /\.(jpg|jpeg|png|gif|webp)$/i.test(f)).sort() : [];
      hasil[kode] = { folder: folderPath, exists, jumlah: files.length, files };
    }

    // Cek template yang tersedia
    const templates = {};
    for (const [jenis, nama] of Object.entries(TEMPLATE_MAP)) {
      const tp = path.join(__dirname, 'templates', nama);
      templates[jenis] = { file: nama, exists: fs.existsSync(tp) };
    }

    res.json({
      lokasi: lokasiData.nama_lokasi,
      tematik: lokasiData.tematik,
      base_dir: baseDir,
      base_dir_exists: fs.existsSync(baseDir),
      foto_per_kategori: hasil,
      templates_tersedia: templates,
    });
  } catch (e) {
    res.status(500).json({ message: 'Error debug', detail: e.message });
  }
});


// ── PROJECT PLAN — 2 modal: input rencana proyek & upload eviden lapangan ────
// Tidak butuh login: siapa saja boleh mengisi form (lihat brief fitur ini).

const PP_DIR = path.join(UPLOADS_DIR, 'project-plan');
const PP_DESIGNATOR = [
  'AC-OF-SM-ADSS-12D', 'AC-OF-SM-ADSS-24D', 'AC-OF-SM-ADSS-48D',
  'AC-OF-SM-ADSS-96D', 'ADAPT-SPRING', 'ODC-ODP-ADAPT-STDR',
  'ODP-SOLID-PB-16', 'ODP-SOLID-PB-8', 'ODP-SOLID-PL-16', 'ODP-SOLID-PL-8',
  'PS-1-16-ODP', 'PS-1-16-ODP-S', 'ODP-SPL', 'ODP-LABEL', 'ODP-VA-PORT',
  'PS-1-2-ODC', 'PS-1-32-ODX', 'PS-1-4-ODC', 'PS-1-8-ODX', 'PS-1-8-ODX-S',
  'PS-2-2-ODC', 'PS-2-4-ODC', 'PS-2-8-ODX',
  'GT-J-HH1-TTP3', 'GT-J-HH2-TTP2', 'HH-PIT-80', 'HH-PIT-DP',
  'HH-PIT-HA', 'HH-PIT-P-HA', 'HH-PIT-P-ODC', 'HH-PIT-P-ODP',
  'MH-HH1', 'MH-HH2', 'Slack Support HH'
];
const PP_JENIS = ['Jasa', 'Material'];
const PP_KATEGORI_EVIDEN = ['Foto Progres', 'Foto Material', 'Foto Instalasi', 'Dokumen PDF', 'Lainnya'];
const PP_MAX_FILES = 50;
const PP_EXT_EVIDEN = ['.jpg', '.jpeg', '.png', '.webp', '.pdf'];
const PP_MIME_EVIDEN = ['image/jpeg', 'image/png', 'image/webp', 'application/pdf'];

// Simpan eviden ke uploads/project-plan/ dengan nama Date.now()_<namaasli>.
// Bila nama sudah terpakai (file serupa dalam satu request) ditambahi _1, _2, ...
const ppUpload = multer({
  storage: multer.diskStorage({
    destination: (req, file, cb) => {
      pastikanFolderEksis(PP_DIR);
      cb(null, PP_DIR);
    },
    filename: (req, file, cb) => {
      const asli = path.basename(String(file.originalname || '')).replace(/[^\w.\-]+/g, '_') || 'eviden';
      const stamp = Date.now();
      let nama = `${stamp}_${asli}`;
      let i = 1;
      while (fs.existsSync(path.join(PP_DIR, nama))) nama = `${stamp}_${i++}_${asli}`;
      cb(null, nama);
    }
  }),
  limits: { fileSize: 15 * 1024 * 1024, files: PP_MAX_FILES }
});

// Validasi payload items dipakai bersama oleh POST maupun PUT.
function ppCekItems(items) {
  if (!Array.isArray(items) || items.length === 0) return 'Minimal 1 baris designator wajib diisi.';
  for (const raw of items) {
    const designator = String(raw && raw.designator || '').trim();
    const jenis = String(raw && raw.jenis || '').trim();
    const satuan = String(raw && raw.satuan || '').trim();
    const volume = Number(raw && raw.volume);
    if (!PP_DESIGNATOR.includes(designator)) return 'Designator "' + designator + '" tidak tersedia dalam daftar.';
    if (!PP_JENIS.includes(jenis)) return 'Jenis harus "Jasa" atau "Material".';
    if (!satuan) return 'Satuan wajib diisi pada setiap baris designator.';
    if (!Number.isFinite(volume) || volume < 0) return 'Volume minimal 0.';
  }
  return null;
}

// [API] SIMPAN PROJECT PLAN — body JSON: { nama_lokasi, items: [...] }
app.post('/api/project-plan', async (req, res) => {
  let projectId = null;
  try {
    const namaLokasi = String(req.body && req.body.nama_lokasi || '').trim();
    if (!namaLokasi) return res.status(400).json({ success: false, message: 'Nama Lokasi wajib diisi.' });

    const kitems = ppCekItems(req.body && req.body.items);
    if (kitems) return res.status(400).json({ success: false, message: kitems });
    const items = req.body.items;

    // RETURNING id dipakai karena saveDb() di db-wrapper mengekspor database
    // setiap INSERT, yang mereset last_insert_rowid() ke 0.
    const barisId = await db.get('INSERT INTO project_plan (nama_lokasi) VALUES (?) RETURNING id', [namaLokasi]);
    projectId = barisId && barisId.id;
    if (!projectId) throw new Error('ID project plan tidak didapat.');

    // Tiap baris mengembalikan id-nya supaya eviden bisa menempel per designator.
    const itemsTersimpan = [];
    for (const raw of items) {
      const designator = String(raw.designator).trim();
      const jenis = String(raw.jenis).trim();
      const volume = Number(raw.volume);
      const satuan = String(raw.satuan).trim();
      const baris = await db.get(
        'INSERT INTO project_plan_item (project_id, designator, jenis, volume, satuan) VALUES (?, ?, ?, ?, ?) RETURNING id',
        [projectId, designator, jenis, volume, satuan]
      );
      itemsTersimpan.push({ id: baris && baris.id, designator, jenis, volume, satuan });
    }

    res.json({ success: true, project_id: projectId, items: itemsTersimpan });
  } catch (e) {
    console.error('[project-plan] Gagal menyimpan:', e);
    // Batalkan data setengah jadi supaya tidak menyisakan baris yatim.
    if (projectId) {
      await db.run('DELETE FROM project_plan_item WHERE project_id = ?', [projectId]).catch(() => {});
      await db.run('DELETE FROM project_plan WHERE id = ?', [projectId]).catch(() => {});
    }
    res.status(500).json({ success: false, message: 'Gagal menyimpan project plan. ' + (e && e.message || '') });
  }
});

// [API] UPDATE PROJECT PLAN — dipakai ketika eviden sudah tersimpan lebih dulu,
// supaya tombol "Simpan Eviden" tidak membuat data ganda.
app.put('/api/project-plan/:id', async (req, res) => {
  try {
    const projectId = Number(req.params.id);
    if (!Number.isInteger(projectId) || projectId <= 0) {
      return res.status(400).json({ success: false, message: 'ID project plan tidak valid.' });
    }
    const ada = await db.get('SELECT id FROM project_plan WHERE id = ?', [projectId]);
    if (!ada) return res.status(404).json({ success: false, message: 'Project plan tidak ditemukan.' });

    const namaLokasi = String(req.body && req.body.nama_lokasi || '').trim();
    if (!namaLokasi) return res.status(400).json({ success: false, message: 'Nama Lokasi wajib diisi.' });

    const kitems = ppCekItems(req.body && req.body.items);
    if (kitems) return res.status(400).json({ success: false, message: kitems });
    const items = req.body.items;

    await db.run('UPDATE project_plan SET nama_lokasi = ? WHERE id = ?', [namaLokasi, projectId]);

    const lama = await db.all('SELECT id FROM project_plan_item WHERE project_id = ? ORDER BY id', [projectId]);
    const punya = new Set(lama.map(r => r.id));
    const dipakai = new Set();
    const hasil = [];

    for (const raw of items) {
      const designator = String(raw.designator).trim();
      const jenis = String(raw.jenis).trim();
      const volume = Number(raw.volume);
      const satuan = String(raw.satuan).trim();
      const idLama = Number(raw.id) || 0;

      if (idLama && punya.has(idLama)) {
        await db.run(
          'UPDATE project_plan_item SET designator = ?, jenis = ?, volume = ?, satuan = ? WHERE id = ?',
          [designator, jenis, volume, satuan, idLama]
        );
        dipakai.add(idLama);
        hasil.push({ id: idLama, designator, jenis, volume, satuan });
      } else {
        const baris = await db.get(
          'INSERT INTO project_plan_item (project_id, designator, jenis, volume, satuan) VALUES (?, ?, ?, ?, ?) RETURNING id',
          [projectId, designator, jenis, volume, satuan]
        );
        hasil.push({ id: baris && baris.id, designator, jenis, volume, satuan });
      }
    }

    // Baris yang dihapus pengguna: eviden & berkas fisiknya ikut dibersihkan.
    const buang = [...punya].filter(x => !dipakai.has(x));
    if (buang.length) {
      const ph = buang.map(() => '?').join(',');
      const ev = await db.all('SELECT file_path FROM project_plan_evidensi WHERE item_id IN (' + ph + ')', buang);
      for (const e of ev) {
        try {
          const fp = path.join(PP_DIR, path.basename(String(e.file_path || '')));
          if (fp.startsWith(PP_DIR + path.sep) && fs.existsSync(fp)) fs.unlinkSync(fp);
        } catch (err) { /* abaikan */ }
      }
      await db.run('DELETE FROM project_plan_evidensi WHERE item_id IN (' + ph + ')', buang).catch(() => {});
      await db.run('DELETE FROM project_plan_item WHERE id IN (' + ph + ')', buang);
    }

    res.json({ success: true, project_id: projectId, items: hasil });
  } catch (e) {
    console.error('[project-plan] Gagal memperbarui:', e);
    res.status(500).json({ success: false, message: 'Gagal memperbarui project plan. ' + (e && e.message || '') });
  }
});

// [API] UPLOAD EVIDEN LAPANGAN — multipart: project_id, item_id?, kategori, files[], keterangan[]
app.post('/api/project-plan/eviden', ppUpload.array('files', PP_MAX_FILES), async (req, res) => {
  const fileTertulis = (req.files || []).map(f => path.join(PP_DIR, f.filename));
  const bersihkanFile = () => fileTertulis.forEach(fp => {
    try { if (fs.existsSync(fp)) fs.unlinkSync(fp); } catch (e) { /* abaikan */ }
  });
  const gagal = (pesan, status = 400) => {
    bersihkanFile();
    return res.status(status).json({ success: false, message: pesan });
  };
  const barisTersimpan = [];

  try {
    const projectId = Number(req.body.project_id);
    if (!Number.isInteger(projectId) || projectId <= 0) return gagal('Lokasi project wajib dipilih.');
    const lokasiAda = await db.get('SELECT id FROM project_plan WHERE id = ?', [projectId]);
    if (!lokasiAda) return gagal('Lokasi project tidak ditemukan.');

    const kategori = String(req.body.kategori || '').trim();
    if (!PP_KATEGORI_EVIDEN.includes(kategori)) return gagal('Kategori berkas tidak valid.');

    // Opsional: eviden diikat ke 1 baris designator (item_id).
    let rawItemId = req.body.item_id;
    if (Array.isArray(rawItemId)) rawItemId = rawItemId[0];
    let itemId = null;
    if (rawItemId !== undefined && rawItemId !== null && String(rawItemId).trim() !== '') {
      itemId = Number(rawItemId);
      if (!Number.isInteger(itemId) || itemId <= 0) return gagal('Baris designator eviden tidak valid.');
      const barisAda = await db.get('SELECT id FROM project_plan_item WHERE id = ? AND project_id = ?', [itemId, projectId]);
      if (!barisAda) return gagal('Baris designator tidak ditemukan pada project ini.');
    }

    const files = req.files || [];
    if (!files.length) return gagal('Pilih minimal 1 file eviden.');
    for (const f of files) {
      const ext = path.extname(String(f.originalname || '')).toLowerCase();
      if (!PP_EXT_EVIDEN.includes(ext)) return gagal('Format file hanya JPG, PNG, WEBP, atau PDF.');
      if (!PP_MIME_EVIDEN.includes(String(f.mimetype || ''))) return gagal('Format file tidak sesuai dengan jenis berkasnya.');
    }

    // keterangan per file: array sejajar 'files', atau satu teks untuk semua file
    let keterangan = req.body.keterangan;
    if (keterangan === undefined || keterangan === null) keterangan = [];
    if (!Array.isArray(keterangan)) keterangan = [keterangan];

    for (let i = 0; i < files.length; i++) {
      const ket = String(keterangan[i] !== undefined ? keterangan[i] : (keterangan[0] || '')).trim();
      const filePath = '/uploads/project-plan/' + files[i].filename;
      await db.run(
        'INSERT INTO project_plan_evidensi (project_id, item_id, kategori, keterangan, file_path) VALUES (?, ?, ?, ?, ?)',
        [projectId, itemId, kategori, ket || null, filePath]
      );
      barisTersimpan.push(filePath);
    }

    res.json({ success: true, count: barisTersimpan.length });
  } catch (e) {
    console.error('[project-plan] Gagal simpan eviden:', e);
    if (barisTersimpan.length) {
      const tanya = barisTersimpan.map(() => '?').join(',');
      await db.run('DELETE FROM project_plan_evidensi WHERE file_path IN (' + tanya + ')', barisTersimpan).catch(() => {});
    }
    bersihkanFile();
    res.status(500).json({ success: false, message: 'Gagal menyimpan eviden. ' + (e && e.message || '') });
  }
});

// [API] DAFTAR LOKASI PROJECT PLAN — untuk dropdown modal eviden
app.get('/api/project-plan/lokasi', async (req, res) => {
  try {
    const data = await db.all('SELECT id, nama_lokasi FROM project_plan ORDER BY created_at DESC, id DESC');
    res.json({ success: true, data });
  } catch (e) {
    console.error('[project-plan] Gagal membaca lokasi:', e);
    res.status(500).json({ success: false, message: 'Gagal memuat daftar lokasi.' });
  }
});

// [API] DAFTAR PROJECT PLAN (terbaru di atas) beserta item & evidensinya
// Khusus dashboard admin — input tetap terbuka tanpa login (lihat brief fitur).
app.get('/api/project-plan', checkAuth('admin'), async (req, res) => {
  try {
    const projects = await db.all('SELECT * FROM project_plan ORDER BY created_at DESC, id DESC');
    const data = [];
    for (const p of projects) {
      const items = await db.all('SELECT * FROM project_plan_item WHERE project_id = ? ORDER BY id ASC', [p.id]);
      const evidensi = await db.all('SELECT * FROM project_plan_evidensi WHERE project_id = ? ORDER BY id ASC', [p.id]);
      data.push({ ...p, items, evidensi });
    }
    res.json({ success: true, data });
  } catch (e) {
    console.error('[project-plan] Gagal membaca:', e);
    res.status(500).json({ success: false, message: 'Gagal memuat data project plan.' });
  }
});

// [API] HAPUS PROJECT PLAN — baris designator, evidensi, dan file fisik ikut terhapus
app.delete('/api/project-plan/:id', checkAuth('admin'), async (req, res) => {
  try {
    const id = Number(req.params.id);
    if (!Number.isInteger(id) || id <= 0) {
      return res.status(400).json({ success: false, message: 'ID project plan tidak valid.' });
    }

    const proyek = await db.get('SELECT id, nama_lokasi FROM project_plan WHERE id = ?', [id]);
    if (!proyek) return res.status(404).json({ success: false, message: 'Project plan tidak ditemukan.' });

    const daftarFile = await db.all('SELECT file_path FROM project_plan_evidensi WHERE project_id = ?', [id]);

    await db.run('DELETE FROM project_plan_evidensi WHERE project_id = ?', [id]);
    await db.run('DELETE FROM project_plan_item WHERE project_id = ?', [id]);
    await db.run('DELETE FROM project_plan WHERE id = ?', [id]);

    // File eviden tersimpan rata di uploads/project-plan/<nama file>
    let fileGagal = 0;
    for (const row of daftarFile) {
      const namaFile = path.basename(String(row.file_path || ''));
      if (!namaFile) continue;
      const fp = path.join(PP_DIR, namaFile);
      if (!fp.startsWith(PP_DIR + path.sep)) continue;
      try { if (fs.existsSync(fp)) fs.unlinkSync(fp); } catch (e) { fileGagal++; }
    }

    res.json({
      success: true,
      message: 'Project plan "' + proyek.nama_lokasi + '" beserta ' + daftarFile.length + ' berkas eviden dihapus.',
      file_gagal: fileGagal
    });
  } catch (e) {
    console.error('[project-plan] Gagal menghapus:', e);
    res.status(500).json({ success: false, message: 'Gagal menghapus project plan. ' + (e && e.message || '') });
  }
});


// ─────────────────────────────────────────────────────────────────────────────
// [API] REKAP OPM RELOKASI  (kategori "INPUT REKAP OPM")
//   1. GET  /api/rekap-opm/:lokasi_id          → daftar ODP+port (dari evidenced OPM)
//                                                 + nilai tersimpan + header rekap
//   2. POST /api/rekap-opm/:lokasi_id/scan      → scan angka tiap foto OPM (AI vision)
//   3. POST /api/rekap-opm/:lokasi_id/simpan     → simpan nilai yang sudah dikoreksi
//   4. POST /api/rekap-opm/:lokasi_id/generate   → isi Excel (rumus tetap) + render
//                                                 gambar per ODP → simpan sebagai berkas
// ─────────────────────────────────────────────────────────────────────────────
const KATEGORI_REKAP_OPM = 'INPUT REKAP OPM';

async function daftarOdpOpm(lokasi_id, kategori = '-EVIDENT HASIL UKUR OPM') {
  const odps = await db.all('SELECT odp, kapasitas FROM opm_odp WHERE lokasi_id = ? ORDER BY odp ASC', [lokasi_id]);
  const fotoRows = await db.all('SELECT sub_kategori, file_name FROM berkas WHERE lokasi_id = ? AND kategori = ? ORDER BY id ASC', [lokasi_id, kategori]);
  const odpMap = new Map();
  odps.forEach(g => odpMap.set(g.odp, g.kapasitas || 8));
  fotoRows.forEach(f => {
    const nm = (f.sub_kategori || '').trim();
    if (nm && !odpMap.has(nm)) odpMap.set(nm, 8);
  });
  return { odpMap, fotoRows };
}

function filePortOpm(fotoRows, nama) {
  // Kembalikan map port -> nama file (dari pola "-P<n>." atau "<n>." di akhir nama)
  const map = {};
  fotoRows.forEach(f => {
    if ((f.sub_kategori || '').trim() !== nama) return;
    const m = (f.file_name || '').match(/-P(\d+)\.(jpg|jpeg|png|webp|pdf)$/i) ||
              (f.file_name || '').match(/(\d+)\.(jpg|jpeg|png|webp|pdf)$/i);
    if (m) map[parseInt(m[1])] = f.file_name;
  });
  return map;
}

app.get('/api/rekap-opm/:lokasi_id', async (req, res) => {
  try {
    const lokasi_id = req.params.lokasi_id;
    const kategori = req.query.kategori || '-EVIDENT HASIL UKUR OPM';

    const lok = await db.get('SELECT tematik, nama_lokasi FROM lokasi WHERE id = ?', [lokasi_id]);
    if (!lok) return res.status(404).json({ message: 'Lokasi tidak ditemukan.' });

    const meta = await db.get('SELECT catuan FROM rekap_opm_meta WHERE lokasi_id = ?', [lokasi_id]) || {};
    const adminInput = await db.get('SELECT proyek, surat_pesanan FROM laporan_rekon WHERE lokasi_id = ?', [lokasi_id]) || {};
    const header = {
      proyek: adminInput.proyek || '-',
      surat_pesanan: adminInput.surat_pesanan || '-',
      lokasi_smile: lok.nama_lokasi || '-',
      catuan: meta.catuan || '-'
    };

    const { odpMap, fotoRows } = await daftarOdpOpm(lokasi_id, kategori);
    const portRows = await db.all('SELECT odp, port, status FROM opm_port WHERE lokasi_id = ? ORDER BY port ASC', [lokasi_id]);
    const nilaiRows = await db.all('SELECT odp, port, nilai FROM rekap_opm WHERE lokasi_id = ?', [lokasi_id]);
    const pilihRows = await db.all('SELECT odp FROM rekap_opm_pilih WHERE lokasi_id = ?', [lokasi_id]);
    const pilihSet = new Set(pilihRows.map(r => r.odp));
    // Bila user belum pernah memilih, anggap semua ODP terpilih (perilaku lama).
    const adaPilihan = pilihRows.length > 0;

    const daftar = Array.from(odpMap.entries()).map(([nm, kap]) => {
      const mapFile = filePortOpm(fotoRows, nm);
      const ports = [];
      for (let p = 1; p <= kap; p++) {
        const pr = portRows.find(x => x.odp === nm && x.port === p);
        const status = pr ? pr.status : 'ukur';
        const nr = nilaiRows.find(x => x.odp === nm && x.port === p);
        ports.push({
          port: p,
          status,
          file: mapFile[p] || null,
          nilai: nr && nr.nilai != null ? nr.nilai : ''
        });
      }
      return { nama: nm, kapasitas: kap, dipilih: adaPilihan ? pilihSet.has(nm) : true, ports };
    });

    const total = await db.get('SELECT COUNT(*) as total FROM rekap_opm WHERE lokasi_id = ?', [lokasi_id]);
    res.json({ daftar, header, tersimpan: !!(total && total.total > 0), ada_pilihan: adaPilihan });
  } catch (e) {
    console.error(e);
    res.status(500).json({ message: 'Gagal memuat rekap OPM.', detail: e.message });
  }
});

app.post('/api/rekap-opm/:lokasi_id/scan', checkAuth('admin'), async (req, res) => {
  try {
    const lokasi_id = req.params.lokasi_id;
    const kategori = req.query.kategori || '-EVIDENT HASIL UKUR OPM';
    const hanyaOdp = (req.body && req.body.odp) ? String(req.body.odp) : null;
    const hanyaPort = (req.body && req.body.port != null) ? parseInt(req.body.port, 10) : null;

    const lok = await db.get('SELECT tematik, nama_lokasi FROM lokasi WHERE id = ?', [lokasi_id]);
    if (!lok) return res.status(404).json({ message: 'Lokasi tidak ditemukan.' });

    const folderFoto = path.join(
      __dirname, 'uploads',
      (lok.tematik || '').trim().replace(/[^a-zA-Z0-9_-]/g, '_'),
      (lok.nama_lokasi || String(lokasi_id)).replace(/[^a-zA-Z0-9_-]/g, '_'),
      kategori.replace(/[^a-zA-Z0-9_-]/g, '_')
    );

    const { odpMap, fotoRows } = await daftarOdpOpm(lokasi_id, kategori);
    const portRows = await db.all('SELECT odp, port, status FROM opm_port WHERE lokasi_id = ?', [lokasi_id]);

    const target = Array.from(odpMap.keys()).filter(nm => !hanyaOdp || nm === hanyaOdp);
    if (hanyaOdp && !target.length) target.push(hanyaOdp);

    const hasil = {};
    let dipindai = 0, gagal = 0;
    const paralel = !!(req.body && req.body.paralel);

    // Kumpulkan dulu foto yang benar-benar perlu dipindai. Mode paralel
    // menjalankan semuanya bersamaan (dibatasi antrean AI di rekapOpm.js),
    // mode biasa tetap berurutan untuk fallback / retry yang lebih lembut.
    const tugas = [];
    for (const nm of target) {
      const mapFile = filePortOpm(fotoRows, nm);
      const kap = odpMap.get(nm) || 8;
      for (let p = 1; p <= kap; p++) {
        if (!isNaN(hanyaPort) && hanyaPort > 0 && p !== hanyaPort) continue;
        const pr = portRows.find(x => x.odp === nm && x.port === p);
        const status = pr ? pr.status : 'ukur';
        if (status === 'isi') { hasil[nm + '|' + p] = 'ISI'; continue; }
        const file = mapFile[p];
        if (!file) continue;
        const fp = path.join(folderFoto, file);
        if (!fs.existsSync(fp)) continue;
        tugas.push({ nm, p, fp });
      }
    }
    const jalankan = async (t) => {
      try {
        const angka = await rekapOpmService.bacaAngkaOpm(t.fp, { odp: t.nm, port: t.p });
        if (angka) { hasil[t.nm + '|' + t.p] = angka; dipindai++; }
      } catch (e) {
        gagal++;
        console.warn('[REKAP OPM] scan gagal', t.nm, t.p, e.message);
      }
    };
    if (paralel) await Promise.all(tugas.map(jalankan));
    else for (const t of tugas) await jalankan(t);
    res.json({ hasil, dipindai, gagal, paralel, dipindaiTotal: tugas.length });
  } catch (e) {
    console.error(e);
    res.status(500).json({ message: 'Gagal scan OPM.', detail: e.message });
  }
});

app.post('/api/rekap-opm/:lokasi_id/simpan', checkAuth('admin'), async (req, res) => {
  try {
    const lokasi_id = req.params.lokasi_id;
    const { nilai, header, odp_dipilih } = req.body || {};

    if (header && header.catuan != null) {
      await db.run('INSERT OR REPLACE INTO rekap_opm_meta (lokasi_id, catuan, updated_at) VALUES (?, ?, CURRENT_TIMESTAMP)',
        [lokasi_id, String(header.catuan || '-')]);
    }

    // Simpan daftar ODP yang dipilih (bila dikirim)
    let pilihSet = null;
    if (Array.isArray(odp_dipilih)) {
      pilihSet = new Set(odp_dipilih.map(x => String(x)));
      await db.run('DELETE FROM rekap_opm_pilih WHERE lokasi_id = ?', [lokasi_id]);
      for (const nm of pilihSet) {
        await db.run('INSERT OR IGNORE INTO rekap_opm_pilih (lokasi_id, odp) VALUES (?, ?)', [lokasi_id, nm]);
      }
    }

    await db.run('DELETE FROM rekap_opm WHERE lokasi_id = ?', [lokasi_id]);
    const entries = [];
    if (nilai && typeof nilai === 'object') {
      for (const k of Object.keys(nilai)) {
        const parts = k.split('|');
        if (parts.length !== 2) continue;
        const port = parseInt(parts[1], 10);
        if (Number.isNaN(port)) continue;
        // Bila ada pilihan ODP, hanya simpan nilai ODP yang dipilih
        if (pilihSet && !pilihSet.has(parts[0])) continue;
        entries.push([lokasi_id, parts[0], port, String(nilai[k] == null ? '' : nilai[k])]);
      }
    }
    for (const e of entries) {
      await db.run('INSERT INTO rekap_opm (lokasi_id, odp, port, nilai) VALUES (?,?,?,?)', e);
    }
    res.json({ success: true, count: entries.length, dipilih: pilihSet ? pilihSet.size : null });
  } catch (e) {
    console.error(e);
    res.status(500).json({ message: 'Gagal menyimpan rekap OPM.', detail: e.message });
  }
});

app.post('/api/rekap-opm/:lokasi_id/generate', checkAuth('admin'), async (req, res) => {
  try {
    const lokasi_id = req.params.lokasi_id;
    const { header, odp_dipilih } = req.body || {};

    const lok = await db.get('SELECT tematik, nama_lokasi FROM lokasi WHERE id = ?', [lokasi_id]);
    if (!lok) return res.status(404).json({ message: 'Lokasi tidak ditemukan.' });

    const meta = await db.get('SELECT catuan FROM rekap_opm_meta WHERE lokasi_id = ?', [lokasi_id]) || {};
    const adminInput = await db.get('SELECT proyek, surat_pesanan FROM laporan_rekon WHERE lokasi_id = ?', [lokasi_id]) || {};
    const h = {
      proyek: adminInput.proyek || '-',
      surat_pesanan: adminInput.surat_pesanan || '-',
      lokasi_smile: lok.nama_lokasi || '-',
      catuan: (header && header.catuan != null) ? header.catuan : (meta.catuan || '-')
    };

    const { odpMap } = await daftarOdpOpm(lokasi_id, '-EVIDENT HASIL UKUR OPM');
    const portRows = await db.all('SELECT odp, port, status FROM opm_port WHERE lokasi_id = ?', [lokasi_id]);
    const nilaiRows = await db.all('SELECT odp, port, nilai FROM rekap_opm WHERE lokasi_id = ?', [lokasi_id]);

    // ODP yang ikut dilampirkan: dari body, atau dari pilihan tersimpan; bila
    // keduanya kosong → semua ODP (perilaku lama).
    let daftarPilih = null;
    if (Array.isArray(odp_dipilih) && odp_dipilih.length) {
      daftarPilih = odp_dipilih.map(String);
    } else {
      const pilihRows = await db.all('SELECT odp FROM rekap_opm_pilih WHERE lokasi_id = ?', [lokasi_id]);
      if (pilihRows.length) daftarPilih = pilihRows.map(r => r.odp);
    }
    const pilihSet = daftarPilih ? new Set(daftarPilih) : null;

    const odpsList = Array.from(odpMap.entries())
      .filter(([nm]) => !pilihSet || pilihSet.has(nm))
      .map(([nm, kap]) => {
        const nilai = [];
        for (let p = 1; p <= kap; p++) {
          const nr = nilaiRows.find(x => x.odp === nm && x.port === p);
          let v = nr && nr.nilai != null ? nr.nilai : '';
          if (!String(v).trim()) {
            const pr = portRows.find(x => x.odp === nm && x.port === p);
            if (pr && pr.status === 'isi') v = 'ISI';
          }
          nilai.push(v);
        }
        return { nama: nm, nilai };
      });

    if (!odpsList.length) {
      return res.status(400).json({ message: 'Belum ada ODP yang dipilih / tersedia pada eviden "EVIDENT HASIL UKUR OPM" untuk lokasi ini.' });
    }

    const namaTematikClean = (lok.tematik || '').trim().replace(/[^a-zA-Z0-9_-]/g, '_');
    const namaLokasiClean  = (lok.nama_lokasi || String(lokasi_id)).replace(/[^a-zA-Z0-9_-]/g, '_');
    const outDir = path.join(__dirname, 'uploads', namaTematikClean, namaLokasiClean, 'INPUT_REKAP_OPM');
    pastikanFolderEksis(outDir);

    const hasilExcel = await rekapOpmService.isiRekapExcel({ header: h, odps: odpsList, outDir, prefix: 'REKAP_OPM' });

    // Render tiap PDF → PNG halaman pertama, simpan sebagai berkas per ODP
    const berkasBaru = [];
    for (let i = 0; i < hasilExcel.pdfs.length; i++) {
      const pdf = hasilExcel.pdfs[i];
      const odpsNm = odpsList[i] ? odpsList[i].nama : ('ODP' + (i + 1));
      const aman = odpsNm.replace(/[^a-zA-Z0-9_-]/g, '_');
      const pngTujuan = path.join(outDir, aman + '.png');
      try {
        const halaman = await renderPdfHalaman(pdf.path);
        if (halaman && halaman.length) {
          fs.copyFileSync(halaman[0], pngTujuan);
          await db.run('DELETE FROM berkas WHERE lokasi_id = ? AND kategori = ? AND sub_kategori = ?',
            [lokasi_id, KATEGORI_REKAP_OPM, odpsNm]);
          await db.run('INSERT INTO berkas (lokasi_id, kategori, sub_kategori, file_name) VALUES (?,?,?,?)',
            [lokasi_id, KATEGORI_REKAP_OPM, odpsNm, path.basename(pngTujuan)]);
          berkasBaru.push({ odp: odpsNm, file: path.basename(pngTujuan) });
        }
      } catch (e) {
        console.warn('[REKAP OPM] render PDF→PNG gagal', odpsNm, e.message);
      }
    }

    // Hapus lampiran lama untuk ODP yang TIDAK dipilih (agar tidak ikut di LACT)
    const berkasLama = await db.all(
      'SELECT id, sub_kategori, file_name FROM berkas WHERE lokasi_id = ? AND kategori = ?',
      [lokasi_id, KATEGORI_REKAP_OPM]
    );
    for (const b of berkasLama) {
      const sub = (b.sub_kategori || '').trim();
      if (sub === 'Excel Rekap') continue;
      if (pilihSet && !pilihSet.has(sub)) {
        try { fs.rmSync(path.join(outDir, b.file_name), { force: true }); } catch (e) { /* biarkan */ }
        await db.run('DELETE FROM berkas WHERE id = ?', [b.id]);
      }
    }

    // Simpan juga file Excel hasil isi (bisa diunduh)
    await db.run('DELETE FROM berkas WHERE lokasi_id = ? AND kategori = ? AND sub_kategori = ?',
      [lokasi_id, KATEGORI_REKAP_OPM, 'Excel Rekap']);
    await db.run('INSERT INTO berkas (lokasi_id, kategori, sub_kategori, file_name) VALUES (?,?,?,?)',
      [lokasi_id, KATEGORI_REKAP_OPM, 'Excel Rekap', path.basename(hasilExcel.xlsx)]);

    await catatAktivitas('rekap_opm', lokasi_id);
    res.json({ success: true, xlsx: path.basename(hasilExcel.xlsx), berkas: berkasBaru });
  } catch (e) {
    console.error(e);
    res.status(500).json({ message: 'Gagal generate rekap OPM.', detail: e.message });
  }
});

// ── MITRATEL: route tambahan mandiri (OPM; ATP menyusul setelah template tersedia) ──
mitratelService.register(app, checkAuth, { renderPdf: renderPdfDwg });

// ── Single-instance guard: cegah dua server berjalan bersamaan (penyebab saveDb EBUSY/ENOENT) ──
const _LOCK_FILE = path.join(__dirname, 'server.lock');
(function _singleInstanceGuard() {
  try {
    if (fs.existsSync(_LOCK_FILE)) {
      const pid = parseInt(fs.readFileSync(_LOCK_FILE, 'utf8').trim(), 10);
      if (pid && pid !== process.pid) {
        let alive = false;
        try { process.kill(pid, 0); alive = true; } catch (e) { alive = false; }
        if (alive) {
          console.error(`\n⚠ Server sudah berjalan (PID ${pid}).\n  Hentikan proses tersebut sebelum menjalankan server baru, agar database tidak terkunci (EBUSY/ENOENT).\n  Contoh: Stop-Process -Id ${pid} -Force\n`);
          process.exit(1);
        }
      }
    }
  } catch (e) { /* abaikan */ }
  try { fs.writeFileSync(_LOCK_FILE, String(process.pid)); } catch (e) { /* abaikan */ }
})();
process.on('exit', () => { try { if (fs.existsSync(_LOCK_FILE) && parseInt(fs.readFileSync(_LOCK_FILE, 'utf8').trim(), 10) === process.pid) fs.unlinkSync(_LOCK_FILE); } catch (e) { /* abaikan */ } });
process.on('SIGINT', () => process.exit(0));
process.on('SIGTERM', () => process.exit(0));
process.on('uncaughtException', (e) => { console.error('uncaughtException:', e && e.message ? e.message : e); /* tetap exit */ try { process.exit(1); } catch (_) {} });

app.listen(PORT, () => {
  console.log(`\n==================================================================`);
  console.log(`  Sistem Aktif Berjalan di: http://localhost:${PORT}/login.html`);
  console.log(`==================================================================\n`);
  // Bersihkan sesi lama/kedaluwarsa agar tabel sessions tidak menumpuk.
  bersihkanSesiLama()
    .then(n => { if (n) console.log(`[sesi] ${n} sesi lama dibersihkan saat startup.`); })
    .catch(e => console.warn('[sesi] gagal membersihkan sesi lama:', e && e.message || e));
  setInterval(() => {
    bersihkanSesiLama().catch(e => console.warn('[sesi] gagal membersihkan sesi lama:', e && e.message || e));
  }, 60 * 60 * 1000);
  // Cadangkan eviden ke Google Drive (bila remote rclone sudah diotorisasi).
  // Tidak menghalangi startup: kegagalan di sini hanya dicatat di log.
  driveSync.mulai().catch(e => console.warn('[DriveSync] gagal mulai:', e && e.message || e));
  // Sinkronisasi master project dari Google Sheets "Data Master"
  // (semua 11.649 baris; kolom G = PROGRAM → tematik, kolom AF = status lapangan).
  // Dinonaktifkan bila DISABLE_GODC_SYNC=1 (uji otomatis): cleanup-nya menghapus
  // lokasi yang tidak ada di sheet sehingga baris fixture uji ikut terhapus.
  if (process.env.DISABLE_GODC_SYNC === '1') {
    console.log('[GODC] Sinkronisasi dinonaktifkan (DISABLE_GODC_SYNC=1).');
  } else {
    godcSync.start();

    // Health check: restart sync if it stops unexpectedly
    setInterval(() => {
      if (!godcSync.isRunning()) {
        console.log('[GODC] Health check: sync stopped, restarting...');
        godcSync.start();
      }
    }, 120000);
  }
});

// ── GODC SYNC (Google Sheets "Data Master" — Semesta OnHand & Revenue Area-2) ─
// Kolom AF (Status Lapangan PMO-2) menentukan proyek yang masuk tahap evidentiary.
// Kolom G (PROGRAM) menjadi tematik.
function needEvidence(it) {
  if (!it) return false;
  if (it.butuh_eviden === 'YA') return true;
  if (it.butuh_eviden === 'TIDAK') return false;
  return godcSync.isTriggerStatus(it.status_progress);
}

app.get('/api/godc', async (req, res) => {
  try {
    // Hanya kolom yang dipakai UI. Nama lokasi di Data Master bisa sangat panjang
    // (deskripsi SPK), jadi dipotong agar payload tetap ringan untuk 11.649 baris.
    const rows = await db.all(`SELECT
        sumber, sumber_rev, pid AS id_pid, substr(nama_lop, 1, 200) AS nama_lop, tematik_project, branch,
        nama_mitra, status_progress, butuh_eviden, status_berkas, pic_berkas,
        nilai_total_drm, po_non_po, drm_nilai, substr(id_sap, 1, 60) AS id_sap
      FROM godc_project ORDER BY nama_lop`);
    // Satu lokasi = satu unit kerja: baris kembar digabung sebelum dihitung.
    const items = godcSync.uniqByLokasi(rows);
    const perTematik = await db.all('SELECT tematik_project AS k, COUNT(*) AS n, SUM(drm_nilai) AS drm FROM godc_project WHERE tematik_project != "" GROUP BY tematik_project ORDER BY n DESC');
    const perStatus  = await db.all('SELECT status_progress AS k, COUNT(*) AS n FROM godc_project WHERE status_progress != "" GROUP BY status_progress ORDER BY n DESC');
    const perSumber  = await db.all('SELECT sumber AS k, COUNT(*) AS n FROM godc_project GROUP BY sumber');
    const perRev     = await db.all("SELECT sumber_rev AS k, COUNT(*) AS n FROM godc_project WHERE COALESCE(sumber_rev,'') != '' GROUP BY sumber_rev ORDER BY n DESC");
    const perBranch  = await db.all("SELECT branch AS k, COUNT(*) AS n FROM godc_project WHERE COALESCE(branch,'') != '' GROUP BY branch ORDER BY n DESC");
    const meta      = await db.all("SELECT key, value FROM godc_meta");
    const metaObj   = {};
    meta.forEach(m => { metaObj[m.key] = m.value; });

    const lokasiRows = await db.all('SELECT l.id, l.nama_lokasi, l.status_eviden, (SELECT COUNT(*) FROM berkas b WHERE b.lokasi_id = l.id) AS n_berkas FROM lokasi l');
    const lokasiMap  = new Map(lokasiRows.map(r => [godcSync.normName(r.nama_lokasi), r]));
    const perluEviden = items.filter(it => {
      if (!needEvidence(it)) return false;
      const m = lokasiMap.get(godcSync.normName(it.nama_lop));
      return !m || m.n_berkas === 0;
    }).map(it => {
      const m = lokasiMap.get(godcSync.normName(it.nama_lop));
      return Object.assign({}, it, { lokasi_id: m ? m.id : null, n_berkas: m ? m.n_berkas : 0, status_eviden: m ? m.status_eviden : null });
    });

    // No. SP (PO) — gabungkan ke tiap item + ringkasan PO.
    // Hanya proyek PO yang statusnya sudah FINISH/GOLIVE yang wajib input No. SP.
    const spRows = await db.all("SELECT nama_lokasi, no_sp FROM godc_sp WHERE COALESCE(no_sp,'') <> ''");
    const spMap  = new Map(spRows.map(r => [godcSync.normName(r.nama_lokasi), r.no_sp]));
    const enriched = items.map(it => {
      const m = lokasiMap.get(godcSync.normName(it.nama_lop));
      return Object.assign({}, it, {
        no_sp: spMap.get(godcSync.normName(it.nama_lop)) || '',
        lokasi_id: m ? m.id : null,
        status_eviden: m ? m.status_eviden : null
      });
    });
    const poFinish = enriched.filter(it => godcSync.poClass(it.po_non_po) === 'PO' && needEvidence(it));
    const sp = {
      po: poFinish.length,
      poTerisi: poFinish.filter(it => it.no_sp).length,
      poKosong: poFinish.filter(it => !it.no_sp).length
    };

    // Trigger stats: berapa banyak project yang match trigger status
    const allTriggerProjects = enriched.filter(it => needEvidence(it));
    const triggerInLokasi = allTriggerProjects.filter(it => {
      const m = lokasiMap.get(godcSync.normName(it.nama_lop));
      return !!m;
    });
    const triggerNotInLokasi = allTriggerProjects.filter(it => {
      const m = lokasiMap.get(godcSync.normName(it.nama_lop));
      return !m;
    });

    res.json({
      ok: true,
      sync: Object.assign({}, godcSync.state, { lastSyncMeta: metaObj.last_sync || null }),
      sheet: { nama: godcSync.DATA_SHEET_NAME, url: godcSync.sheetUrl(0) },
      ringkasan: { perTematik, perStatus, perSumber, perRev, perBranch },
      perluEviden,
      lokasiTerdaftar: lokasiRows.length,
      triggerStats: {
        totalTrigger: allTriggerProjects.length,
        triggerInLokasi: triggerInLokasi.length,
        triggerNotInLokasi: triggerNotInLokasi.length,
        triggerNotInLokasiNames: triggerNotInLokasi.map(it => it.nama_lop)
      },
      sp,
      items: enriched
    });
  } catch (e) {
    res.status(500).json({ ok: false, message: 'Gagal memuat data GODC', detail: e.message });
  }
});

// ─────────────────────────────────────────────────────────────────────────────
// [API] REKAP STATUS — Tabel rekapitulasi per Tematik → SP → PO/Non PO
// Sumber: godc_project (Google Sheets). Kategori status:
//   - Done Pemberkasan : berkas lengkap 100%
//   - Finish Instalasi : status GOLIVE/FINISH tapi berkas belum 100%
//   - Need Evidence    : masih butuh kelengkapan eviden (berkas < 100%, belum finish)
// Setiap sel menampilkan: {n} LOP | Rp {m} M
// ─────────────────────────────────────────────────────────────────────────────
function godcTematikRowServer(t) {
  const v = String(t || '').trim().toUpperCase();
  if (/HEM/.test(v)) return 'HEM';
  if (/QEREL|RELOKASI|RELOLASI/.test(v) && !/PROYEK/.test(v)) return 'Relokasi';
  if (/MICRODEMAND/.test(v)) return 'MICRODEMAND';
  if (/NODE\s*B/.test(v)) return 'Node B';
  if (/MITRATEL|MORATEL|M-?\s?TEL/.test(v)) return 'M-tel';
  if (/REPUBLIK|MY\s*REP/.test(v)) return 'MyRepublik';
  if (/^JLM|JL\b/.test(v)) return 'Jlm';
  if (/OLO\b/.test(v)) return 'OLO';
  if (/UNSC|PT2/.test(v)) return 'PT2 UNSC';
  if (/IFORTE|FOR3E/.test(v)) return 'IFORTE';
  return t || '(tanpa tematik)';
}

  app.get('/api/rekap-status', async (req, res) => {
  try {
    // Filter BRANCH opsional (?branch=). Nilai dinormalisasi agar "Bandung",
    // "BANDUNG", dan "jaksel"/"JAKARTA SELATAN" regarded cabang yang sama.
    const reqBranch = godcSync.normBranch(req.query.branch || '');
    const reqBranchKey = reqBranch ? godcSync.normBranchKey(reqBranch) : '';
    // Filter STATUS PEKERJAAN opsional (?status=) — nilai persis seperti kolom AF
    // "Status Lapangan PMO-2" (mis. "03. MATDEV", "04. INSTALASI", "07. TC").
    const reqStatus = String(req.query.status || '').trim();
    const rows = await db.all('SELECT id, sumber, sumber_rev, branch, nama_lop, tematik_project, status_progress, butuh_eviden, po_non_po, nilai_total_drm, drm_nilai FROM godc_project');
    const items = godcSync.uniqByLokasi(rows);
    const spRows = await db.all("SELECT nama_lokasi, no_sp FROM godc_sp WHERE COALESCE(no_sp,'') <> ''");
    const spMap = new Map(spRows.map(r => [godcSync.normName(r.nama_lokasi), r.no_sp]));

    const lokasiRows = await db.all('SELECT l.id, l.nama_lokasi, l.tematik, l.status_eviden FROM lokasi l');
    const lokasiByNorm = new Map();
    lokasiRows.forEach(l => lokasiByNorm.set(godcSync.normName(l.nama_lokasi), l));

    // Status 6 dokumen pemberkasan per LOP (tabel dokumen_status).
    // Dipakai kolom Progres pada tabel rekap, panel Detail LOP & dashboard dokumen.
    await seedDokumenStatus();
    const petaDok = await petaDokumen();

    // Peta kelengkapan berkas per lokasi (replikasi logika /api/progres)
    const progresByLokasi = new Map();
    if (lokasiRows.length) {
      const lokasiIds = lokasiRows.map(l => l.id);
      const berkasRows = await db.all(
        `SELECT lokasi_id, kategori FROM berkas WHERE lokasi_id IN (${lokasiIds.map(() => '?').join(',')})`,
        lokasiIds
      );
      const inputRows = await db.all(
        `SELECT lokasi_id, kategori FROM input_khusus WHERE lokasi_id IN (${lokasiIds.map(() => '?').join(',')})`,
        lokasiIds
      );
      const boqRows = await db.all(
        `SELECT lokasi_id, COUNT(*) AS n FROM boq_aktual WHERE lokasi_id IN (${lokasiIds.map(() => '?').join(',')}) GROUP BY lokasi_id`,
        lokasiIds
      );
      const byLok = {};
      const inputByLok = {};
      for (const b of berkasRows) { (byLok[b.lokasi_id] = byLok[b.lokasi_id] || new Set()).add(b.kategori); }
      for (const b of inputRows) { (inputByLok[b.lokasi_id] = inputByLok[b.lokasi_id] || new Set()).add(b.kategori); }
      const boqSet = new Set(boqRows.map(b => b.lokasi_id));
      for (const l of lokasiRows) {
        const daftarWajib = DATABASE_TEMATIK_WAJIB[l.tematik] || [];
        const setWajib = new Set(daftarWajib);
        const kategoriAda = Array.from(byLok[l.id] || [])
          .concat(Array.from(inputByLok[l.id] || []))
          .concat(boqSet.has(l.id) ? ['BOQ AKTUAL'] : []);
        const sudah = new Set(kategoriAda.filter(k => setWajib.has(k))).size;
        const total = daftarWajib.length;
        progresByLokasi.set(l.id, total > 0 ? Math.round((sudah / total) * 100) : 0);
      }
    }

    // Wadah agregasi: tematik -> { spKey -> { po:{finish,need,done}, nonpo:{finish,need,done} }, total:{...} }
    const catKey = (po) => (po === 'PO' ? 'po' : 'nonpo');
    const statusKey = ['finish', 'need', 'done'];
    const addCell = (bucket, po, st, drm) => {
      const k = catKey(po);
      bucket[k][st].n += 1;
      bucket[k][st].drm += drm;
    };

    const tree = new Map(); // sumber -> Map(tematik -> Map(spKey -> bucket))
    const lokasiByKey = new Map(); // "sumber|tematik" -> array lokasi
    let totalAll = { po: { finish:{n:0,drm:0}, need:{n:0,drm:0}, done:{n:0,drm:0} }, nonpo: { finish:{n:0,drm:0}, need:{n:0,drm:0}, done:{n:0,drm:0} } };

    // Daftar BRANCH & STATUS PEKERJAAN untuk mengisi dropdown filter. Dihitung
    // dari SELURUH item yang lolos trigger (sebelum filter), supaya pilihan tidak
    // ikut berubah saat cabang/status lain sedang difilter.
    const branchList = new Map();
    const statusList = new Map();
    for (const it of items) {
      if (!needEvidence(it)) continue;
      const b = godcSync.normBranch(it.branch) || '(Tanpa Branch)';
      branchList.set(b, (branchList.get(b) || 0) + 1);
      const sk = String(it.status_progress || '').trim() || '(Tanpa Status)';
      statusList.set(sk, (statusList.get(sk) || 0) + 1);
    }
    const branches = Array.from(branchList.entries())
      .sort(function (a, b) { return b[1] - a[1] || a[0].localeCompare(b[0]); })
      .map(function (e) { return { branch: e[0], n: e[1] }; });
    const statuses = Array.from(statusList.entries())
      .sort(function (a, b) { return b[1] - a[1] || a[0].localeCompare(b[0]); })
      .map(function (e) { return { status: e[0], n: e[1] }; });

    for (const it of items) {
      // Sumber rekap mengikuti kolom E "Rev" (INTERNAL / EKSTERNAL), bukan
      // kolom B "Source Data" (Apras/Non Apras) yang hanya dipakai sebagai badge.
      const sumber = (it.sumber_rev || '').trim() || 'Data Master';
      const tematik = godcTematikRowServer(it.tematik_project);
      const noSp = spMap.get(godcSync.normName(it.nama_lop)) || '';
      const po = godcSync.poClass(it.po_non_po) === 'PO' ? 'PO' : 'NON PO';
      const drm = Number(it.drm_nilai) || Number(it.nilai_total_drm) || 0;
      const branch = godcSync.normBranch(it.branch);

      const lok = lokasiByNorm.get(godcSync.normName(it.nama_lop));
      const progres = lok ? (progresByLokasi.get(lok.id) || 0) : 0;
      const isFinish = needEvidence(it);

      // Rekap hanya menampilkan proyek yang sudah final (kolom AF = FINISH
      // INSTALASI / OA atau GO LIVE). Proyek yang masih berjalan tidak dihitung.
      if (!isFinish) continue;

      // Filter BRANCH: hanya untuk memilih cabang, bukan mengubah hitungan.
      if (reqBranchKey && godcSync.normBranchKey(branch) !== reqBranchKey) continue;

      // Filter STATUS PEKERJAAN (kolom AF): "03. MATDEV", "04. INSTALASI", dst.
      if (reqStatus && String(it.status_progress || '').trim() !== reqStatus) continue;

      let st;
      // Gunakan status_eviden jika ada (manual override)
      if (lok && lok.status_eviden) {
        if (lok.status_eviden === 'done') st = 'done';
        else if (lok.status_eviden === 'progres') st = 'finish';
        else st = 'need';
      } else {
        // Default: hitung dari kelengkapan berkas
        if (progres >= 100) st = 'done';       // berkas lengkap -> Done Pemberkasan
        else if (progres > 0) st = 'finish';   // sudah golive & punya sebagian berkas -> Progres
        else st = 'need';                      // golive tapi belum punya eviden -> Need Evidence
      }

      const dokLok = (lok && petaDok.get(lok.id)) || {};
      const dok = {}, genAt = {}, gen = [];
      DOK_STATUS_JENIS.forEach(function (j) {
        const stDok = dokLok[j] ? dokLok[j].status : 'BELUM';
        dok[j] = stDok;
        genAt[j] = dokLok[j] ? dokLok[j].generated_at : null;
        if (stDok === 'TERGENERATE') gen.push(j);
      });

      const sampKey = sumber + '|' + tematik;
      if (!lokasiByKey.has(sampKey)) lokasiByKey.set(sampKey, []);
      lokasiByKey.get(sampKey).push({
        nama: it.nama_lop, sp: noSp, po: catKey(po), st, drm, branch,
        status_progress: String(it.status_progress || '').trim(),
        status_eviden: lok ? lok.status_eviden : null, lokasi_id: lok ? lok.id : null,
        gen: gen, dok: dok, gen_at: genAt
      });

      if (!tree.has(sumber)) tree.set(sumber, new Map());
      const byTem = tree.get(sumber);
      if (!byTem.has(tematik)) byTem.set(tematik, new Map());
      const spBuckets = byTem.get(tematik);
      const spKey = noSp || '(Tanpa SP)';
      if (!spBuckets.has(spKey)) {
        spBuckets.set(spKey, { sp: noSp, po: { finish:{n:0,drm:0}, need:{n:0,drm:0}, done:{n:0,drm:0} }, nonpo: { finish:{n:0,drm:0}, need:{n:0,drm:0}, done:{n:0,drm:0} } });
      }
      addCell(spBuckets.get(spKey), po, st, drm);
      addCell(totalAll, po, st, drm);
    }

    const sumberOut = [];
    for (const [sumber, byTem] of tree.entries()) {
      const tematikOut = [];
      for (const [tematik, spBuckets] of byTem.entries()) {
        const sps = Array.from(spBuckets.values())
          .filter(b => b.sp || b.po.finish.n + b.po.need.n + b.po.done.n + b.nonpo.finish.n + b.nonpo.need.n + b.nonpo.done.n > 0)
          .map(b => {
            const spTot = { po: {finish:{n:0,drm:0},need:{n:0,drm:0},done:{n:0,drm:0}}, nonpo: {finish:{n:0,drm:0},need:{n:0,drm:0},done:{n:0,drm:0}} };
            for (const k of ['po','nonpo']) for (const s of statusKey) {
              spTot[k][s].n = b[k][s].n;
              spTot[k][s].drm = b[k][s].drm;
            }
            return { sp: b.sp, ...spTot };
          })
          .sort((a, b) => (a.sp || '').localeCompare(b.sp || ''));
        const tot = { po: {finish:{n:0,drm:0},need:{n:0,drm:0},done:{n:0,drm:0}}, nonpo: {finish:{n:0,drm:0},need:{n:0,drm:0},done:{n:0,drm:0}} };
        for (const b of spBuckets.values()) for (const k of ['po','nonpo']) for (const s of statusKey) {
          tot[k][s].n += b[k][s].n;
          tot[k][s].drm += b[k][s].drm;
        }
        const lokasi = (lokasiByKey.get(sumber + '|' + tematik) || []).slice().sort((a, b) => (a.nama || '').localeCompare(b.nama || ''));
        tematikOut.push({ tematik, sps, total: tot, lokasi });
      }
      tematikOut.sort((a, b) => (a.tematik || '').localeCompare(b.tematik || ''));
      sumberOut.push({ sumber, tematik: tematikOut });
    }
    sumberOut.sort((a, b) => (a.sumber || '').localeCompare(b.sumber || ''));

    res.json({
      ok: true,
      updated_at: new Date().toISOString(),
      filter_branch: reqBranch || '',
      filter_status: reqStatus || '',
      branches,
      statuses,
      total: totalAll,
      sumber: sumberOut
    });
  } catch (e) {
    res.status(500).json({ ok: false, message: 'Gagal memuat rekap status', detail: e.message });
  }
});

app.get('/api/godc/sync/status', (req, res) => {
  res.json({
    ok: true,
    running: godcSync.isRunning(),
    state: godcSync.state
  });
});

app.post('/api/godc/sync', async (req, res) => {
  try {
    const st = await godcSync.runSync({ cleanup: req.body && req.body.cleanup === true });
    const diag = st.autoCreateDiagnostic || null;
    res.json({
      ok: true,
      sync: st,
      diagnostic: diag,
      cleanup: st.cleanup || null
    });
  } catch (e) {
    res.status(500).json({ ok: false, message: 'Gagal sinkron GODC', detail: e.message });
  }
});

// Lihat/hapus lokasi yang namanya sudah tidak ada di Data Master.
//   ?dryRun=1 → simulasi (default), ?force=1 → benar-benar hapus
// Force juga membersihkan baris tabel anak yang yatim (induk lokasi sudah hilang).
app.post('/api/godc/cleanup', checkAuth('admin'), async (req, res) => {
  try {
    const force = req.query.force === '1' || (req.body && req.body.force === true);
    const dryRun = req.query.dryRun === '1' || !force;
    if (force) {
      const st = await godcSync.runSync({ cleanup: true });
      if (st.lastError) return res.status(500).json({ ok: false, message: 'Gagal sinkron GODC', detail: st.lastError });
      return res.json({ ok: true, cleanup: st.cleanup || null });
    }
    const stats = await godcSync.readSheets();
    await godcSync.ensureSchema();
    const info = await godcSync.cleanupOrphanLokasi(stats.merged, { dryRun: true });
    const yatim = await godcSync.countOrphanChildren();
    res.json({ ok: true, rowsRead: stats.rowsRead, cleanup: info, orphanChildren: yatim });
  } catch (e) {
    res.status(500).json({ ok: false, message: 'Gagal memproses pembersihan lokasi', detail: e.message });
  }
});

// Simpan No. SP proyek (PO) — tulis ke godc_sp + lokasi yang cocok
app.post('/api/godc/sp', checkAuth('admin'), async (req, res) => {
  try {
    const { nama_lop, no_sp } = req.body || {};
    if (!nama_lop) return res.status(400).json({ ok: false, message: 'nama_lop wajib diisi.' });
    const hasil = await godcSync.saveSp(nama_lop, String(no_sp == null ? '' : no_sp).trim());
    res.json({ ok: true, ...hasil });
  } catch (e) {
    res.status(500).json({ ok: false, message: 'Gagal menyimpan No. SP', detail: e.message });
  }
});

// ── REPORT GODC — baca tab Dasboard 2026 (gid 0) → tampilan rapi di menu Report ──
const reportCache = { t: 0, gid: 0, rows: null };
app.get('/api/godc/report', async (req, res) => {
  try {
    const gid = Number(req.query.gid) || 0;
    const force = req.query.force === '1';
    const now = Date.now();
    if (!reportCache.rows || reportCache.gid !== gid || force || now - reportCache.t > 60000) {
      reportCache.rows = await godcSync.readReport(gid);
      reportCache.gid = gid;
      reportCache.t = now;
    }
    res.json({
      ok: true,
      sheet: gid === 0 ? 'Dasboard 2026' : 'gid ' + gid,
      gid,
      url: `https://docs.google.com/spreadsheets/d/${godcSync.SPREADSHEET_ID}/edit?gid=${gid}#gid=${gid}`,
      fetchedAt: new Date(reportCache.t).toISOString(),
      rows: reportCache.rows
    });
  } catch (e) {
    res.status(500).json({ ok: false, message: 'Gagal memuat report GODC', detail: e.message });
  }
});

// ── Status pencadangan eviden ke Google Drive ────────────────────────────────
// Dipakai admin untuk melihat antrean/hasil sinkronisasi (services/driveSync.js).
app.get('/api/drive/status', checkAuth(), (req, res) => {
  res.json(driveSync.status());
});

// ── ERROR HANDLER GLOBAL (multer & lainnya) ──────────────────────────────────
// Memastikan error upload (file terlalu besar, field salah, dll) dikembalikan
// sebagai JSON yang jelas, bukan halaman HTML bawaan Express.
app.use((err, req, res, next) => {
  const pesan = String((err && err.message) || err || 'Kesalahan tidak diketahui');
  console.error('[Global Error]', pesan);
  if (res.headersSent) return next(err);

  // File kelebihan ukuran (multer LIMIT_FILE_SIZE / "File too large").
  // Kembalikan pesan yang memuat angka batasnya supaya user tahu harus
  // memperkecil file — jangan tampilkan detail teknis "File too large".
  const batasMb = (err && err.limit) ? Math.round(err.limit / (1024 * 1024)) : 15;
  if (err && (err.code === 'LIMIT_FILE_SIZE' || err.name === 'MulterError' && /too large|LIMIT_FILE_SIZE/i.test(pesan))) {
    return res.status(400).json({
      message: `Berkas terlalu besar: melebihi batas ${batasMb} MB per file. Perkecil fotonya dulu (kompres/ubah resolusi) lalu upload ulang.`,
      code: 'FILE_TOO_LARGE',
      limit_mb: batasMb,
    });
  }
  if (err && (err.name === 'MulterError' || pesan.includes('File terlalu besar') || pesan.includes('LIMIT_FILE_SIZE'))) {
    return res.status(400).json({
      message: 'Berkas gagal dikirim: ukuran file melebihi batas atau format tidak sesuai.',
      detail: pesan,
    });
  }
  res.status(err && err.status || 500).json({ message: 'Terjadi kesalahan pada server.', detail: pesan });
});
