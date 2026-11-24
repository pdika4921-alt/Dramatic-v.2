/* =====================================================================
   boqDrmMaster.js — BACA SAJA data master BOQ DRM ALL
   ---------------------------------------------------------------------
   Dipakai oleh tombol "CEK DATA MASTER BOQ" di panel Input RAB/DRM dan
   Input Aktual/REKON (halaman RAB & Rekonsiliasi).

   SUMBER UTAMA (default): Google Sheet "BOQ AMD4 PAKET 5"
     https://docs.google.com/spreadsheets/d/1j5ORzwBqRpLW2oTtYcQTGMxrfAu8XRbbZxRxT_dYnAw/edit?gid=1487274428
   FALLBACK               : templates/BOQ DRM ALL.xlsx (file lokal)

   Aturan sheet (sama untuk kedua sumber, TIDAK diubah):
     • baris 4 : NAMA LOKASI, mulai kolom K, tiap lokasi lebar 4 kolom
                 (K, O, S, W, …)
     • baris 5 : sub-header DRM | REKON | TAMBAH | KURANG
     • baris 6+: C = designator M-, D = designator J-, F = satuan,
                 G = harga material, H = harga jasa,
                 4 kolom pertama tiap blok = volume DRM / REKON.

   Google Sheet dibaca lewat export CSV publik (gviz) — TANPA login,
   TANPA service account, TANPA OAuth. File master tidak pernah ditulis.
   Hasil di-cache di memori (TTL default 30 menit) agar cepat.

   Env opsional:
     BOQ_DRM_MASTER_SOURCE = google | file   (default: google)
     BOQ_DRM_MASTER_TTL_MS = ms cache Google (default: 1800000)
     BOQ_DRM_GOOGLE_ID / BOQ_DRM_GOOGLE_GID = ganti sheet
   ===================================================================== */
'use strict';

const fs = require('fs');
const path = require('path');
const PizZip = require('pizzip');
const { bacaShared, lepasTeks } = require('./boqTamkur');

const PATH_MASTER = path.join(__dirname, '..', 'templates', 'BOQ DRM ALL.xlsx');
const NAMA_SHEET = 'BOQ AMD4 PAKET 5';
const KOLOM_AWAL = 'K';
const LEBAR_BLOK = 4;

const SUMBER = String(process.env.BOQ_DRM_MASTER_SOURCE || 'google').toLowerCase();
const GOOGLE_ID = process.env.BOQ_DRM_GOOGLE_ID || '1j5ORzwBqRpLW2oTtYcQTGMxrfAu8XRbbZxRxT_dYnAw';
const GOOGLE_GID = process.env.BOQ_DRM_GOOGLE_GID || '1487274428';
const GOOGLE_URL = 'https://docs.google.com/spreadsheets/d/' + GOOGLE_ID +
  '/gviz/tq?tqx=out:csv&gid=' + GOOGLE_GID;
const TTL_MS = Math.max(60000, Number(process.env.BOQ_DRM_MASTER_TTL_MS) || 1800000);

let _cache = null;
let _cacheKunci = '';
let _loading = null;
let _loadingKunci = '';

// ── helper umum ───────────────────────────────────────────────────────────
function keIndeks(kolom) {
  let n = 0;
  for (const ch of String(kolom).toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n;
}
function keKolom(n) {
  let s = '';
  while (n > 0) {
    const sisa = (n - 1) % 26;
    s = String.fromCharCode(65 + sisa) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}
function lepasEntitas(s) {
  return String(s || '')
    .replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"').replace(/&apos;/g, "'");
}
function normNama(s) {
  return String(s == null ? '' : s)
    .toUpperCase()
    .replace(/[^A-Z0-9]/g, '')
    .replace(/^\d+/, ''); // "1.Relokasi …" → "RELOKASI…"
}
function angka(v) {
  if (v == null || v === '') return 0;
  let s = String(v).trim();
  if (!s) return 0;
  if (/^-?\d+(\.\d+)?$/.test(s)) return Number(s);
  s = s.replace(/[^\d,.-]/g, '');
  if (!s) return 0;
  if (s.includes(',') && s.includes('.')) {
    return Number(s.lastIndexOf(',') > s.lastIndexOf('.')
      ? s.replace(/\./g, '').replace(',', '.')
      : s.replace(/,/g, '')) || 0;
  }
  if (s.includes(',')) return Number(s.replace(/,/g, '')) || 0;
  return Number(s.replace(/\./g, '')) || 0;
}

// Kumpulkan blok lokasi dariBARIS 4 (nama) + BARIS 5 (sub-header).
function susunLokasi(baris4, baris5) {
  const lokasi = [];
  const awal = keIndeks(KOLOM_AWAL);
  for (let i = 0; i < 300; i++) {
    const basis = awal + i * LEBAR_BLOK; // 1-based (K = 11)
    const p = baris4[basis - 1];
    const nama = String(p == null ? '' : p).trim();
    if (!nama) {
      const q = baris4[basis + LEBAR_BLOK - 1];
      if (i > 0 && !String(q == null ? '' : q).trim()) break;
      continue;
    }
    const sub = [0, 1, 2, 3].map(k =>
      String(baris5[basis + k - 1] == null ? '' : baris5[basis + k - 1]).trim().toUpperCase());
    lokasi.push({
      nama,
      namaNorm: normNama(nama),
      kolom: keKolom(basis),
      drm: keKolom(basis),
      rekon: keKolom(basis + 1),
      tambah: keKolom(basis + 2),
      kurang: keKolom(basis + 3),
      sub,
      items: []
    });
  }
  if (!lokasi.length) throw new Error('Tidak ada nama lokasi di baris 4 sheet ' + NAMA_SHEET + '.');
  return lokasi;
}

// Isi item per lokasi dari baris data (array of arraysel).
// get(row, kolomKe) → nilai sel.
function isiItem(lokasi, dataRows, get) {
  for (const lok of lokasi) {
    const indeksDrm = keIndeks(lok.drm) - 1;
    const indeksRekon = keIndeks(lok.rekon) - 1;
    const agregat = new Map();
    for (const row of dataRows) {
      const vDrm = angka(get(row, indeksDrm));
      const vRek = angka(get(row, indeksRekon));
      if (!vDrm && !vRek) continue;

      const desM = String(get(row, 2) || '').trim().toUpperCase(); // C
      const desJ = String(get(row, 3) || '').trim().toUpperCase(); // D
      const satuan = String(get(row, 5) || '').trim();             // F
      const hMat = angka(get(row, 6));                             // G
      const hJas = angka(get(row, 7));                             // H
      const pasangan = [];
      if (/^M-/.test(desM)) pasangan.push({ des: desM, jasa: 0, material: hMat });
      if (/^J-/.test(desJ)) pasangan.push({ des: desJ, jasa: hJas, material: 0 });
      if (!pasangan.length) continue;

      for (const p of pasangan) {
        let a = agregat.get(p.des);
        if (!a) {
          a = { designator: p.des, jasa: 0, material: 0, satuan: '', drm: 0, rekon: 0 };
          agregat.set(p.des, a);
        }
        if (!a.jasa && p.jasa) a.jasa = p.jasa;
        if (!a.material && p.material) a.material = p.material;
        if (!a.satuan && satuan) a.satuan = satuan;
        a.drm += vDrm;
        a.rekon += vRek;
      }
    }
    lok.items = Array.from(agregat.values());
  }
}

// ── SUMBER 1: Google Sheet (CSV publik) ───────────────────────────────────
function parseCsv(teks) {
  const rows = [];
  let row = [], field = '', quoted = false;
  const s = String(teks || '');
  for (let i = 0; i < s.length; i++) {
    const c = s[i];
    if (quoted) {
      if (c === '"') {
        if (s[i + 1] === '"') { field += '"'; i++; }
        else quoted = false;
      } else field += c;
    } else if (c === '"') quoted = true;
    else if (c === ',') { row.push(field); field = ''; }
    else if (c === '\n' || c === '\r') {
      if (c === '\r' && s[i + 1] === '\n') i++;
      row.push(field); rows.push(row); row = []; field = '';
    } else field += c;
  }
  if (field || row.length) { row.push(field); rows.push(row); }
  return rows;
}

async function ambilTeks(url) {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), 25000);
  try {
    const r = await fetch(url, { signal: controller.signal });
    if (!r.ok) throw new Error('HTTP ' + r.status + ' saat mengambil Google Sheet.');
    return await r.text();
  } finally { clearTimeout(timer); }
}

async function muatMasterGoogle(kunci) {
  const rentang = '&range=A4:APZ';
  // Baris 4/5 diambil terpisah supaya nama lokasi tidak tercampur
  // dengan header grup (gviz CSV menggabungkan baris bila tanpa range).
  const [teks4, teks5, teksPenuh] = await Promise.all([
    ambilTeks(GOOGLE_URL + rentang + '4'),
    ambilTeks(GOOGLE_URL + rentang + '5'),
    ambilTeks(GOOGLE_URL)
  ]);
  const baris4 = (parseCsv(teks4)[0] || []);
  const baris5 = (parseCsv(teks5)[0] || []);
  const csv = parseCsv(teksPenuh);
  // Tanpa range, gviz menggabungkan 5 baris header jadi 1 baris → data mulai index 1.
  const dataRows = csv.slice(1);
  if (!baris4.length || !dataRows.length) throw new Error('Google Sheet BOQ DRM ALL kosong.');

  const lokasi = susunLokasi(baris4, baris5);
  isiItem(lokasi, dataRows, (row, i) => row[i]);
  // Sheet bisa saja sedang dikosongkan/diedit. Kalau baris datanya cuma
  // sisa header/total (atau tidak ada volume yang terbaca), jangan simpan
  // cache kosong — langsung pakai fallback file lokal.
  const lokasiBeriisi = lokasi.filter(l => l.items.length).length;
  if (dataRows.length < 20 || lokasiBeriisi < 10) {
    throw new Error('Google Sheet BOQ DRM ALL belum memuat data volume (' +
      dataRows.length + ' baris, ' + lokasiBeriisi + ' lokasi terisi).');
  }
  return {
    sheet: NAMA_SHEET,
    sumber: 'google',
    file: GOOGLE_URL,
    jumlahLokasi: lokasi.length,
    lokasi,
    _kunci: kunci
  };
}

// ── SUMBER 2: file lokal (fallback) ───────────────────────────────────────
function kunciFile() {
  if (!fs.existsSync(PATH_MASTER)) throw new Error('File master tidak ditemukan: ' + PATH_MASTER);
  const st = fs.statSync(PATH_MASTER);
  return 'file|' + st.mtimeMs + '|' + st.size;
}

function cariFileSheet(zip, namaSheet) {
  const wb = zip.file('xl/workbook.xml');
  if (!wb) return null;
  const xml = wb.asText();
  const reSheet = /<sheet\b([^>]*)\/?>/g;
  let m, rid = null;
  while ((m = reSheet.exec(xml))) {
    const a = m[1] || '';
    const n = /\bname="([^"]+)"/.exec(a);
    if (n && lepasEntitas(n[1]).trim() === namaSheet) {
      const r = /\br:id="([^"]+)"/.exec(a);
      rid = r && r[1];
      break;
    }
  }
  if (!rid) return null;
  const rels = zip.file('xl/_rels/workbook.xml.rels');
  if (!rels) return 'xl/worksheets/sheet2.xml';
  const rx = new RegExp('<Relationship\\b[^>]*\\bId="' + rid + '"[^>]*>', 'i');
  const mm = rx.exec(rels.asText());
  const t = mm && /\bTarget="([^"]+)"/.exec(mm[0]);
  if (!t) return null;
  let target = lepasEntitas(t[1]).replace(/^\//, '');
  if (!/^xl\//i.test(target)) target = 'xl/' + target;
  return zip.file(target) ? target : null;
}

function bacaBaris(sheetXml, shared) {
  const baris = [];
  const reRow = /<row\b([^>]*)>([\s\S]*?)<\/row>/g;
  let m;
  while ((m = reRow.exec(sheetXml))) {
    const r = /\br="(\d+)"/.exec(m[1] || '');
    if (!r) continue;
    const cells = {};
    const reC = /<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g;
    let c;
    while ((c = reC.exec(m[2]))) {
      const ref = /\br="([A-Z]+\d+)"/.exec(c[1] || '');
      if (!ref) continue;
      cells[/^[A-Z]+/.exec(ref[1])[0]] = lepasTeks((c[1] || '') + ' ' + (c[2] || ''), shared);
    }
    baris.push({ nomor: Number(r[1]), cells });
  }
  return baris;
}

function muatMasterFile(kunci) {
  const zip = new PizZip(fs.readFileSync(PATH_MASTER));
  const shared = bacaShared(zip);
  const target = cariFileSheet(zip, NAMA_SHEET);
  if (!target) throw new Error('Sheet "' + NAMA_SHEET + '" tidak ditemukan di file master.');
  const rows = bacaBaris(zip.file(target).asText(), shared);
  const r4 = rows.find(r => r.nomor === 4);
  const r5 = rows.find(r => r.nomor === 5);
  if (!r4 || !r5) throw new Error('Baris 4/5 tidak ditemukan di file master.');

  // Ubah peta {K:'…', O:'…'} jadi array 0-based supaya sama dengan CSV.
  const keArray = cells => Array.from({ length: 1200 }, (_, i) => cells[keKolom(i + 1)] || '');
  const lokasi = susunLokasi(keArray(r4.cells), keArray(r5.cells));
  const dataRows = rows.filter(r => r.nomor >= 6);
  isiItem(lokasi, dataRows, (row, i) => row.cells[keKolom(i + 1)]);
  return {
    sheet: NAMA_SHEET,
    sumber: 'file',
    file: PATH_MASTER,
    jumlahLokasi: lokasi.length,
    lokasi,
    _kunci: kunci
  };
}

// ── dispatcher + cache ────────────────────────────────────────────────────
async function muatMaster() {
  const kunci = SUMBER === 'file'
    ? kunciFile()
    : 'google|' + GOOGLE_ID + '|' + GOOGLE_GID + '|' + Math.floor(Date.now() / TTL_MS);
  if (_cache && _cacheKunci === kunci) return _cache;
  // Hindari beberapa request bersamaan mengunduh sheet yang sama.
  if (_loading && _loadingKunci === kunci) return _loading;
  _loadingKunci = kunci;

  _loading = (async () => {
    try {
      return SUMBER === 'file' ? muatMasterFile(kunci) : await muatMasterGoogle(kunci);
    } catch (e) {
      if (SUMBER !== 'file') {
        try { return muatMasterFile(kunciFile()); }
        catch (e2) { throw new Error(e.message + ' Fallback file gagal: ' + e2.message); }
      }
      throw e;
    }
  })();

  try {
    const master = await _loading;
    _cache = master;
    _cacheKunci = kunci;
    return master;
  } finally {
    _loading = null;
  }
}

function itemsUntuk(lok, tipe) {
  const kunci = tipe === 'rekon' ? 'rekon' : 'drm';
  return lok.items
    .filter(it => it[kunci] > 0)
    .map(it => ({
      designator: it.designator,
      jasa: it.jasa,
      material: it.material,
      volume: it[kunci],
      satuan: it.satuan
    }))
    .sort((a, b) => a.designator.localeCompare(b.designator));
}

// ── API publik ───────────────────────────────────────────────────────────
async function cariLokasiDiMasterDrm(namaLokasi) {
  const master = await muatMaster();
  const q = normNama(namaLokasi);
  const semuaNama = master.lokasi.map(l => l.nama);
  const dasar = { sheet: master.sheet, sumber: master.sumber, jumlahMaster: master.jumlahLokasi };

  if (!q) {
    return Object.assign({ ditemukan: false, saran: semuaNama.slice(0, 10) }, dasar);
  }

  let ketemu = master.lokasi.find(l => l.namaNorm === q);
  if (!ketemu && q.length >= 8) {
    ketemu = master.lokasi.find(l => l.namaNorm.length >= 8 &&
      (l.namaNorm.includes(q) || q.includes(l.namaNorm)));
  }
  if (!ketemu) {
    const saran = master.lokasi
      .filter(l => l.namaNorm.includes(q.slice(0, Math.min(q.length, 12))) || q.includes(l.namaNorm.slice(0, 8)))
      .slice(0, 8).map(l => l.nama);
    return Object.assign({
      ditemukan: false,
      saran: saran.length ? saran : semuaNama.slice(0, 8)
    }, dasar);
  }

  const drm = itemsUntuk(ketemu, 'drm');
  const rekon = itemsUntuk(ketemu, 'rekon');
  return Object.assign({
    ditemukan: true,
    nama_lokasi: String(namaLokasi),
    nama_master: ketemu.nama,
    kolom: ketemu.kolom,
    blok: { drm: ketemu.drm, rekon: ketemu.rekon, tambah: ketemu.tambah, kurang: ketemu.kurang },
    sub_header: ketemu.sub,
    drm: { items: drm, jumlah: drm.length, adaData: drm.length > 0 },
    rekon: { items: rekon, jumlah: rekon.length, adaData: rekon.length > 0 }
  }, dasar);
}

async function daftarNamaMasterDrm() {
  const master = await muatMaster();
  return master.lokasi.map(l => ({ nama: l.nama, kolom: l.kolom }));
}

module.exports = {
  cariLokasiDiMasterDrm, daftarNamaMasterDrm,
  PATH_MASTER, NAMA_SHEET, GOOGLE_URL
};
