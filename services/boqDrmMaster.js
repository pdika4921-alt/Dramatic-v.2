/* =====================================================================
   boqDrmMaster.js — BACA SAJA data master BOQ DRM ALL.xlsx
   ---------------------------------------------------------------------
   Dipakai oleh tombol "CEK DATA MASTER BOQ" di panel Input RAB/DRM dan
   Input Aktual/REKON (halaman RAB & Rekonsiliasi).

   Aturan file master (tidak diubah, dibaca saja):
     • sheet   : "BOQ AMD4 PAKET 5"
     • baris 4 : NAMA LOKASI, mulai kolom K, tiap lokasi lebar 4 kolom
                 (K, O, S, W, …)
     • baris 5 : sub-header DRM | REKON | TAMBAH | KURANG
     • baris 6+: C = designator M-, D = designator J-, F = satuan,
                 G = harga material, H = harga jasa,
                 4 kolom pertama tiap blok = volume DRM / REKON.

   File master BERSIFAT READ-ONLY. Tidak pernah ditulis/ disimpan.
   Hasil baca di-cache di memori, hangus otomatis saat file berubah
   (mtime + ukuran) atau saat server di-restart.
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

let _cache = null;
let _cacheKunci = '';

// ── helper ───────────────────────────────────────────────────────────────
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

// Temukan file XML sheet berdasarkan nama (bukan berdasarkan nomor urut).
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
  if (!rels) return 'xl/worksheets/sheet2.xml'; // fallback untuk file lama
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

// ── baca master (cache) ──────────────────────────────────────────────────
function muatMaster() {
  if (!fs.existsSync(PATH_MASTER)) {
    throw new Error('File master tidak ditemukan: ' + PATH_MASTER);
  }
  const st = fs.statSync(PATH_MASTER);
  const kunci = st.mtimeMs + '|' + st.size;
  if (_cache && _cacheKunci === kunci) return _cache;

  const zip = new PizZip(fs.readFileSync(PATH_MASTER));
  const shared = bacaShared(zip);
  const target = cariFileSheet(zip, NAMA_SHEET);
  if (!target) throw new Error('Sheet "' + NAMA_SHEET + '" tidak ditemukan di file master.');

  const rows = bacaBaris(zip.file(target).asText(), shared);
  const r4 = rows.find(r => r.nomor === 4);
  const r5 = rows.find(r => r.nomor === 5);
  if (!r4 || !r5) throw new Error('Baris 4 (nama lokasi) / baris 5 (sub-header) tidak ditemukan.');

  // Kumpulkan blok lokasi: 4 kolom berturut mulai K, O, S, …
  const lokasi = [];
  const awal = keIndeks(KOLOM_AWAL);
  for (let i = 0; i < 300; i++) {
    const basis = awal + i * LEBAR_BLOK;
    const cNama = keKolom(basis);
    const nama = String(r4.cells[cNama] || '').trim();
    if (!nama) {
      // Berhenti setelah dua blok kosong berturut (antisipasi celah kosong).
      const berikut = String(r4.cells[keKolom(basis + LEBAR_BLOK)] || '').trim();
      if (i > 0 && !berikut) break;
      continue;
    }
    const sub = [0, 1, 2, 3].map(k =>
      String(r5.cells[keKolom(basis + k)] || '').trim().toUpperCase());
    lokasi.push({
      nama,
      namaNorm: normNama(nama),
      kolom: cNama,
      drm: keKolom(basis),
      rekon: keKolom(basis + 1),
      tambah: keKolom(basis + 2),
      kurang: keKolom(basis + 3),
      sub,
      items: []
    });
  }
  if (!lokasi.length) throw new Error('Tidak ada nama lokasi di baris 4 sheet ' + NAMA_SHEET + '.');

  // Baris data mulai 6. Designator M- di kolom C, J- di kolom D.
  // Satu baris bisa memuat dua designator → dua item dengan volume sama.
  const dataRows = rows.filter(r => r.nomor >= 6);
  for (const lok of lokasi) {
    const agregat = new Map();
    for (const row of dataRows) {
      const vDrm  = angka(row.cells[lok.drm]);
      const vRek  = angka(row.cells[lok.rekon]);
      if (!vDrm && !vRek) continue;

      const desM = String(row.cells.C || '').trim().toUpperCase();
      const desJ = String(row.cells.D || '').trim().toUpperCase();
      const satuan = String(row.cells.F || '').trim();
      const hMat = angka(row.cells.G);
      const hJas = angka(row.cells.H);
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

  _cache = {
    sheet: NAMA_SHEET,
    file: PATH_MASTER,
    jumlahLokasi: lokasi.length,
    lokasi
  };
  _cacheKunci = kunci;
  return _cache;
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
function cariLokasiDiMasterDrm(namaLokasi) {
  const master = muatMaster();
  const q = normNama(namaLokasi);
  const semuaNama = master.lokasi.map(l => l.nama);

  if (!q) {
    return { ditemukan: false, sheet: master.sheet, jumlahMaster: master.jumlahLokasi, saran: semuaNama.slice(0, 10) };
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
    return {
      ditemukan: false,
      sheet: master.sheet,
      jumlahMaster: master.jumlahLokasi,
      saran: saran.length ? saran : semuaNama.slice(0, 8)
    };
  }

  const drm = itemsUntuk(ketemu, 'drm');
  const rekon = itemsUntuk(ketemu, 'rekon');
  return {
    ditemukan: true,
    nama_lokasi: String(namaLokasi),
    nama_master: ketemu.nama,
    sheet: master.sheet,
    kolom: ketemu.kolom,
    blok: { drm: ketemu.drm, rekon: ketemu.rekon, tambah: ketemu.tambah, kurang: ketemu.kurang },
    sub_header: ketemu.sub,
    drm: { items: drm, jumlah: drm.length, adaData: drm.length > 0 },
    rekon: { items: rekon, jumlah: rekon.length, adaData: rekon.length > 0 },
    jumlahMaster: master.jumlahLokasi
  };
}

function daftarNamaMasterDrm() {
  return muatMaster().lokasi.map(l => ({ nama: l.nama, kolom: l.kolom }));
}

module.exports = { cariLokasiDiMasterDrm, daftarNamaMasterDrm, PATH_MASTER, NAMA_SHEET };
