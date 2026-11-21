/* =====================================================================
   boqRekonUpload.js — Baca file Excel REKON yang di-upload admin
   ---------------------------------------------------------------------
   Yang diambil HANYA:
     • designator material (awalan M-)  → kolom "Designator Material"
     • designator jasa    (awalan J-)  → kolom "Designator Jasa"
     • nilai REKON                     → kolom ber-header "REKON"
   Baris dengan nilai 0 TIDAK diambil — tidak ditampilkan di textarea
   maupun di tabel rekonsiliasi.
   Kolom harga (G/H) TIDAK dibaca — harga tetap memakai BOQ TAMKUR.
   Contoh struktur: BOQ REKON EBONI  METLAND.xlsx (sheet "Paket5",
   header baris 11, data mulai baris 12, REKON di kolom K).

   Format DRM alternatif (mis. RELOKASI BATU TULIS -DRM.xlsx):
   sheet "RAB FFD FDM", header baris 6 — DESIGNATOR di kolom B (satu
   kolom untuk M-/J-), SATUAN di D, VOL di G. Parser mencari berdasarkan
   teks header, bukan posisi kolom tetap, jadi template lain dengan
   penempatan berbeda tetap terbaca.
   ===================================================================== */
'use strict';

const PizZip = require('pizzip');
const { bacaShared, lepasTeks } = require('./boqTamkur');

function angka(v) {
  if (typeof v === 'number') return v;
  let s = String(v == null ? '' : v).trim();
  if (!s) return 0;
  s = s.replace(/[^\d,.-]/g, '');
  if (/^\d{1,3}(\.\d{3})+(,\d+)?$/.test(s)) {
    s = s.replace(/\./g, '').replace(',', '.');
  } else if (s.indexOf(',') >= 0 && s.indexOf('.') >= 0) {
    s = s.replace(/\./g, '').replace(',', '.');
  } else if (/^\d+,\d+$/.test(s)) {
    s = s.replace(',', '.');
  }
  const n = parseFloat(s);
  return isNaN(n) ? 0 : n;
}

function teks(v) { return String(v == null ? '' : v).trim(); }

function bacaBaris(sheetXml, shared) {
  const baris = [];
  const reRow = /<row\b([^>]*)>([\s\S]*?)<\/row>/g;
  let m;
  while ((m = reRow.exec(sheetXml))) {
    const r = /\br="(\d+)"/.exec(m[1] || '');
    if (!r) continue;
    const cells = {};
    // <c r="H8" s="52"/> adalah sel kosong. Regex lama menggabungkannya
    // dengan </c> sel berikutnya sehingga nilainya bergeser satu kolom.
    const reC = /<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g;
    let c;
    while ((c = reC.exec(m[2]))) {
      const ref = /\br="([A-Z]+\d+)"/.exec(c[1] || '');
      if (!ref) continue;
      const kol = /^[A-Z]+/.exec(ref[1])[0];
      cells[kol] = lepasTeks((c[1] || '') + ' ' + (c[2] || ''), shared);
    }
    baris.push({ nomor: Number(r[1]), cells });
  }
  return baris;
}

function temukanHeader(baris, mode) {
  mode = mode || 'rekon';
  let material = null, jasa = null, nilai = null, satuan = null;
  let barisDes = 0, barisNilai = 0;

  // REKON → kolom ber-header "REKON" (mis. K11 pada BOQ REKON EBONI).
  // DRM   → kolom ber-header "TOTAL VOLUME" (mis. I8 pada BOQ RELOKASI)
  //          atau "VOL" (mis. G6 pada RELOKASI BATU TULIS -DRM.xlsx).
  // Header designator dan header nilai boleh berada di baris berbeda.
  const cariNilai = (s) => {
    if (mode !== 'drm') return s === 'REKON' || s.indexOf('VOLUME REKON') >= 0;
    return s === 'TOTAL VOLUME' || s.indexOf('TOTAL VOLUME') >= 0 ||
           s === 'TOTAL VOL' || s === 'VOL' || s === 'VOL.' || s === 'VOLUME';
  };

  for (const b of baris) {
    for (const [kol, v] of Object.entries(b.cells)) {
      const s = teks(v).toUpperCase();
      if (!s) continue;
      if (!material && s.indexOf('DESIGNATOR MATERIAL') >= 0) { material = kol; barisDes = b.nomor; }
      if (!jasa     && s.indexOf('DESIGNATOR JASA') >= 0)     { jasa = kol;     barisDes = b.nomor; }
      if (!satuan   && s === 'SATUAN')                        { satuan = kol; }
      if (!nilai    && cariNilai(s))                          { nilai = kol;    barisNilai = b.nomor; }
    }
  }
  if (material && jasa && nilai) {
    return {
      nomor: Math.max(barisDes, barisNilai),
      material, jasa, rekon: nilai,
      satuan: satuan || 'F'
    };
  }

  // Format alternatif: satu kolom "DESIGNATOR" (mis. B) berisi M-/J-,
  // total volume di kolom "VOL"/"TOTAL VOLUME" (mis. G). Contoh:
  // RELOKASI BATU TULIS -DRM.xlsx. Fallback ini HANYA dipakai bila header
  // DESIGNATOR MATERIAL + DESIGNATOR JASA tidak ditemukan, jadi format
  // lama (Metland) tetap menjadi prioritas.
  let desTunggal = null;
  for (const b of baris) {
    for (const [kol, v] of Object.entries(b.cells)) {
      if (teks(v).toUpperCase() === 'DESIGNATOR') {
        desTunggal = kol; barisDes = b.nomor; break;
      }
    }
    if (desTunggal) break;
  }
  if (!desTunggal) {
    for (const b of baris) {
      for (const [kol, v] of Object.entries(b.cells)) {
        if (teks(v).toUpperCase().indexOf('DESIGNATOR') >= 0) {
          desTunggal = kol; barisDes = b.nomor; break;
        }
      }
      if (desTunggal) break;
    }
  }
  if (desTunggal && nilai) {
    return {
      nomor: Math.max(barisDes, barisNilai),
      material: desTunggal, jasa: desTunggal, rekon: nilai,
      satuan: satuan || 'F'
    };
  }
  return null;
}

// Daftar worksheet urut sesuai workbook.xml (bukan urutan file di zip).
function daftarSheet(zip) {
  const hasil = [];
  if (!zip.file('xl/workbook.xml')) return hasil;
  const wb = zip.file('xl/workbook.xml').asText();
  const relasi = {};
  if (zip.file('xl/_rels/workbook.xml.rels')) {
    const rx = /<Relationship\b([^>]*)\/>/g;
    let m;
    while ((m = rx.exec(zip.file('xl/_rels/workbook.xml.rels').asText()))) {
      const id = /\bId="([^"]+)"/.exec(m[1]);
      const target = /\bTarget="([^"]+)"/.exec(m[1]);
      if (id && target && /worksheet/i.test(m[1])) relasi[id[1]] = target[1];
    }
  }
  const sx = /<sheet\b([^>]*)\/>/g;
  let s;
  while ((s = sx.exec(wb))) {
    const nama = /\bname="([^"]+)"/.exec(s[1]);
    const rid = /\br:id="([^"]+)"/.exec(s[1]);
    if (!nama || !rid || !relasi[rid[1]]) continue;
    let target = relasi[rid[1]].replace(/^\//, '');
    if (!/^xl\//.test(target)) target = 'xl/' + target;
    if (zip.file(target)) hasil.push({ nama: nama[1], path: target });
  }
  return hasil;
}

/**
 * bacaRekonDariExcel(buffer)
 * Return: { sheet, headerBaris, kolom, items, laporan }
 * items: [{ designator, satuan, rekon, jenis }]
 */
function bacaRekonDariExcel(buffer, mode) {
  mode = mode || 'rekon';
  const zip = new PizZip(buffer);
  const shared = bacaShared(zip);
  let sheets = daftarSheet(zip);
  if (!sheets.length) {
    const fallback = Object.keys(zip.files).find(n => /^xl\/worksheets\/sheet\d+\.xml$/.test(n));
    if (!fallback) throw new Error('Worksheet tidak ditemukan di file Excel.');
    sheets = [{ nama: 'Sheet1', path: fallback }];
  }

  // Pilih sheet yang benar-benar memuat header DESIGNATOR + REKON.
  let dipilih = null;
  for (const s of sheets) {
    const b = bacaBaris(zip.file(s.path).asText(), shared);
    const h = temukanHeader(b, mode);
    if (h) { dipilih = { ...s, baris: b, header: h }; break; }
  }
  if (!dipilih) {
    const s = sheets[0];
    const b = bacaBaris(zip.file(s.path).asText(), shared);
    return {
      sheet: s.nama, headerBaris: null, kolom: null, items: [],
      laporan: { totalBaris: b.length, headerBaris: null, material: 0, jasa: 0, tanpaRekon: 0, dilewati: 0 }
    };
  }

  const { nama: namaSheet, baris, header } = dipilih;
  const laporan = {
    totalBaris: baris.length, headerBaris: header.nomor,
    material: 0, jasa: 0, tanpaRekon: 0, dilewati: 0
  };

  const items = [];
  const seen = new Set();
  baris.forEach(b => {
    if (b.nomor <= header.nomor) return;
    const c = b.cells;
    const rekon = angka(c[header.rekon]);
    const satuan = teks(c[header.satuan]);

    // Baca kedua kolom designator. Jenis ditentukan dari awalannya sendiri:
    // M- = material, J- = jasa. (Pada file nyata, sebagian designator J-
    // bisa berada di kolom C, jadi tidak boleh bergantung pada posisi kolom.)
    [teks(c[header.material]), teks(c[header.jasa])].forEach(mentah => {
      if (!mentah) return;
      const d = mentah.toUpperCase();
      if (!/^[MJ]-/.test(d) || !/^[A-Z0-9][A-Z0-9._/-]*-[A-Z0-9._/-]+$/.test(d)) {
        laporan.dilewati++;
        return;
      }
      const jenis = d.indexOf('M-') === 0 ? 'material' : 'jasa';

      // Nilai 0 tidak diambil sama sekali — jangan memenuhi textarea dan
      // tabel rekonsiliasi dengan baris yang tidak berubah.
      // (seen diisi setelahnya supaya baris 0 tidak menghalangi baris
      //  designator yang sama jika kemudian ada nilainya.)
      if (rekon === 0) {
        laporan.tanpaRekon++;
        return;
      }

      if (seen.has(d)) return;
      seen.add(d);
      items.push({ designator: d, satuan, rekon, jenis });
      laporan[jenis]++;
    });
  });

  return { sheet: namaSheet, headerBaris: header.nomor, kolom: header, items, laporan };
}

// DRM memakai parser yang sama, hanya kolom nilainya "TOTAL VOLUME".
function bacaDrmDariExcel(buffer) { return bacaRekonDariExcel(buffer, 'drm'); }

module.exports = { bacaRekonDariExcel, bacaDrmDariExcel, angka };
