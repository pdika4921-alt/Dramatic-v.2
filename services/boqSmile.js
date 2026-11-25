/* =====================================================================
   boqSmile.js — Generator "NEW BOW SMILE" dari master BOQ SMILE
   ---------------------------------------------------------------------
   Fitur BARU, berdiri sendiri. Tidak memanggil/ubah alur BOQ TAMKUR,
   BA Tambah Kurang, DRM/REKON, atau endpoint lama.

   Master : templates/NEW BOW SMILE.xlsx  (sheet TPT47, dibaca saja)
     A = no_doc / grup paket (dipakai sebagai pilihan)
     B = kd_harga
     C = part_number / designator (kunci VLOOKUP)
     D = deskripsi
     E = satuan
     F = {{lokasi}} → diisi volume REKON dari Excel yang di-upload admin

   Kunci VLOOKUP memakai designator LENGKAP (M-... / J-... tetap dipakai),
   huruf besar dan tanda baca dibuang — sama seperti kunciCocok() di RAB.
   Pengecualian: J-PRELIMINARY-PROJECT diarahkan ke baris khusus
     Preliminary Project | S20240650134 | Prelim-000000001 | ... | Unit
   ===================================================================== */
'use strict';

const fs = require('fs');
const path = require('path');
const PizZip = require('pizzip');

const PATH_XLSX = path.join(__dirname, '..', 'templates', 'NEW BOW SMILE.xlsx');
const KUNCI_PRELIM = 'JPRELIMINARYPROJECT';
const KD_PRELIM = 'S20240650134';
const PART_PRELIM = 'PRELIM000000001';

function kunciSmile(s) {
  return String(s == null ? '' : s).toUpperCase().replace(/[^A-Z0-9]/g, '');
}

function xmlTeks(s) {
  return String(s == null ? '' : s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function bacaShared(zip) {
  if (!zip.file('xl/sharedStrings.xml')) return [];
  return (zip.file('xl/sharedStrings.xml').asText().match(/<si\b[^>]*>[\s\S]*?<\/si>/g) || [])
    .map(si => (si.match(/<t[^>]*>([\s\S]*?)<\/t>/g) || [])
      .map(x => x.replace(/<t[^>]*>|<\/t>/g, ''))
      .join(''));
}

function lepasTeks(inner, shared) {
  const t = /\bt="([^"]+)"/.exec(inner);
  const tipe = t ? t[1] : 'n';
  if (tipe === 's') {
    const v = /<v>([\s\S]*?)<\/v>/.exec(inner);
    return v ? (shared[Number(v[1])] || '') : '';
  }
  if (tipe === 'inlineStr') {
    return (inner.match(/<t[^>]*>([\s\S]*?)<\/t>/g) || [])
      .map(x => x.replace(/<t[^>]*>|<\/t>/g, '')).join('');
  }
  const v = /<v>([\s\S]*?)<\/v>/.exec(inner);
  return v ? v[1] : '';
}

// Urut sesuai workbook.xml, bukan urutan file di zip.
function daftarSheet(zip) {
  const hasil = [];
  if (!zip.file('xl/workbook.xml')) return hasil;
  const relasi = {};
  const rx = /<Relationship\b([^>]*)\/>/g;
  let m;
  if (zip.file('xl/_rels/workbook.xml.rels')) {
    while ((m = rx.exec(zip.file('xl/_rels/workbook.xml.rels').asText()))) {
      const id = /\bId="([^"]+)"/.exec(m[1]);
      const target = /\bTarget="([^"]+)"/.exec(m[1]);
      if (id && target && /worksheet/i.test(m[1])) relasi[id[1]] = target[1];
    }
  }
  const sx = /<sheet\b([^>]*)\/>/g;
  while ((m = sx.exec(zip.file('xl/workbook.xml').asText()))) {
    const nama = /\bname="([^"]+)"/.exec(m[1]);
    const rid = /\br:id="([^"]+)"/.exec(m[1]);
    if (!nama || !rid || !relasi[rid[1]]) continue;
    let target = relasi[rid[1]].replace(/^\//, '');
    if (!/^xl\//.test(target)) target = 'xl/' + target;
    if (zip.file(target)) hasil.push({ nama: nama[1], path: target });
  }
  return hasil;
}

function bacaBaris(sheetXml, shared) {
  const baris = [];
  const reRow = /<row\b([^>]*)>([\s\S]*?)<\/row>/g;
  let m;
  while ((m = reRow.exec(sheetXml))) {
    const r = /\br="(\d+)"/.exec(m[1] || '');
    if (!r) continue;
    const cells = {};
    // Sel kosong <c r="F5" s="2"/> tidak boleh menyita isi sel berikutnya.
    const reC = /<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g;
    let c;
    while ((c = reC.exec(m[2]))) {
      const ref = /\br="([A-Z]+\d+)"/.exec(c[1] || '');
      if (!ref) continue;
      cells[/^[A-Z]+/.exec(ref[1])[0]] = lepasTeks((c[1] || '') + ' ' + (c[2] || ''), shared);
    }
    baris.push({ nomor: Number(r[1]), xml: m[0], cells });
  }
  return baris;
}

function muatMaster() {
  if (!fs.existsSync(PATH_XLSX)) {
    throw new Error('templates/NEW BOW SMILE.xlsx tidak ditemukan.');
  }
  const zip = new PizZip(fs.readFileSync(PATH_XLSX));
  const sheet = daftarSheet(zip)[0];
  const pathSheet = sheet ? sheet.path : 'xl/worksheets/sheet1.xml';
  if (!zip.file(pathSheet)) throw new Error('Worksheet master BOQ SMILE tidak ditemukan.');
  const shared = bacaShared(zip);
  const xml = zip.file(pathSheet).asText();
  return { zip, pathSheet, shared, xml, baris: bacaBaris(xml, shared), sheet: sheet ? sheet.nama : '-' };
}

/** Daftar pilihan paket dari kolom A (tanpa header & tanpa baris Preliminary). */
function daftarPaket() {
  const { baris } = muatMaster();
  const out = [];
  baris.forEach(b => {
    if (b.nomor <= 1) return;
    const a = String(b.cells.A || '').trim();
    if (!a || kunciSmile(a) === 'PRELIMINARYPROJECT' || a === 'Preliminary Project') return;
    if (out.indexOf(a) < 0) out.push(a);
  });
  return out;
}

function angka(v) {
  const n = Number(v);
  return Number.isFinite(n) ? n : 0;
}

function formatAngka(n) {
  const v = Math.round(angka(n) * 1e6) / 1e6;
  return String(v);
}

// F1 master memakai placeholder huruf kecil "{{lokasi}}" (bukan {{LOKASI}}).
function isiHeaderLokasi(zip, lokasi) {
  const ss = zip.file('xl/sharedStrings.xml');
  const nama = String(lokasi == null ? '' : lokasi).trim();
  if (!ss || !nama) return false;
  const sebelum = ss.asText();
  const sesudah = sebelum.replace(
    new RegExp('<si>\\s*<t([^>]*)>\\s*\\{\\{\\s*lokasi\\s*\\}\\}\\s*</t>\\s*</si>', 'gi'),
    (utuh, attr) => '<si><t' + attr + '>' + xmlTeks(nama) + '</t></si>'
  );
  if (sesudah !== sebelum) zip.file('xl/sharedStrings.xml', sesudah);
  return sesudah !== sebelum;
}

function tulisRow(xmlRow, nomorBaru, nilaiF) {
  let out = String(xmlRow || '');
  // Nomor row.
  out = out.replace(/\br="(\d+)"/, 'r="' + nomorBaru + '"');
  // Referensi sel A1..F1 mengikuti nomor row baru.
  out = out.replace(/\br="([A-Z]+)\d+"/g, (m, kol) => 'r="' + kol + nomorBaru + '"');
  // Ganti sel F (kolom hasil VLOOKUP) tanpa mengubah A-E.
  const reF = /<c\b[^>]*\br="F\d+"[^>]*(?:\/>|>[\s\S]*?<\/c>)/g;
  const lama = reF.exec(out);
  const gaya = lama ? (/\bs="([^"]+)"/.exec(lama[0]) || [])[1] : null;
  out = out.replace(reF, '');
  const selF = '<c r="F' + nomorBaru + '"' + (gaya ? ' s="' + gaya + '"' : '') +
    '><v>' + formatAngka(nilaiF) + '</v></c>';
  return out.replace(/<\/row>\s*$/, selF + '</row>');
}

/**
 * bangunBOQSmile({ paket, items, lokasi })
 * items: [{ designator, volume }] dari Excel upload (volume REKON).
 * Return: { ok, buffer, sheet, paket, totalBaris, cocok, baris, preliminary,
 *           identitas, alasan }
 */
function bangunBOQSmile(opsi) {
  opsi = opsi || {};
  const paket = String(opsi.paket || '').trim();
  const hasil = {
    ok: false, buffer: null, sheet: '-', paket, totalBaris: 0, cocok: 0,
    baris: [], preliminary: null, identitas: { lokasi: false }, alasan: ''
  };
  try {
    const master = muatMaster();
    hasil.sheet = master.sheet;
    const tersedia = daftarPaket();
    if (!paket || tersedia.indexOf(paket) < 0) {
      hasil.alasan = 'Paket BOQ SMILE tidak valid.';
      return hasil;
    }

    const peta = {};
    let volPrelim = 0;
    (Array.isArray(opsi.items) ? opsi.items : []).forEach(it => {
      const k = kunciSmile(it && it.designator);
      const v = angka(it && it.volume);
      if (!k || !v) return;
      if (k === KUNCI_PRELIM || k === 'PRELIMINARYPROJECT') volPrelim += v;
      else peta[k] = (peta[k] || 0) + v;
    });

    const barisPaket = master.baris.filter(b =>
      b.nomor > 1 && String(b.cells.A || '').trim() === paket
    );
    if (!barisPaket.length) {
      hasil.alasan = 'Baris untuk paket ini tidak ditemukan di master.';
      return hasil;
    }

    // Baris khusus Preliminary Project (kode tetap, bukan milik paket mana pun).
    const barisPrelim = master.baris.find(b =>
      kunciSmile(b.cells.B) === KD_PRELIM && kunciSmile(b.cells.C) === PART_PRELIM
    );

    // J-Preliminary-Project milik paket sengaja tidak dipakai: nilainya masuk
    // ke baris khusus di bawah, supaya hanya ada satu baris Preliminary.
    const dipilih = barisPaket.filter(b => kunciSmile(b.cells.C) !== KUNCI_PRELIM);
    if (barisPrelim) dipilih.push(barisPrelim);

    const header = master.baris.find(b => b.nomor === 1);
    const baru = [];
    let nomor = 1;
    let cocok = 0;
    // Baris judul lokasi (F1 = {{lokasi}}) dipertahankan apa adanya:
    // nilainya diganti lewat sharedStrings oleh isiHeaderLokasi() di bawah.
    if (header) baru.push(header.xml);

    dipilih.forEach(b => {
      nomor++;
      const khas = b === barisPrelim;
      const nilai = khas ? volPrelim : (peta[kunciSmile(b.cells.C)] || 0);
      if (nilai) cocok++;
      baru.push(tulisRow(b.xml, nomor, nilai));
      hasil.baris.push({
        no: nomor,
        a: b.cells.A || '', b: b.cells.B || '', c: b.cells.C || '',
        d: b.cells.D || '', e: b.cells.E || '', f: nilai
      });
      if (khas) hasil.preliminary = hasil.baris[hasil.baris.length - 1];
    });

    const sheetBaru = master.xml.replace(
      /<sheetData>([\s\S]*?)<\/sheetData>/,
      '<sheetData>' + baru.join('') + '</sheetData>'
    ).replace(/<dimension\b[^>]*\/>/, '<dimension ref="A1:F' + nomor + '"/>')
      .replace(/<autoFilter\b[^>]*\/>/, '<autoFilter ref="A1:F' + nomor + '"/>');

    master.zip.file(master.pathSheet, sheetBaru);
    hasil.identitas.lokasi = isiHeaderLokasi(master.zip, opsi.lokasi);
    hasil.buffer = master.zip.generate({ type: 'nodebuffer', compression: 'DEFLATE' });
    hasil.ok = true;
    hasil.totalBaris = nomor;
    hasil.cocok = cocok;
  } catch (e) {
    hasil.alasan = e.message || String(e);
  }
  return hasil;
}

module.exports = { daftarPaket, bangunBOQSmile, kunciSmile, PATH_XLSX };
