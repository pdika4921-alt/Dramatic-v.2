/* =====================================================================
   boqTamkur.js —Sinkronisasi BA Tambah Kurang → BOQ TAMKUR.xlsx
   ---------------------------------------------------------------------
   • Membaca nilai DRM/REKON yang sudah dihitung generator dari database,
     lalu menuliskannya ke kolom K (DRM) dan L (REKON) pada baris yang
     designatornya cocok.
   • Pencocokan memakai kolom C (designator material, awalan M-) dan
     kolom D (designator jasa, awalan J-), dengan fallback ke kolom B.
   • Nama sheet diseragamkan menjadi "BOQ TAMKUR".
   • bacaHargaBoqTamkur('amd4'|'amd5') = VLOOKUP harga jasa/material dari
     BOQ TAMKUR (AMD4) atau BOQ TAMKUR AMD 5 PAKET 5. File AMD5 dibaca saja.
   • Memakai PizZip (sudah ada di package.json) — tanpa dependency baru.
   • Kegagalan menulis TIDAK boleh menggagalkan generate BA.
   ===================================================================== */
'use strict';

const fs = require('fs');
const path = require('path');
const PizZip = require('pizzip');

const PATH_XLSX = path.join(__dirname, '..', 'templates', 'BOQ TAMKUR.xlsx');
const NAMA_SHEET = 'BOQ TAMKUR';
const PATH_LOCK = path.join(__dirname, '..', 'templates', '~$BOQ TAMKUR.xlsx');

// Master harga alternatif (AMD5). Dipakai untuk baca harga/VLOOKUP dan,
// bila pemilih VLOOKUP AMD5 aktif, sebagai template generate/download BOQ.
const PATH_XLSX_AMD5 = path.join(__dirname, '..', 'templates', 'BOQ TAMKUR AMD 5 PAKET 5.xlsx');
const NAMA_SHEET_AMD5 = 'AMD 5 Paket5';
const PATH_LOCK_AMD5 = path.join(__dirname, '..', 'templates', '~$BOQ TAMKUR AMD 5 PAKET 5.xlsx');

// Pilih template sesuai pemilih VLOOKUP di toolbar.
// Default 'amd4' = alur lama (BOQ TAMKUR.xlsx), 'amd5' = BOQ AMD 5 PAKET 5.
function pilihTemplate(opsi) {
  const amd5 = String((opsi && opsi.sumberMaster) || '').toLowerCase() === 'amd5';
  return amd5
    ? { amd5: true,  file: PATH_XLSX_AMD5, lock: PATH_LOCK_AMD5, sheet: NAMA_SHEET_AMD5 }
    : { amd5: false, file: PATH_XLSX,      lock: PATH_LOCK,      sheet: NAMA_SHEET };
}

function norm(s) {
  return String(s == null ? '' : s).trim().toUpperCase().replace(/\s+/g, '');
}

function lepasTeks(inner, shared) {
  const mT = /\bt="([^"]+)"/.exec(inner) || [];
  const tipe = mT[1] || 'n';
  if (tipe === 's') {
    const mV = /<v>([\s\S]*?)<\/v>/.exec(inner);
    return mV ? (shared[Number(mV[1])] || '') : '';
  }
  if (tipe === 'inlineStr') {
    return (inner.match(/<t[^>]*>([\s\S]*?)<\/t>/g) || [])
      .map(x => x.replace(/<t[^>]*>|<\/t>/g, ''))
      .join('');
  }
  const mV = /<v>([\s\S]*?)<\/v>/.exec(inner);
  return mV ? mV[1] : '';
}

function bacaShared(zip) {
  if (!zip.file('xl/sharedStrings.xml')) return [];
  const xml = zip.file('xl/sharedStrings.xml').asText();
  return (xml.match(/<si\b[^>]*>[\s\S]*?<\/si>/g) || []).map(si =>
    (si.match(/<t[^>]*>([\s\S]*?)<\/t>/g) || [])
      .map(x => x.replace(/<t[^>]*>|<\/t>/g, ''))
      .join('')
  );
}

function indeksSheet(sheetXml, shared) {
  const indeks = { C: {}, D: {}, B: {} };
  // Sel kosong bertipe <c r="H8" s="52"/> harus dibaca terpisah,
  // kalau tidak nilainya bergeser ke kolom sebelumnya.
  const reC = /<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g;
  let m;
  while ((m = reC.exec(sheetXml))) {
    const atrib = m[1] || '';
    const r = /\br="([A-Z]+\d+)"/.exec(atrib);
    if (!r) continue;
    const ref = r[1];
    const kol = /^[A-Z]+/.exec(ref)[0];
    if (!indeks[kol]) continue;
    const baris = Number(/\d+/.exec(ref)[0]);
    const nilai = norm(lepasTeks(atrib + ' ' + (m[2] || ''), shared));
    if (nilai) indeks[kol][nilai] = baris;
  }
  return indeks;
}

function cariBaris(indeks, designator) {
  const d = norm(designator);
  if (!d) return null;
  if (/^M-/.test(d) && indeks.C[d]) return indeks.C[d];
  if (/^J-/.test(d) && indeks.D[d]) return indeks.D[d];
  const tanpa = d.replace(/^[MJ]-/, '');
  if (indeks.C[tanpa]) return indeks.C[tanpa];
  if (indeks.D[tanpa]) return indeks.D[tanpa];
  if (indeks.B[tanpa]) return indeks.B[tanpa];
  if (indeks.B[d]) return indeks.B[d];
  return null;
}

function hapusCell(inner, kolom, baris) {
  const re = new RegExp('<c\\b[^>]*\\br="' + kolom + baris + '"[^>]*(?:\\/>|>[\\s\\S]*?<\\/c>)', 'g');
  return inner.replace(re, '');
}

function tulisPasangan(sheetXml, baris, drm, rekon) {
  const reRow = /<row\b([^>]*)>([\s\S]*?)<\/row>/g;
  let ketemu = false;
  const out = sheetXml.replace(reRow, (utuh, atrib, inner) => {
    const r = /\br="(\d+)"/.exec(atrib || '');
    if (!r || Number(r[1]) !== baris) return utuh;
    ketemu = true;
    let baru = hapusCell(hapusCell(inner, 'K', baris), 'L', baris);
    const sel = '<c r="K' + baris + '"><v>' + Number(drm || 0) + '</v></c>'
      + '<c r="L' + baris + '"><v>' + Number(rekon || 0) + '</v></c>';
    const mSesudah = /(<c\b[^>]*\br="M\d+")/.exec(baru);
    baru = mSesudah ? baru.replace(mSesudah[0], sel + mSesudah[0]) : baru + sel;
    return '<row' + atrib + '>' + baru + '</row>';
  });
  return { xml: out, ketemu };
}

function enamBelas(n) { return String(Math.round(Number(n) * 100) / 100); }

function xmlTeks(s) {
  return String(s == null ? '' : s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

/**
 * isiIdentitasBoq(zip, opsi)
 * Isi placeholder {{AREA}} / {{LOKASI}} di xl/sharedStrings.xml.
 * Hanya shared string yang isinya PERSIS placeholder yang diganti —
 * sel, rumus, designator, dan harga tidak disentuh sama sekali.
 * Tanpa opsi.area / opsi.lokasi → file dibiarkan apa adanya (alur lama).
 */
function isiIdentitasBoq(zip, opsi) {
  const hasil = { area: false, lokasi: false };
  if (!zip.file('xl/sharedStrings.xml')) return hasil;
  opsi = opsi || {};
  const area   = String(opsi.area   == null ? '' : opsi.area).trim();
  const lokasi = String(opsi.lokasi == null ? '' : opsi.lokasi).trim();
  if (!area && !lokasi) return hasil;

  const sebelum = zip.file('xl/sharedStrings.xml').asText();
  const ganti = (teks, placeholder, nilai) => teks.replace(
    new RegExp('<si>\\s*<t([^>]*)>\\s*\\{\\{' + placeholder + '\\}\\}\\s*<\\/t>\\s*<\\/si>', 'g'),
    (utuh, attr) => '<si><t' + attr + '>' + xmlTeks(nilai) + '</t></si>'
  );

  const sesudahArea = area ? ganti(sebelum, 'AREA', area) : sebelum;
  hasil.area = sesudahArea !== sebelum;
  const sesudah = lokasi ? ganti(sesudahArea, 'LOKASI', lokasi) : sesudahArea;
  hasil.lokasi = sesudah !== sesudahArea;

  if (sesudah !== sebelum) zip.file('xl/sharedStrings.xml', sesudah);
  return hasil;
}

/**
 * tulisBoqTamkur(rows, opsi)
 * rows: [{ designator, drm, rekon }, ...]
 * opsi (opsional): { area, lokasi, pathXlsx, pathLock }
 * Return: { ok, ditulis, dilewati, alasan, identitas }
 */
function tulisBoqTamkur(rows, opsi) {
  opsi = opsi || {};
  const tpl = pilihTemplate(opsi);
  const fileXlsx = opsi.pathXlsx || tpl.file;
  const fileLock = opsi.pathLock || tpl.lock;
  const hasil = {
    ok: false, ditulis: 0, dilewati: [], alasan: '',
    identitas: { area: false, lokasi: false },
    sumber_master: tpl.amd5 ? 'amd5' : 'amd4', sheet: tpl.sheet
  };

  if (!fs.existsSync(fileXlsx)) {
    hasil.alasan = 'templates/' + path.basename(tpl.file) + ' tidak ditemukan';
    return hasil;
  }
  if (fs.existsSync(fileLock)) {
    hasil.alasan = 'file sedang dibuka di Excel — tutup dulu agar bisa ditulis';
    return hasil;
  }
  if (!Array.isArray(rows) || !rows.length) {
    hasil.alasan = 'tidak ada baris rekon untuk ditulis';
    return hasil;
  }

  try {
    const zip = new PizZip(fs.readFileSync(fileXlsx));
    const shared = bacaShared(zip);
    const sheetPath = zip.file('xl/worksheets/sheet1.xml')
      ? 'xl/worksheets/sheet1.xml'
      : (Object.keys(zip.files).find(n => /^xl\/worksheets\/sheet\d+\.xml$/.test(n)) || '');
    if (!sheetPath) {
      hasil.alasan = 'worksheet tidak ditemukan di file Excel';
      return hasil;
    }

    let sheetXml = zip.file(sheetPath).asText();
    const indeks = indeksSheet(sheetXml, shared);

    rows.forEach(item => {
      const baris = cariBaris(indeks, item && item.designator);
      if (!baris) {
        hasil.dilewati.push(String((item && item.designator) || '-'));
        return;
      }
      const t = tulisPasangan(sheetXml, baris, enamBelas(item.drm), enamBelas(item.rekon));
      if (t.ketemu) { sheetXml = t.xml; hasil.ditulis++; }
      else hasil.dilewati.push(String(item.designator) + ' (baris ' + baris + ' tidak ditemukan)');
    });

    // Isi identitas lokasi/area pada placeholder template (kalau dikirim).
    hasil.identitas = isiIdentitasBoq(zip, opsi);

    // Seragamkan nama sheet, termasuk referensi di definedNames.
    // AMD5 memakai sheet aslinya ("AMD 5 Paket5") — jangan diganti ke
    // "BOQ TAMKUR" karena itu nama sheet template AMD4.
    if (zip.file('xl/workbook.xml')) {
      let wb = zip.file('xl/workbook.xml').asText();
      if (!tpl.amd5) wb = wb.split('BOQ AMD4 PAKET 5').join(NAMA_SHEET);
      // Paksa Excel menghitung ulang SUMIF/IF setelah kolom K/L berubah.
      if (/<calcPr\b/.test(wb)) {
        wb = wb.replace(/<calcPr\b([^>]*)\/?>/, (utuh, a) => {
          if (/fullCalcOnLoad=/.test(a)) return utuh;
          const bersih = a.replace(/\/\s*$/, '');
          return '<calcPr' + bersih + ' fullCalcOnLoad="1"/>';
        });
      } else {
        wb = wb.replace('</workbook>', '<calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>');
      }
      zip.file('xl/workbook.xml', wb);
    }

    zip.file(sheetPath, sheetXml);
    zip.remove('xl/calcChain.xml'); // biar Excel rebuild sendiri

    const buf = zip.generate({ type: 'nodebuffer', compression: 'DEFLATE' });
    const tmp = fileXlsx + '.tmp';
    fs.writeFileSync(tmp, buf);
    fs.renameSync(tmp, fileXlsx);
    hasil.ok = true;
    return hasil;
  } catch (e) {
    // Jangan pernah menggagalkan generate BA gara-gara sinkronisasi Excel.
    try { const t = fileXlsx + '.tmp'; if (fs.existsSync(t)) fs.unlinkSync(t); } catch (_) {}
    hasil.alasan = e.message;
    return hasil;
  }
}

// ─────────────────────────────────────────────────────────────────────────────
//  BUILDER BARU: BOQ Tambah Kurang → Buffer .xlsx (tanpa menimpa template)
//  Fungsi lama tulisBoqTamkur() sengaja tidak diubah: generator BA masih
//  memakai itu. Fitur download Excel memakai fungsi-fungsi di bawah ini.
// ─────────────────────────────────────────────────────────────────────────────

/** Isi nilai cache (<v>) pada sel rumus, rumusnya tetap dipertahankan. */
function isiCacheRumus(sheetXml, ref, nilai) {
  const re = new RegExp('<c\\b[^>]*\\br="' + ref + '"[^>]*>([\\s\\S]*?)<\\/c>');
  return sheetXml.replace(re, utuh => {
    if (!/<f\b/.test(utuh)) return utuh;
    if (/<v>[\s\S]*?<\/v>/.test(utuh)) {
      return utuh.replace(/<v>[\s\S]*?<\/v>/, '<v>' + enamBelas(nilai) + '</v>');
    }
    return utuh.replace(/<\/c>$/, '<v>' + enamBelas(nilai) + '</v></c>');
  });
}

// Versi aman khusus builder download. Helper lama hapusCell() punya
// kasus tepi: sel mandiri <c r="K6"/> tersalah tangkap sampai </c>
// berikutnya sehingga sel M ikut hilang. Fungsi lama sengaja dibiarkan.
function hapusCellAman(inner, kolom, baris) {
  const ref = kolom + baris;
  const reSelf = new RegExp('<c\\b[^>]*\\br="' + ref + '"[^>]*\\/>', 'g');
  const reFull = new RegExp('<c\\b[^>]*\\br="' + ref + '"[^>]*>[\\s\\S]*?<\\/c>', 'g');
  return inner.replace(reSelf, '').replace(reFull, '');
}

function ambilGaya(inner, kolom, baris) {
  const ref = kolom + baris;
  const re = new RegExp('<c\\b[^>]*\\br="' + ref + '"([^>]*)>');
  const m = re.exec(inner);
  if (!m) return '';
  const s = /\bs="(\d+)"/.exec(m[1]);
  return s ? ' s="' + s[1] + '"' : '';
}

function tulisPasanganAman(sheetXml, baris, drm, rekon) {
  const reRow = /<row\b([^>]*)>([\s\S]*?)<\/row>/g;
  let ketemu = false;
  const out = sheetXml.replace(reRow, (utuh, atrib, inner) => {
    const r = /\br="(\d+)"/.exec(atrib || '');
    if (!r || Number(r[1]) !== baris) return utuh;
    ketemu = true;
    const gayaK = ambilGaya(inner, 'K', baris);
    const gayaL = ambilGaya(inner, 'L', baris);
    let baru = hapusCellAman(hapusCellAman(inner, 'K', baris), 'L', baris);
    const sel = '<c r="K' + baris + '"' + gayaK + '><v>' + Number(drm || 0) + '</v></c>'
      + '<c r="L' + baris + '"' + gayaL + '><v>' + Number(rekon || 0) + '</v></c>';
    const mSesudah = /(<c\b[^>]*\br="M\d+")/.exec(baru);
    baru = mSesudah ? baru.replace(mSesudah[0], sel + mSesudah[0]) : baru + sel;
    return '<row' + atrib + '>' + baru + '</row>';
  });
  return { xml: out, ketemu };
}

/** Segarkan kolom turunan I/J/M/N sesuai K/L yang baru ditulis. */
function segarkanTurunan(sheetXml, baris, drm, rekon) {
  const d = Number(drm) || 0;
  const r = Number(rekon) || 0;
  sheetXml = isiCacheRumus(sheetXml, 'I' + baris, r); // VOLUME REKON
  sheetXml = isiCacheRumus(sheetXml, 'J' + baris, d); // VOLUME DRM
  sheetXml = isiCacheRumus(sheetXml, 'M' + baris, Math.max(r - d, 0)); // TAMBAH
  sheetXml = isiCacheRumus(sheetXml, 'N' + baris, Math.max(d - r, 0)); // KURANG
  return sheetXml;
}

/** Nama sheet + paksa Excel menghitung ulang (dipakai builder download). */
function siapSheetBoq(zip, sheetPath, sheetXml, tpl) {
  if (zip.file('xl/workbook.xml')) {
    let wb = zip.file('xl/workbook.xml').asText();
    // AMD5 mempertahankan sheet aslinya; AMD4 tetap diseragamkan.
    if (!tpl || !tpl.amd5) wb = wb.split('BOQ AMD4 PAKET 5').join(NAMA_SHEET);
    if (/<calcPr\b/.test(wb)) {
      wb = wb.replace(/<calcPr\b([^>]*)\/?>/, (utuh, a) => {
        if (/fullCalcOnLoad=/.test(a)) return utuh;
        return '<calcPr' + a.replace(/\/\s*$/, '') + ' fullCalcOnLoad="1"/>';
      });
    } else {
      wb = wb.replace('</workbook>', '<calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>');
    }
    zip.file('xl/workbook.xml', wb);
  }
  zip.file(sheetPath, sheetXml);
  zip.remove('xl/calcChain.xml');
}

/**
 * bangunBoqTamkur(rows, opsi)
 * Sama seperti tulisBoqTamkur, tetapi template tidak ditimpa: hasil
 * dikembalikan sebagai Buffer .xlsx siap di-download.
 * rows: [{ designator, drm, rekon }, ...]
 * opsi (opsional): { area, lokasi }
 * Return: { ok, ditulis, dilewati, alasan, identitas, buffer }
 */
function bangunBoqTamkur(rows, opsi) {
  opsi = opsi || {};
  const tpl = pilihTemplate(opsi);
  const hasil = {
    ok: false, ditulis: 0, dilewati: [], alasan: '',
    identitas: { area: false, lokasi: false }, buffer: null,
    sumber_master: tpl.amd5 ? 'amd5' : 'amd4', sheet: tpl.sheet
  };

  if (!fs.existsSync(tpl.file)) {
    hasil.alasan = 'templates/' + path.basename(tpl.file) + ' tidak ditemukan';
    return hasil;
  }
  if (!Array.isArray(rows) || !rows.length) {
    hasil.alasan = 'tidak ada baris rekon untuk ditulis';
    return hasil;
  }

  try {
    const zip = new PizZip(fs.readFileSync(tpl.file)); // read-only
    const shared = bacaShared(zip);
    const sheetPath = zip.file('xl/worksheets/sheet1.xml')
      ? 'xl/worksheets/sheet1.xml'
      : (Object.keys(zip.files).find(n => /^xl\/worksheets\/sheet\d+\.xml$/.test(n)) || '');
    if (!sheetPath) {
      hasil.alasan = 'worksheet tidak ditemukan di file Excel';
      return hasil;
    }

    let sheetXml = zip.file(sheetPath).asText();
    const indeks = indeksSheet(sheetXml, shared);
    const terpakai = {};

    rows.forEach(item => {
      const baris = cariBaris(indeks, item && item.designator);
      if (!baris) {
        hasil.dilewati.push(String((item && item.designator) || '-'));
        return;
      }
      const drm  = Number(item.drm)  || 0;
      const rekon = Number(item.rekon) || 0;
      const t = tulisPasanganAman(sheetXml, baris, enamBelas(drm), enamBelas(rekon));
      if (!t.ketemu) {
        hasil.dilewati.push(String(item.designator) + ' (baris ' + baris + ' tidak ditemukan)');
        return;
      }
      sheetXml = t.xml;
      terpakai[baris] = { baris, drm, rekon };
      hasil.ditulis++;
    });

    Object.keys(terpakai).forEach(k => {
      const x = terpakai[k];
      sheetXml = segarkanTurunan(sheetXml, x.baris, x.drm, x.rekon);
    });

    hasil.identitas = isiIdentitasBoq(zip, opsi);
    siapSheetBoq(zip, sheetPath, sheetXml, tpl);
    hasil.buffer = zip.generate({ type: 'nodebuffer', compression: 'DEFLATE' });
    hasil.ok = true;
    return hasil;
  } catch (e) {
    hasil.alasan = e.message;
    return hasil;
  }
}

/**
 * bacaHargaBoqTamkur(sumber)
 * Harga satuan untuk VLOOKUP:
 *   • sumber 'amd4' (DEFAULT, alur lama) → templates/BOQ TAMKUR.xlsx
 *   • sumber 'amd5'                    → templates/BOQ TAMKUR AMD 5 PAKET 5.xlsx
 * Kedua file memakai aturan kolom yang sama:
 *   • designator M- di kolom C → harga material kolom G
 *   • designator J- di kolom D → harga jasa    kolom H
 * Return: { "M-...": 7098, "J-...": 2169, ... }
 */
function bacaHargaBoqTamkur(sumber) {
  const peta = {};
  const pakaiAmd5 = String(sumber == null ? '' : sumber).toLowerCase() === 'amd5';
  const filePath = pakaiAmd5 ? PATH_XLSX_AMD5 : PATH_XLSX;
  if (!fs.existsSync(filePath)) return peta;
  try {
    const zip = new PizZip(fs.readFileSync(filePath));
    const shared = bacaShared(zip);
    const sheetPath = zip.file('xl/worksheets/sheet1.xml')
      ? 'xl/worksheets/sheet1.xml'
      : (Object.keys(zip.files).find(n => /^xl\/worksheets\/sheet\d+\.xml$/.test(n)) || '');
    if (!sheetPath) return peta;
    const xml = zip.file(sheetPath).asText();

    const reRow = /<row\b([^>]*)>([\s\S]*?)<\/row>/g;
    let baris;
    while ((baris = reRow.exec(xml))) {
      const cells = {};
      const reC = /<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g;
      let c;
      while ((c = reC.exec(baris[2]))) {
        const ref = /\br="([A-Z]+\d+)"/.exec(c[1] || '');
        if (!ref) continue;
        cells[/^[A-Z]+/.exec(ref[1])[0]] = lepasTeks((c[1] || '') + ' ' + (c[2] || ''), shared);
      }
      const pasang = (kolDes, kolHar) => {
        const d = norm(cells[kolDes]);
        if (!/^[MJ]-/.test(d)) return;
        const n = Number(String(cells[kolHar] || '').replace(/[^\d.-]/g, ''));
        peta[d] = isNaN(n) ? 0 : n;
      };
      pasang('C', 'G');
      pasang('D', 'H');
    }
  } catch (e) {
    console.warn('[BOQ TAMKUR] gagal membaca harga ' + path.basename(filePath) + ':', e.message);
  }
  return peta;
}

module.exports = {
  tulisBoqTamkur, bacaHargaBoqTamkur, bangunBoqTamkur,
  PATH_XLSX, NAMA_SHEET, PATH_XLSX_AMD5, NAMA_SHEET_AMD5,
  isiIdentitasBoq, bacaShared, lepasTeks, indeksSheet, cariBaris
};
