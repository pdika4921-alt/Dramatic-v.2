/* =====================================================================
   boqTamkur.js —Sinkronisasi BA Tambah Kurang → BOQ TAMKUR.xlsx
   ---------------------------------------------------------------------
   • Membaca nilai DRM/REKON yang sudah dihitung generator dari database,
     lalu menuliskannya ke kolom K (DRM) dan L (REKON) pada baris yang
     designatornya cocok.
   • Pencocokan memakai kolom C (designator material, awalan M-) dan
     kolom D (designator jasa, awalan J-), dengan fallback ke kolom B.
   • Nama sheet diseragamkan menjadi "BOQ TAMKUR".
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

/**
 * tulisBoqTamkur(rows)
 * rows: [{ designator, drm, rekon }, ...]
 * Return: { ok, ditulis, dilewati, alasan }
 */
function tulisBoqTamkur(rows, opsi) {
  opsi = opsi || {};
  const fileXlsx = opsi.pathXlsx || PATH_XLSX;
  const fileLock = opsi.pathLock || PATH_LOCK;
  const hasil = { ok: false, ditulis: 0, dilewati: [], alasan: '' };

  if (!fs.existsSync(fileXlsx)) {
    hasil.alasan = 'templates/BOQ TAMKUR.xlsx tidak ditemukan';
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

    // Seragamkan nama sheet, termasuk referensi di definedNames.
    if (zip.file('xl/workbook.xml')) {
      let wb = zip.file('xl/workbook.xml').asText();
      wb = wb.split('BOQ AMD4 PAKET 5').join(NAMA_SHEET);
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
function siapSheetBoq(zip, sheetPath, sheetXml) {
  if (zip.file('xl/workbook.xml')) {
    let wb = zip.file('xl/workbook.xml').asText();
    wb = wb.split('BOQ AMD4 PAKET 5').join(NAMA_SHEET);
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
 * bangunBoqTamkur(rows)
 * Sama seperti tulisBoqTamkur, tetapi template tidak ditimpa: hasil
 * dikembalikan sebagai Buffer .xlsx siap di-download.
 * rows: [{ designator, drm, rekon }, ...]
 * Return: { ok, ditulis, dilewati, alasan, buffer }
 */
function bangunBoqTamkur(rows) {
  const hasil = { ok: false, ditulis: 0, dilewati: [], alasan: '', buffer: null };

  if (!fs.existsSync(PATH_XLSX)) {
    hasil.alasan = 'templates/BOQ TAMKUR.xlsx tidak ditemukan';
    return hasil;
  }
  if (!Array.isArray(rows) || !rows.length) {
    hasil.alasan = 'tidak ada baris rekon untuk ditulis';
    return hasil;
  }

  try {
    const zip = new PizZip(fs.readFileSync(PATH_XLSX)); // read-only
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

    siapSheetBoq(zip, sheetPath, sheetXml);
    hasil.buffer = zip.generate({ type: 'nodebuffer', compression: 'DEFLATE' });
    hasil.ok = true;
    return hasil;
  } catch (e) {
    hasil.alasan = e.message;
    return hasil;
  }
}

/**
 * bacaHargaBoqTamkur()
 * Harga satuan diambil dari templates/BOQ TAMKUR.xlsx:
 *   • designator M- di kolom C → harga material kolom G
 *   • designator J- di kolom D → harga jasa    kolom H
 * Return: { "M-...": 7098, "J-...": 2169, ... }
 */
function bacaHargaBoqTamkur() {
  const peta = {};
  if (!fs.existsSync(PATH_XLSX)) return peta;
  try {
    const zip = new PizZip(fs.readFileSync(PATH_XLSX));
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
    console.warn('[BOQ TAMKUR] gagal membaca harga:', e.message);
  }
  return peta;
}

module.exports = {
  tulisBoqTamkur, bacaHargaBoqTamkur, bangunBoqTamkur,
  PATH_XLSX, NAMA_SHEET, bacaShared, lepasTeks, indeksSheet, cariBaris
};
