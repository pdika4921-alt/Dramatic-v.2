// Lapisan 2 (bagian kualitas) — ketajaman (varians Laplacian), kecerahan,
// resolusi minimum, kecocokan format berkas, dan estimasi kualitas JPEG.
const fs = require('fs');
const sharp = require('sharp');
const cfg = require('./verifyConfig');

/* ── Deteksi format dari isi berkas (magic bytes) ────────────────────────── */
const FORMAT_DARI_MAGIC = [
  { nama: 'jpeg', cek: b => b[0] === 0xff && b[1] === 0xd8 && b[2] === 0xff },
  { nama: 'png', cek: b => b[0] === 0x89 && b[1] === 0x50 && b[2] === 0x4e && b[3] === 0x47 },
  { nama: 'gif', cek: b => b[0] === 0x47 && b[1] === 0x49 && b[2] === 0x46 },
  { nama: 'webp', cek: b => b.slice(0, 4).toString('latin1') === 'RIFF' && b.slice(8, 12).toString('latin1') === 'WEBP' },
  { nama: 'avif', cek: b => b.slice(4, 8).toString('latin1') === 'ftyp' },
  { nama: 'bmp', cek: b => b[0] === 0x42 && b[1] === 0x4d },
  { nama: 'tiff', cek: b => (b[0] === 0x49 && b[1] === 0x49) || (b[0] === 0x4d && b[1] === 0x4d) }
];

function formatDariMagic(buf) {
  for (const f of FORMAT_DARI_MAGIC) {
    try { if (f.cek(buf)) return f.nama; } catch (e) { /* buffer terlalu pendek */ }
  }
  return null;
}

function formatDariEkstensi(namaFile) {
  const m = /\.([a-z0-9]+)$/i.exec(String(namaFile || ''));
  const e = m ? m[1].toLowerCase() : '';
  if (e === 'jpg' || e === 'jpeg') return 'jpeg';
  if (e === 'jfif') return 'jpeg';
  return e || null;
}

/* ── Estimasi kualitas JPEG dari tabel kuantisasi (DQT) ─────────────────── */
// Tabel luminans standar IJG (kualitas 50)
const TABEL_STANDAR = [
  16, 11, 10, 16, 24, 40, 51, 61,
  12, 12, 14, 19, 26, 58, 60, 55,
  14, 13, 16, 24, 40, 57, 69, 56,
  14, 17, 22, 29, 51, 87, 80, 62,
  18, 22, 37, 56, 68, 109, 103, 77,
  24, 35, 55, 64, 81, 104, 113, 92,
  49, 64, 78, 87, 103, 121, 120, 101,
  72, 92, 95, 98, 112, 100, 103, 99
];

function tabelUntukKualitas(q) {
  const F = q < 50 ? Math.floor(5000 / q) : 200 - 2 * q;
  return TABEL_STANDAR.map(v => {
    const t = Math.floor((v * F + 50) / 100);
    return Math.max(1, Math.min(255, t));
  });
}

function bacaDqtLuma(buf) {
  let pos = 2; // lewati SOI
  while (pos + 4 < buf.length) {
    if (buf[pos] !== 0xff) { pos++; continue; }
    const marker = buf[pos + 1];
    if (marker === 0xd8 || marker === 0xd9 || marker === 0x01 || (marker >= 0xd0 && marker <= 0xd7)) { pos += 2; continue; }
    const panjang = buf.readUInt16BE(pos + 2);
    if (marker === 0xdb) {
      let p = pos + 4;
      const akhir = pos + 2 + panjang;
      while (p < akhir) {
        const info = buf[p];
        const presisi = info >> 4;
        const id = info & 0x0f;
        const lebar = presisi === 1 ? 128 : 64;
        if (id === 0 && lebar === 64) return Array.from({ length: 64 }, (_, i) => buf[p + 1 + i]);
        p += 1 + lebar;
      }
    }
    if (marker === 0xda) break; // awal data scan — cukup
    pos += 2 + panjang;
  }
  return null;
}

function estimasiKualitasJpeg(buf) {
  try {
    const luma = bacaDqtLuma(buf);
    if (!luma || luma.length !== 64) return null;
    let terbaik = null;
    for (let q = 1; q <= 100; q++) {
      const t = tabelUntukKualitas(q);
      let selisih = 0;
      for (let i = 0; i < 64; i++) selisih += Math.abs(t[i] - luma[i]);
      if (!terbaik || selisih < terbaik.selisih) terbaik = { q, selisih };
    }
    if (!terbaik) return null;
    return {
      kualitas: terbaik.q,
      tabelKustom: terbaik.selisih / 64 > 8 // tabel tidak menyerupai standar IJG — biasanya hasil ekspor editor
    };
  } catch (e) {
    return null;
  }
}

/* ── Ketajaman (varians Laplacian) & kecerahan dari piksel greyscale ─────── */
function analisisPiksel(data, lebar, tinggi) {
  if (lebar < 16 || tinggi < 16) return { ketajaman: null, luminans: null };
  let jumlah = 0;
  for (let i = 0; i < data.length; i++) jumlah += data[i];
  const luminans = jumlah / data.length;

  // Varians Laplacian 4-tetangga pada area interior
  let n = 0, sum = 0, sum2 = 0;
  for (let y = 1; y < tinggi - 1; y++) {
    for (let x = 1; x < lebar - 1; x++) {
      const i = y * lebar + x;
      const lap = 4 * data[i] - data[i - 1] - data[i + 1] - data[i - lebar] - data[i + lebar];
      sum += lap;
      sum2 += lap * lap;
      n++;
    }
  }
  if (!n) return { ketajaman: null, luminans };
  const rata = sum / n;
  const varian = sum2 / n - rata * rata;
  return { ketajaman: Math.round(Math.max(0, varian) * 10) / 10, luminans: Math.round(luminans * 10) / 10 };
}

/* ── Pemeriksaan utama ───────────────────────────────────────────────────── */
async function cekKualitas(filePath, namaFile) {
  const hasil = { flags: [], alasan: [], detail: null };
  try {
    const sumber = Buffer.isBuffer(filePath) ? filePath : await fs.promises.readFile(filePath);
    const meta = await sharp(sumber).metadata();

    hasil.detail = {
      lebar: meta.width || null,
      tinggi: meta.height || null,
      piksel: (meta.width && meta.height) ? meta.width * meta.height : null,
      formatIsi: formatDariMagic(sumber),
      formatEkstensi: formatDariEkstensi(namaFile),
      ketajaman: null,
      luminans: null,
      jpegKualitas: null
    };

    // 1. Format isi vs ekstensi
    const fmt = cekFormat(sumber, namaFile);
    if (fmt) {
      hasil.detail.formatSesuai = fmt.sesuai;
      if (fmt.sesuai === false) {
        hasil.flags.push('format_tidak_sesuai');
        hasil.alasan.push(`Format berkas tidak sesuai ekstensi — isi ${fmt.isi || 'tidak dikenal'}, ekstensi .${fmt.ext}.`);
      }
    }

    if (hasil.detail.formatIsi === 'jpeg') {
      const q = estimasiKualitasJpeg(sumber);
      if (q) hasil.detail.jpegKualitas = q.kualitas;
    }

    // 2. Ketajaman & kecerahan (hanya untuk gambar raster)
    if (meta.width && meta.height && meta.width * meta.height <= 40e6) {
      const { data } = await sharp(sumber)
        .flatten({ background: { r: 255, g: 255, b: 255 } })
        .greyscale()
        .raw()
        .toBuffer({ resolveWithObject: true });
      const a = analisisPiksel(data, meta.width, meta.height);
      hasil.detail.ketajaman = a.ketajaman;
      hasil.detail.luminans = a.luminans;

      if (a.ketajaman !== null && a.ketajaman < cfg.kualitas.blurMin) {
        hasil.flags.push('foto_buram');
        hasil.alasan.push(`Foto tidak tajam — ketajaman ${a.ketajaman} (minimal ${cfg.kualitas.blurMin}).`);
      }
      if (a.luminans !== null && a.luminans < cfg.kualitas.lumaGelap) {
        hasil.flags.push('foto_gelap');
        hasil.alasan.push(`Foto gelap — luminans rata-rata ${a.luminans} (minimal ${cfg.kualitas.lumaGelap}).`);
      }
    }

    // 3. Resolusi minimum (sisi terkecil)
    if (meta.width && meta.height) {
      const sisi = Math.min(meta.width, meta.height);
      if (sisi < cfg.kualitas.minSisi) {
        hasil.flags.push('resolusi_rendah');
        hasil.alasan.push(`Resolusi rendah — ${meta.width}×${meta.height} px, sisi terkecil ${sisi} < ${cfg.kualitas.minSisi}.`);
      }
    }
  } catch (e) {
    hasil.alasan.push('Pemeriksaan kualitas gagal: ' + (e.message || e));
  }
  return hasil;
}

// Kecocokan format diperiksa terpisah karena butuh nama berkas
function cekFormat(sumber, namaFile) {
  try {
    const buf = Buffer.isBuffer(sumber) ? sumber : null;
    if (!buf) return null;
    const isi = formatDariMagic(buf);
    const ext = formatDariEkstensi(namaFile);
    if (!isi || !ext) return { isi, ext, sesuai: null };
    const setara = (isi === 'jpeg' && ext === 'jpeg') || isi === ext;
    return { isi, ext, sesuai: setara };
  } catch (e) {
    return null;
  }
}

module.exports = { cekKualitas, cekFormat, estimasiKualitasJpeg, formatDariMagic, formatDariEkstensi };
