// Lapisan 2 — pemeriksaan forensik piksel sederhana (Error Level Analysis).
// File disimpan ulang pada kualitas tinggi lalu dibandingkan piksel demi piksel:
// area yang "menyala" kuat dicurigai pernah diedit.
const fs = require('fs');
const path = require('path');
const sharp = require('sharp');
const cfg = require('./verifyConfig');

async function muatPiksel(sumber) {
  const buf = Buffer.isBuffer(sumber) ? sumber : await fs.promises.readFile(sumber);
  const { data, info } = await sharp(buf)
    .flatten({ background: { r: 255, g: 255, b: 255 } })
    .removeAlpha()
    .toColourspace('srgb')
    .raw()
    .toBuffer({ resolveWithObject: true });
  return { data, lebar: info.width, tinggi: info.height, channels: info.channels };
}

// Analisis artefak blok 8×8 pada area datar (rata-rata/varian rendah).
function hitungBlok8(asli) {
  try {
    const { data, lebar: w, tinggi: h, channels: ch } = asli;
    if (w < 64 || h < 64) return null;
    const lp = new Float64Array(w * h);
    for (let p = 0, i = 0; p < w * h; p++, i += ch) {
      lp[p] = ch >= 3 ? 0.299 * data[i] + 0.587 * data[i + 1] + 0.114 * data[i + 2] : data[i];
    }
    const bs = 8;
    let blokDatar = 0, sumBatas = 0, nBatas = 0, sumDalam = 0, nDalam = 0;
    for (let by = 0; by + bs <= h; by += bs) {
      for (let bx = 0; bx + bs <= w; bx += bs) {
        let s = 0, s2 = 0;
        for (let y = 0; y < bs; y++) {
          for (let x = 0; x < bs; x++) {
            const v = lp[(by + y) * w + bx + x];
            s += v; s2 += v * v;
          }
        }
        const rata = s / 64;
        const varian = s2 / 64 - rata * rata;
        if (varian > 300 || rata < 20 || rata > 235) continue; // bukan area datar
        blokDatar++;
        for (let y = 0; y < bs; y++) {
          const baris = (by + y) * w;
          if (bx > 0) { sumBatas += Math.abs(lp[baris + bx] - lp[baris + bx - 1]); nBatas++; }
          for (let x = bx + 1; x < bx + bs; x++) {
            sumDalam += Math.abs(lp[baris + x] - lp[baris + x - 1]); nDalam++;
          }
        }
        for (let x = 0; x < bs; x++) {
          if (by > 0) { sumBatas += Math.abs(lp[by * w + bx + x] - lp[(by - 1) * w + bx + x]); nBatas++; }
          for (let y = by + 1; y < by + bs; y++) {
            sumDalam += Math.abs(lp[y * w + bx + x] - lp[(y - 1) * w + bx + x]); nDalam++;
          }
        }
      }
    }
    if (!blokDatar || !nBatas || !nDalam) return { blokDatar, rasio: null, gradienBatas: 0, gradienDalam: 0 };
    const batas = sumBatas / nBatas;
    const dalam = sumDalam / nDalam;
    return {
      blokDatar,
      gradienBatas: Math.round(batas * 100) / 100,
      gradienDalam: Math.round(dalam * 100) / 100,
      rasio: dalam > 0.05 ? Math.round((batas / dalam) * 100) / 100 : null
    };
  } catch (e) {
    return null;
  }
}

async function ela(filePath, opsi) {
  const o = Object.assign({}, cfg.ela, opsi || {});
  const hasil = { skor: 0, anomali: 0, flags: [], alasan: [], heatmapPath: null, blok: null };
  try {
    // Buffer dulu — sharp(path) menahan handle berkas di Windows
    const sumber = Buffer.isBuffer(filePath) ? filePath : await fs.promises.readFile(filePath);
    const asli = await muatPiksel(sumber);
    const ulangBuf = await sharp(sumber).jpeg({ quality: o.kualitasUlang }).toBuffer();
    const ulang = await muatPiksel(ulangBuf);

    if (asli.lebar !== ulang.lebar || asli.tinggi !== ulang.tinggi || asli.channels !== ulang.channels) {
      hasil.alasan.push('Ukuran gambar berubah saat pemeriksaan ELA — pemeriksaan dilewati.');
      return hasil;
    }

    const n = asli.lebar * asli.tinggi * asli.channels;
    const diff = new Float64Array(asli.lebar * asli.tinggi);
    let total = 0;
    for (let i = 0, p = 0; i < n; i += asli.channels, p++) {
      let maks = 0;
      for (let c = 0; c < asli.channels; c++) {
        const d = Math.abs(asli.data[i + c] - ulang.data[i + c]);
        if (d > maks) maks = d;
      }
      diff[p] = maks;
      total += maks;
    }

    const mean = total / (asli.lebar * asli.tinggi);
    hasil.skor = Math.min(100, Math.round(mean * 4));

    // Blok 16x16: fraksi blok yang rata-ratanya jauh di atas rata-rata global
    const bs = 16;
    const batas = Math.max(3 * mean, 4);
    let blokBlok = 0, blokBlokAnomali = 0;
    for (let by = 0; by + bs <= asli.tinggi; by += bs) {
      for (let bx = 0; bx + bs <= asli.lebar; bx += bs) {
        let s = 0;
        for (let y = 0; y < bs; y++)
          for (let x = 0; x < bs; x++) s += diff[(by + y) * asli.lebar + bx + x];
        const rata = s / (bs * bs);
        blokBlok++;
        if (rata > batas) blokBlokAnomali++;
      }
    }
    hasil.anomali = blokBlok ? blokBlokAnomali / blokBlok : 0;

    if (hasil.anomali > o.ambangAnomali && hasil.skor >= o.skorMinimalTinggi) {
      hasil.flags.push('ela_tinggi');
      hasil.alasan.push(`Anomali ELA ${Math.round(hasil.anomali * 100)}% dari area foto (skor ${hasil.skor}).`);
    }

    // Artefak blok 8×8 pada area datar — bekas kompresi ulang/upscale/tempelan.
    // Gradien pada batas blok (x%8==0 atau y%8==0) jauh lebih kuat daripada
    // gradien di dalam blok → rasio tinggi menandakan kompresi ulang berulang.
    const analisisBlok = hitungBlok8(asli);
    if (analisisBlok) {
      hasil.blok = analisisBlok;
      if (analisisBlok.blokDatar >= o.blokMinimal &&
          analisisBlok.rasio !== null &&
          analisisBlok.rasio >= o.blokRasio &&
          analisisBlok.gradienBatas >= 1) {
        hasil.flags.push('blok_artefak');
        hasil.alasan.push(`Artefak blok 8×8 kuat — gradien batas blok ${analisisBlok.rasio.toFixed(2)}× lebih tinggi dari dalam blok pada ${analisisBlok.blokDatar} area datar.`);
      }
    }

    if (o.heatmap) hasil.heatmapPath = await simpanHeatmap(diff, asli, o);
  } catch (e) {
    hasil.alasan.push('Pemeriksaan ELA gagal: ' + (e.message || e));
  }
  return hasil;
}

async function simpanHeatmap(diff, asli, o) {
  try {
    const buf = Buffer.alloc(asli.lebar * asli.tinggi);
    let maks = 1;
    for (let i = 0; i < diff.length; i++) if (diff[i] > maks) maks = diff[i];
    for (let i = 0; i < diff.length; i++) buf[i] = Math.round((diff[i] / maks) * 255);
    const folder = path.join(__dirname, '..', 'uploads', '_ela');
    fs.mkdirSync(folder, { recursive: true });
    const nama = `${Date.now()}-${Math.random().toString(36).slice(2, 8)}.jpg`;
    await sharp(buf, { raw: { width: asli.lebar, height: asli.tinggi, channels: 1 } })
      .resize(512, 512, { fit: 'inside' })
      .jpeg({ quality: 70 })
      .toFile(path.join(folder, nama));
    return `/uploads/_ela/${nama}`;
  } catch (e) {
    return null;
  }
}

module.exports = { ela };
