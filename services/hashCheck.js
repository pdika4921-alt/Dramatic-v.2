// Lapisan 1 (bagian hash) — SHA-256 + perceptual hash (dHash & pHash) + jarak Haversine.
const crypto = require('crypto');
const fs = require('fs');
const sharp = require('sharp');

function sha256(buf) {
  return crypto.createHash('sha256').update(buf).digest('hex');
}

function hexDariBit(bit) {
  let hex = '';
  for (let i = 0; i < bit.length; i += 4) {
    hex += parseInt(bit.slice(i, i + 4), 2).toString(16);
  }
  return hex;
}

function bitDariHex(hex) {
  return String(hex || '').split('').map(c => parseInt(c, 16).toString(2).padStart(4, '0')).join('');
}

async function pikselAbu(filePath, sisi) {
  // Baca ke buffer dulu: sharp(path) menahan handle berkas di Windows
  // sehingga file tidak bisa dihapus/dipindah selama server berjalan.
  const buf = Buffer.isBuffer(filePath) ? filePath : await fs.promises.readFile(filePath);
  const { data } = await sharp(buf)
    .resize(sisi, sisi, { fit: 'fill' })
    .greyscale()
    .raw()
    .toBuffer({ resolveWithObject: true });
  return data;
}

// dHash: bandingkan piksel bersebelahan pada gambar 9x8 → 64 bit
async function dhash(filePath) {
  const d = await pikselAbu(filePath, 9);
  let bit = '';
  for (let y = 0; y < 8; y++) {
    for (let x = 0; x < 8; x++) {
      bit += d[y * 9 + x] > d[y * 9 + x + 1] ? '1' : '0';
    }
  }
  return hexDariBit(bit);
}

// pHash: DCT 32x32 → ambil 8x8 frekuensi rendah → hash median
function cosTab(n) {
  const t = [];
  for (let u = 0; u < n; u++) {
    t[u] = [];
    for (let x = 0; x < n; x++) t[u][x] = Math.cos(((2 * x + 1) * u * Math.PI) / (2 * n));
  }
  return t;
}

async function phash(filePath) {
  const n = 32;
  const p = await pikselAbu(filePath, n);
  const cos = cosTab(n);

  // DCT 2D: baris lalu kolom (n kecil — perhitungan langsung sudah cukup cepat)
  const tmp = new Float64Array(n * n);
  for (let y = 0; y < n; y++) {
    for (let u = 0; u < n; u++) {
      let sum = 0;
      for (let x = 0; x < n; x++) sum += p[y * n + x] * cos[u][x];
      tmp[y * n + u] = sum;
    }
  }
  const matriks = new Float64Array(n * n);
  for (let u = 0; u < n; u++) {
    for (let v = 0; v < n; v++) {
      let sum = 0;
      for (let y = 0; y < n; y++) sum += tmp[y * n + u] * cos[v][y];
      matriks[v * n + u] = sum;
    }
  }

  const nilai = [];
  for (let y = 0; y < 8; y++) {
    for (let x = 0; x < 8; x++) {
      if (x === 0 && y === 0) continue; // buang komponen DC
      nilai.push(matriks[y * n + x]);
    }
  }
  const urut = nilai.slice().sort((a, b) => a - b);
  const median = urut[Math.floor(urut.length / 2)];
  let bit = '';
  for (const v of nilai) bit += v > median ? '1' : '0';
  return hexDariBit(bit);
}

function hamming(a, b) {
  const ba = bitDariHex(a), bb = bitDariHex(b);
  if (ba.length !== bb.length || !ba.length) return 64;
  let n = 0;
  for (let i = 0; i < ba.length; i++) if (ba[i] !== bb[i]) n++;
  return n;
}

function haversineMeter(lat1, lng1, lat2, lng2) {
  const R = 6371000;
  const rad = d => (d * Math.PI) / 180;
  const dLat = rad(lat2 - lat1);
  const dLng = rad(lng2 - lng1);
  const a = Math.sin(dLat / 2) ** 2 +
    Math.cos(rad(lat1)) * Math.cos(rad(lat2)) * Math.sin(dLng / 2) ** 2;
  return 2 * R * Math.asin(Math.sqrt(a));
}

async function hitungHash(filePath) {
  const buf = Buffer.isBuffer(filePath) ? filePath : fs.readFileSync(filePath);
  const [d, p] = await Promise.all([dhash(buf), phash(buf)]);
  return { sha256: sha256(buf), dhash: d, phash: p };
}

module.exports = { sha256, dhash, phash, hamming, haversineMeter, hitungHash };
