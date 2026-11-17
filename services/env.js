// Loader berkas .env sederhana tanpa dependency tambahan.
// Nilai yang sudah ada di process.env TIDAK ditimpa (aman untuk uji/produksi).
const fs = require('fs');
const path = require('path');

let sudahDimuat = false;

function loadEnv(file) {
  if (sudahDimuat) return;
  sudahDimuat = true;
  const target = file || path.join(__dirname, '..', '.env');
  let isi;
  try {
    isi = fs.readFileSync(target, 'utf8');
  } catch (e) {
    return; // .env opsional — tidak wajib ada
  }
  for (const baris of isi.split(/\r?\n/)) {
    const m = /^\s*([A-Za-z_][A-Za-z0-9_]*)\s*=\s*(.*)\s*$/.exec(baris);
    if (!m) continue;
    if (process.env[m[1]] !== undefined) continue;
    let nilai = m[2].trim();
    if ((nilai.startsWith('"') && nilai.endsWith('"')) || (nilai.startsWith("'") && nilai.endsWith("'"))) {
      nilai = nilai.slice(1, -1);
    }
    process.env[m[1]] = nilai;
  }
}

module.exports = { loadEnv };
