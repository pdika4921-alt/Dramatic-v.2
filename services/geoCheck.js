// Lapisan 1 (bagian lokasi) — reverse geocoding Nominatim/OSM untuk memeriksa
// apakah tag lokasi (EXIF GPS) foto sesuai dengan nama lokasi pekerjaan.
// Hasil: 'cocok' | 'tidak_cocok' | 'netral' (tanpa GPS / nama lokasi berupa kode /
// geocoding gagal — dinilai netral, tidak memberi penalti selain flag ringan).
const cfg = require('./verifyConfig');

/* ── Tabel cache koordinat → hasil geocoding ─────────────────────────────── */
async function initTabel(db) {
  await db.exec(`CREATE TABLE IF NOT EXISTS geo_cache (
    lat_r       TEXT NOT NULL,
    lng_r       TEXT NOT NULL,
    hasil       TEXT,
    err         TEXT,
    created_at  TEXT DEFAULT (datetime('now','localtime')),
    PRIMARY KEY (lat_r, lng_r)
  )`);
  console.log(' Tabel geo_cache siap.');
}

/* ── Token nama wilayah dari nama lokasi ─────────────────────────────────── */
// Kata umum proyek/kode yang bukan nama wilayah
const STOPWORD = new Set([
  'relokasi', 'eviden', 'lokasi', 'kerja', 'pekerjaan', 'titik', 'kavling',
  'expand', 'golive', 'datin', 'booo', 'odp', 'odc', 'olt', 'feeder', 'rev',
  'stok', 'sto', 'passive', 'active', 'network', 'proyek', 'laporan', 'foto',
  'sebelum', 'sesudah', 'progres', 'cutover', 'instalasi', 'pemasangan'
]);

function tokenNamaLokasi(nama) {
  const potong = String(nama || '').toLowerCase().split(/[^a-z0-9]+/).filter(Boolean);
  const hasil = new Set();
  for (const t of potong) {
    if (t.length < 4) continue;
    if (!/[a-z]/.test(t)) continue;
    if (STOPWORD.has(t)) continue;
    hasil.add(t);
  }
  return [...hasil];
}

// Indeks frekuensi token di seluruh nama lokasi — token yang muncul di banyak
// nama berbeda umumnya kode proyek, bukan nama wilayah (hindari alarm palsu).
let indeksToken = { isi: null, peta: new Map(), saat: 0 };
async function bangunIndeksToken(db) {
  const TTL = 10 * 60 * 1000;
  if (indeksToken.peta.size && Date.now() - indeksToken.saat < TTL) return indeksToken.peta;
  try {
    const rows = await db.all('SELECT nama_lokasi FROM lokasi');
    const peta = new Map();
    for (const r of rows) {
      for (const t of tokenNamaLokasi(r.nama_lokasi)) peta.set(t, (peta.get(t) || 0) + 1);
    }
    indeksToken = { isi: rows.length, peta, saat: Date.now() };
    return peta;
  } catch (e) {
    return indeksToken.peta || new Map();
  }
}

/* ── Antrean panggilan (kebijakan Nominatim: maksimal ~1 req/detik) ──────── */
const tunggu = ms => new Promise(r => setTimeout(r, ms));
let antreTerakhir = Promise.resolve();
let waktuPanggil = 0;
function jadwal(fn) {
  const jalan = antreTerakhir.then(async () => {
    const sisa = cfg.geo.jedaMs - (Date.now() - waktuPanggil);
    if (sisa > 0) await tunggu(sisa);
    waktuPanggil = Date.now();
    return fn();
  });
  antreTerakhir = jalan.catch(() => {});
  return jalan;
}

function kunciKoordinat(lat, lng) {
  return { lat: lat.toFixed(3), lng: lng.toFixed(3) }; // ≈110 m
}

async function ambilCache(db, lat, lng) {
  const k = kunciKoordinat(lat, lng);
  try {
    const row = await db.get('SELECT hasil, err, created_at FROM geo_cache WHERE lat_r = ? AND lng_r = ?', [k.lat, k.lng]);
    if (!row) return null;
    if (row.hasil) { try { return { data: JSON.parse(row.hasil) }; } catch (e) { return null; } }
    if (row.err) {
      // Galat sementara hanya di-cache 5 menit, lalu dicoba lagi
      const umur = Date.now() - new Date((row.created_at || '').replace(' ', 'T')).getTime();
      if (Number.isFinite(umur) && umur < 5 * 60 * 1000) return { error: row.err };
      return null;
    }
    return null;
  } catch (e) {
    return null;
  }
}

async function simpanCache(db, lat, lng, data, err) {
  const k = kunciKoordinat(lat, lng);
  try {
    await db.run(
      `INSERT INTO geo_cache (lat_r, lng_r, hasil, err, created_at) VALUES (?,?,?,?,datetime('now','localtime'))
       ON CONFLICT(lat_r, lng_r) DO UPDATE SET hasil=excluded.hasil, err=excluded.err, created_at=excluded.created_at`,
      [k.lat, k.lng, data ? JSON.stringify(data) : null, err || null]
    );
  } catch (e) { /* cache gagal bukan masalah fatal */ }
}

async function panggilNominatim(lat, lng) {
  const url = `${cfg.geo.url}?format=jsonv2&lat=${encodeURIComponent(lat)}&lon=${encodeURIComponent(lng)}&zoom=18&addressdetails=1&accept-language=id`;
  const res = await fetch(url, {
    headers: { 'User-Agent': cfg.geo.userAgent, 'Accept-Language': 'id' },
    signal: AbortSignal.timeout(cfg.geo.timeoutMs)
  });
  if (!res.ok) throw new Error('Nominatim HTTP ' + res.status);
  const d = await res.json();
  if (d && d.error) throw new Error('Nominatim: ' + String(d.error).slice(0, 80));
  return {
    display_name: d && d.display_name ? String(d.display_name) : '',
    address: (d && d.address) || {}
  };
}

function teksGeocode(data) {
  if (!data) return '';
  const bagian = [data.display_name || ''];
  for (const v of Object.values(data.address || {})) if (v) bagian.push(String(v));
  return bagian.join(' ').toLowerCase();
}

/* ── Pemeriksaan utama ───────────────────────────────────────────────────── */
async function cekLokasi(db, { lat, lng, namaLokasi }) {
  const dasar = { hasil: 'netral', alamat: null, tokenWilayah: [], tokenCocok: [], teks: '' };

  if (!Number.isFinite(lat) || !Number.isFinite(lng)) {
    return Object.assign(dasar, { teks: 'Tidak ada GPS pada foto — verifikasi lokasi dinilai netral.' });
  }
  if (!cfg.geo.aktif) {
    return Object.assign(dasar, { teks: 'Reverse geocoding dimatikan (GEO_NOMINATIM=0) — lokasi dinilai netral.' });
  }

  const token = tokenNamaLokasi(namaLokasi);
  const frekuensi = await bangunIndeksToken(db);
  // Token kemungkinan nama wilayah: muncul di ≤2 nama lokasi berbeda
  const tokenWilayah = token.filter(t => (frekuensi.get(t) || 0) <= 2);

  // Ambil hasil geocoding (cache → antrean → jaringan)
  let data = null;
  let error = null;
  const cache = await ambilCache(db, lat, lng);
  if (cache && cache.data) {
    data = cache.data;
  } else if (cache && cache.error && cache.error !== 'sebelumnya gagal') {
    error = cache.error;
  } else {
    try {
      data = await jadwal(() => panggilNominatim(lat, lng));
      await simpanCache(db, lat, lng, data, null);
    } catch (e) {
      error = e.message || String(e);
      await simpanCache(db, lat, lng, null, error);
    }
  }

  if (!data) {
    return Object.assign(dasar, {
      teks: `Geocoding gagal (${error}) — verifikasi lokasi dinilai netral.`
    });
  }

  const geoteks = teksGeocode(data);
  const alamat = data.display_name || '(tanpa nama)';
  const cocok = token.filter(t => geoteks.includes(t));

  if (cocok.length) {
    return {
      hasil: 'cocok',
      alamat,
      tokenWilayah,
      tokenCocok: cocok,
      teks: `Cocok — GPS berada di "${alamat}" dan sesuai nama lokasi (token: ${cocok.join(', ')}).`
    };
  }
  if (tokenWilayah.length) {
    return {
      hasil: 'tidak_cocok',
      alamat,
      tokenWilayah,
      tokenCocok: [],
      teks: `Tidak cocok — GPS berada di "${alamat}" tetapi tidak menyebut nama wilayah dari lokasi ini (${tokenWilayah.join(', ')}).`
    };
  }
  return {
    hasil: 'netral',
    alamat,
    tokenWilayah: [],
    tokenCocok: [],
    teks: `Nama lokasi tidak mengandung nama wilayah yang bisa dibandingkan — GPS di "${alamat}", dinilai netral.`
  };
}

module.exports = { initTabel, cekLokasi, tokenNamaLokasi };
