// Konfigurasi terpusat fitur "Verifikasi Keaslian Eviden".
// Semua bobot/ambang diubah lewat file ini atau variabel lingkungan (.env).
const { loadEnv } = require('./env');

loadEnv();

function angka(nilai, bawaan) {
  const n = Number(nilai);
  return Number.isFinite(n) ? n : bawaan;
}

const STATUS = {
  VALID: 'valid',
  CEK: 'cek',
  CURIGA: 'curiga',
  BELUM: 'belum'
};

// Daftar flag verifikasi: bobot = pengurang skor per lapisan, berat = auto DICURIGAI
const FLAG = {
  software_editor:  { berat: true,  bobot: 45, label: 'Software editor terdeteksi pada EXIF' },
  gps_jauh:         { berat: true,  bobot: 40, label: 'GPS foto jauh dari titik lokasi pekerjaan' },
  duplikat_lokasi_lain: { berat: true, bobot: 45, label: 'Foto identik terdeteksi di lokasi/titik lain' },
  foto_layar:       { berat: true,  bobot: 35, label: 'AI mendeteksi indikasi foto layar/screenshot' },
  editan_tinggi:    { berat: true,  bobot: 50, label: 'AI mendeteksi indikasi manipulasi kuat' },
  tanggal_luar:     { berat: false, bobot: 20, label: 'Tanggal foto di luar rentang pekerjaan lokasi' },
  metadata_hilang:  { berat: false, bobot: 12, label: 'Metadata EXIF tidak ada pada berkas' },
  kamera_bervariasi:{ berat: false, bobot: 10, label: 'Kamera berbeda dalam satu kategori eviden' },
  ela_tinggi:       { berat: false, bobot: 12, label: 'Anomali kompresi (ELA) tinggi — ada area dicurigai' },
  stempel_tak_terbaca: { berat: false, bobot: 10, label: 'Stempel/tanggal pada foto tidak terbaca' },
  isi_tidak_sesuai: { berat: false, bobot: 20, label: 'Isi foto dinilai tidak sesuai pekerjaan' },
  kualitas_ragu:    { berat: false, bobot: 6,  label: 'AI ragu menilai kualitas/keaslian foto' },

  // Flag baru — cek lokasi (reverse geocoding)
  lokasi_tidak_cocok: { berat: false, bobot: 30, label: 'Tag lokasi (GPS) tidak cocok dengan nama lokasi' },
  // Flag baru — konsistensi metadata
  metadata_dismatch: { berat: false, bobot: 12, label: 'Dimensi pada EXIF tidak sama dengan ukuran file' },
  // Flag baru — struktur & kualitas gambar (Lapisan 2)
  format_tidak_sesuai: { berat: false, bobot: 10, label: 'Isi berkas tidak sesuai ekstensi (format salah)' },
  blok_artefak:     { berat: false, bobot: 10, label: 'Artefak blok 8×8 kuat — bekas kompresi ulang/tempelan' },
  foto_buram:       { berat: false, bobot: 10, label: 'Foto buram / tidak tajam' },
  foto_gelap:       { berat: false, bobot: 8,  label: 'Foto terlalu gelap' },
  resolusi_rendah:  { berat: false, bobot: 8,  label: 'Resolusi foto di bawah standar' },
  // Flag baru — hasil AI tambahan
  latar_mencurigakan: { berat: false, bobot: 15, label: 'AI mencurigai latar belakang tidak wajar' }
};

// Flag yang berasal dari Lapisan 1 (metadata & konsistensi)
const FLAG_LAPIS1 = ['software_editor', 'gps_jauh', 'tanggal_luar', 'metadata_hilang', 'kamera_bervariasi', 'duplikat_lokasi_lain', 'metadata_dismatch', 'lokasi_tidak_cocok'];
// Flag yang berasal dari Lapisan 2 (forensik piksel/ELA + kualitas)
const FLAG_LAPIS2 = ['ela_tinggi', 'format_tidak_sesuai', 'blok_artefak', 'foto_buram', 'foto_gelap', 'resolusi_rendah'];

// Pemetaan flag → check ringkasan (dipakai UI galeri: popover & strip alasan)
const CHECK_FLAG = {
  duplikat: ['duplikat_lokasi_lain'],
  lokasi:   ['lokasi_tidak_cocok', 'gps_jauh'],
  exif:     ['software_editor', 'metadata_dismatch', 'metadata_hilang', 'tanggal_luar', 'kamera_bervariasi'],
  kualitas: ['foto_buram', 'foto_gelap', 'resolusi_rendah', 'format_tidak_sesuai'],
  forensik: ['ela_tinggi', 'blok_artefak'],
  isi:      ['isi_tidak_sesuai'],
  ai:       ['editan_tinggi', 'foto_layar', 'latar_mencurigakan', 'stempel_tak_terbaca', 'kualitas_ragu']
};

const SOFTWARE_EDITOR = [
  'photoshop', 'lightroom', 'gimp', 'canva', 'picsart', 'snapseed', 'inshot',
  'photogadget', 'photo editor', 'photoscape', 'paint.net', 'microsoft paint',
  'faceapp', 'picsart', 'snapedit', 'retouch', 'watermark',
  'meitu', 'pixelcut', 'photopea', 'beautyplus', 'fotor', 'photolab',
  'photo director', 'photodirector', 'cyberlink', 'corel paint'
];

const cfg = {
  status: STATUS,
  flag: FLAG,
  flagLapis1: FLAG_LAPIS1,
  flagLapis2: FLAG_LAPIS2,
  checkFlag: CHECK_FLAG,
  softwareEditor: SOFTWARE_EDITOR,

  // Lapisan 4 — ambang status akhir
  ambang: {
    valid: angka(process.env.VERIFY_AMBANG_VALID, 80),
    curiga: angka(process.env.VERIFY_AMBANG_CURIGA, 50)
  },
  bobot: {
    ai: 0.55,
    lapis1: 0.30,
    lapis2: 0.15
  },

  // Lapisan 1 — toleransi jarak GPS ke titik lokasi (meter)
  toleransiGpsMeter: angka(process.env.VERIFY_GPS_TOLERANCE_M, 100),

  // Lapisan 1 — deteksi duplikat: jarak hamming maksimal dari 64 bit
  // (1 ≈ kemiripan ≥98,4% ≈ "99%"); SHA-256 identik selalu dianggap duplikat
  duplikat: {
    maksHamming: angka(process.env.VERIFY_DUP_MAX_HAMMING, 1)
  },

  // Lapisan 1 — reverse geocoding (Nominatim/OSM) untuk cek tag lokasi
  geo: {
    aktif: process.env.GEO_NOMINATIM !== '0',
    url: process.env.GEO_NOMINATIM_URL || 'https://nominatim.openstreetmap.org/reverse',
    userAgent: process.env.GEO_USER_AGENT || 'webbogor-verificator/1.0',
    timeoutMs: angka(process.env.GEO_TIMEOUT_MS, 8000),
    jedaMs: angka(process.env.GEO_JEDA_MS, 1100)
  },

  // Kualitas gambar (Lapisan 2) — nilai default hasil kalibrasi 166 foto asli:
  // ketajaman min 311/p05 595 · luminans min 42/median 79 · sisi min 563
  kualitas: {
    minSisi: angka(process.env.VERIFY_MIN_SISI, 500),      // sisi terkecil (px)
    blurMin: angka(process.env.VERIFY_BLUR_MIN, 120),      // varians Laplacian minimal
    lumaGelap: angka(process.env.VERIFY_GELAP_MAKS, 35)    // luminans rata-rata maks dianggap gelap
  },

  // Lapisan 2 — parameter ELA
  ela: {
    kualitasUlang: 90,
    ambangAnomali: 0.40,     // fraksi blok anomali yang dianggap tinggi
    skorMinimalTinggi: 8,
    blokRasio: angka(process.env.VERIFY_BLOK_RASIO, 2.2), // hasil kalibrasi 166 foto: median 1.72, p90 2.05
    blokMinimal: angka(process.env.VERIFY_BLOK_MIN, 30)    // minimal blok datar agar dinilai
  },

  // Lapisan 3 — penyedia AI
  ai: {
    provider: (process.env.AI_PROVIDER || 'mock').toLowerCase(),
    model: process.env.AI_MODEL || '',
    apiKey: process.env.AI_API_KEY || '',
    timeoutMs: angka(process.env.AI_TIMEOUT_MS, 30000),
    retry: angka(process.env.AI_RETRY, 4),
    maxConcurrent: 3,
    maxSisi: 1600
  },

  // Verifikasi otomatis saat upload (mati secara bawaan)
  autoVerify: process.env.VERIFY_AUTO === '1',

  // Simpan heatmap ELA (opsional, mati secara bawaan)
  heatmap: process.env.VERIFY_HEATMAP === '1'
};

module.exports = cfg;
