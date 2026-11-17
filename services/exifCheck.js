// Lapisan 1 — pembacaan & pemeriksaan metadata EXIF (software, GPS, tanggal, kamera).
const exifr = require('exifr');
const cfg = require('./verifyConfig');
const { haversineMeter } = require('./hashCheck');

async function bacaExif(filePath) {
  try {
    const p = await exifr.parse(filePath, { gps: true, ifd0: true, exif: true, tiff: true, translateValues: true });
    return p || null;
  } catch (e) {
    return null;
  }
}

function tanggalLokal(d) {
  if (!(d instanceof Date) || isNaN(d.getTime())) return null;
  const p = n => String(n).padStart(2, '0');
  return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}`;
}

function tanggalDariExif(exif) {
  if (!exif) return null;
  const kandidat = [exif.DateTimeOriginal, exif.CreateDate, exif.ModifyDate, exif.DateTime];
  for (const k of kandidat) {
    const s = tanggalLokal(k instanceof Date ? k : (k ? new Date(k) : null));
    if (s) return s;
  }
  return null;
}

function softwareDariExif(exif) {
  if (!exif) return null;
  const v = exif.Software || exif.CreatorTool || null;
  return v ? String(v) : null;
}

function kameraDariExif(exif) {
  if (!exif) return null;
  const make = exif.Make ? String(exif.Make).trim() : '';
  const model = exif.Model ? String(exif.Model).trim() : '';
  const gabung = (make + ' ' + model).trim();
  return gabung || null;
}

function softwareEditor(software) {
  if (!software) return false;
  const s = String(software).toLowerCase();
  return cfg.softwareEditor.some(kunci => s.includes(kunci));
}

function angkaRational(v) {
  if (typeof v === 'number') return v;
  if (Array.isArray(v)) return Number(v.join(' '));
  return Number(v);
}

// Acuan dari lokasi.detail_tambahan (JSON) — bersifat opsional & netral bila tidak ada
function acuanDariDetail(detailTambahan) {
  const acuan = { lat: null, lng: null, mulai: null, akhir: null };
  if (!detailTambahan) return acuan;
  let d = detailTambahan;
  if (typeof d === 'string') {
    try { d = JSON.parse(d); } catch (e) { return acuan; }
  }
  if (!d || typeof d !== 'object') return acuan;

  const ambil = (...kunci) => {
    for (const k of kunci) if (d[k] !== undefined && d[k] !== null && d[k] !== '') return d[k];
    return null;
  };
  const lat = angkaRational(ambil('lat', 'latitude', 'gps_lat', 'lintang'));
  const lng = angkaRational(ambil('lng', 'lon', 'long', 'longitude', 'gps_lng', 'bujur'));
  acuan.lat = Number.isFinite(lat) ? lat : null;
  acuan.lng = Number.isFinite(lng) ? lng : null;

  const mulai = ambil('tanggal_mulai', 'tgl_mulai', 'mulai', 'tanggal', 'start_date');
  const akhir = ambil('tanggal_selesai', 'tgl_selesai', 'akhir', 'selesai', 'end_date');
  if (mulai) acuan.mulai = tanggalLokal(mulai instanceof Date ? mulai : new Date(mulai)) || String(mulai).slice(0, 10);
  if (akhir) acuan.akhir = tanggalLokal(akhir instanceof Date ? akhir : new Date(akhir)) || String(akhir).slice(0, 10);
  return acuan;
}

// Periksa EXIF satu foto → flag + rincian untuk Lapisan 1.
// @param acuan {lat, lng, mulai, akhir} — netral (tanpa penalti) bila field kosong
// @param dimensiFile {w, h} — ukuran asli file, untuk cek konsistensi dimensi EXIF
function periksa({ exif, acuan, flagLain, dimensiFile }) {
  const flags = [];
  const rincian = {
    ada: false,
    software: null,
    kamera: null,
    tanggal: null,
    lat: null,
    lng: null,
    jarakMeter: null,
    alasan: []
  };

  // EXIF "bermakna" = ada data kamera (tanggal/GPS/software/pabrikan).
  // Blok EXIF minimal buatan penerus gambar (hanya Orientation/ExifVersion)
  // tidak dianggap sebagai metadata asli.
  const bermakna = !!exif && !!(
    exif.Make || exif.Model || exif.Software || exif.CreatorTool ||
    exif.DateTimeOriginal || exif.CreateDate || exif.ModifyDate ||
    Number.isFinite(exif.latitude) || Number.isFinite(exif.longitude)
  );
  rincian.ada = bermakna;

  if (!bermakna) {
    flags.push('metadata_hilang');
    rincian.alasan.push('Metadata EXIF kamera (tanggal/GPS/software) tidak ditemukan pada berkas.');
    return { flags: flags.concat(flagLain || []), rincian };
  }

  rincian.software = softwareDariExif(exif);
  rincian.kamera = kameraDariExif(exif);
  rincian.tanggal = tanggalDariExif(exif);

  if (softwareEditor(rincian.software)) {
    flags.push('software_editor');
    rincian.alasan.push(`Software editor terdeteksi: ${rincian.software}.`);
  }

  // Konsistensi dimensi: EXIF menyimpan ukuran asli kamera. Upload men-simpan
  // ulang dengan resize (ukuran boleh beda), tetapi RASIO ASPEK harus sama.
  // Rasio berubah → gambar pernah dipotong/ditempel/diubah bentuk.
  if (dimensiFile && Number.isFinite(dimensiFile.w) && Number.isFinite(dimensiFile.h) && dimensiFile.w > 0 && dimensiFile.h > 0) {
    const ew = Number(exif.ExifImageWidth || exif.ImageWidth);
    const eh = Number(exif.ExifImageHeight || exif.ImageHeight);
    if (Number.isFinite(ew) && Number.isFinite(eh) && ew > 0 && eh > 0) {
      const aspekExif = ew / eh;
      const aspekFile = dimensiFile.w / dimensiFile.h;
      const beda = Math.abs(aspekExif / aspekFile - 1) > 0.02 && Math.abs(aspekFile / aspekExif - 1) > 0.02;
      if (beda) {
        flags.push('metadata_dismatch');
        rincian.alasan.push(`Rasio gambar pada EXIF ${ew}×${eh} tidak sesuai file ${dimensiFile.w}×${dimensiFile.h} — gambar pernah dipotong/diubah.`);
      }
    }
  }

  const gps = Number.isFinite(exif.latitude) && Number.isFinite(exif.longitude)
    ? { lat: exif.latitude, lng: exif.longitude }
    : null;
  if (gps) { rincian.lat = gps.lat; rincian.lng = gps.lng; }

  if (gps && acuan && Number.isFinite(acuan.lat) && Number.isFinite(acuan.lng)) {
    const jarak = haversineMeter(gps.lat, gps.lng, acuan.lat, acuan.lng);
    rincian.jarakMeter = Math.round(jarak * 10) / 10;
    if (jarak > cfg.toleransiGpsMeter) {
      flags.push('gps_jauh');
      rincian.alasan.push(`GPS ${rincian.jarakMeter} m dari titik lokasi (toleransi ${cfg.toleransiGpsMeter} m).`);
    }
  }

  if (rincian.tanggal && acuan && (acuan.mulai || acuan.akhir)) {
    const mulai = acuan.mulai, akhir = acuan.akhir;
    if ((mulai && rincian.tanggal < mulai) || (akhir && rincian.tanggal > akhir)) {
      flags.push('tanggal_luar');
      rincian.alasan.push(`Tanggal foto ${rincian.tanggal} di luar rentang ${mulai || '…'} s/d ${akhir || '…'}.`);
    }
  }

  return { flags: flags.concat(flagLain || []), rincian };
}

module.exports = { bacaExif, periksa, acuanDariDetail, softwareEditor, tanggalDariExif, kameraDariExif };
