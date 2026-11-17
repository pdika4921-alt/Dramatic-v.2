// Orkestrator verifikasi keaslian eviden — Lapisan 1..4 + batch + summary + override.
const fs = require('fs');
const path = require('path');
const cfg = require('./verifyConfig');
const { hitungHash, hamming } = require('./hashCheck');
const { bacaExif, periksa, acuanDariDetail } = require('./exifCheck');
const { ela } = require('./elaCheck');
const { analisisFoto } = require('./aiVision');
const { cekLokasi } = require('./geoCheck');
const { cekKualitas } = require('./qualityCheck');
const sharp = require('sharp');

const S = cfg.status;
const GAMBAR = /\.(jpe?g|png|webp|avif)$/i;

/* ── Skema tabel ─────────────────────────────────────────────────────────── */
async function initTabel(db) {
  await db.exec(`CREATE TABLE IF NOT EXISTS verifikasi_eviden (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    foto_id         INTEGER NOT NULL UNIQUE,
    lokasi_id       INTEGER NOT NULL,
    checklist       TEXT,
    sha256          TEXT,
    dhash           TEXT,
    phash           TEXT,
    exif_json       TEXT,
    ela_json        TEXT,
    ela_score       REAL,
    gps_lat         REAL,
    gps_lng         REAL,
    jarak_meter     REAL,
    software        TEXT,
    ai_json         TEXT,
    flags_json      TEXT,
    skor            REAL,
    status          TEXT DEFAULT 'belum',
    alasan          TEXT,
    diverifikasi_oleh TEXT,
    diverifikasi_at TEXT DEFAULT (datetime('now','localtime')),
    override_status TEXT,
    override_catatan TEXT,
    updated_at      TEXT DEFAULT (datetime('now','localtime'))
  )`);
  await db.exec('CREATE INDEX IF NOT EXISTS idx_verif_lokasi ON verifikasi_eviden(lokasi_id)');
  await db.exec('CREATE INDEX IF NOT EXISTS idx_verif_sha ON verifikasi_eviden(sha256)');
  console.log(' Tabel verifikasi_eviden siap.');
}

/* ── Utilitas ────────────────────────────────────────────────────────────── */
function clamp(n, min, max) { return Math.max(min, Math.min(max, n)); }

function pathFoto(lokasi, berkas) {
  const ntc = (lokasi.tematik || 'Umum').replace(/[^a-zA-Z0-9_-]/g, '_');
  const nlc = String(lokasi.nama_lokasi).replace(/[^a-zA-Z0-9_-]/g, '_');
  const kc = String(berkas.kategori).replace(/[^a-zA-Z0-9_-]/g, '_');
  return path.join(__dirname, '..', 'uploads', ntc, nlc, kc, berkas.file_name);
}

async function ambilFoto(db, fotoId) {
  const berkas = await db.get('SELECT * FROM berkas WHERE id = ?', [fotoId]);
  if (!berkas) return null;
  const lokasi = await db.get('SELECT * FROM lokasi WHERE id = ?', [berkas.lokasi_id]);
  if (!lokasi) return null;
  return { berkas, lokasi, filePath: pathFoto(lokasi, berkas) };
}

/* ── Indeks hash SEMUA foto yang pernah di-input (untuk deteksi duplikat) ── */
const hashFoto = new Map();     // fotoId -> {sha256, dhash, phash, lokasiId}
const sedangHitung = new Map(); // fotoId -> Promise (hindari hitung ganda)

async function indexHashFoto(db) {
  const rows = await db.all(
    `SELECT b.id, b.lokasi_id, b.kategori, b.file_name, l.tematik, l.nama_lokasi
     FROM berkas b JOIN lokasi l ON l.id = b.lokasi_id`
  );
  await Promise.all(rows.filter(r => GAMBAR.test(r.file_name || '')).map(r => {
    if (hashFoto.has(r.id)) return null;
    if (sedangHitung.has(r.id)) return sedangHitung.get(r.id);
    const p = (async () => {
      try {
        const pth = pathFoto({ tematik: r.tematik, nama_lokasi: r.nama_lokasi }, r);
        if (!fs.existsSync(pth)) return;
        const h = await hitungHash(pth);
        hashFoto.set(r.id, { sha256: h.sha256, dhash: h.dhash, phash: h.phash, lokasiId: r.lokasi_id });
      } catch (e) { /* file gagal dibaca — lewati */ } finally { sedangHitung.delete(r.id); }
    })();
    sedangHitung.set(r.id, p);
    return p;
  }));
  return hashFoto;
}

// Lupakan hash foto yang dihapus (dipanggil route DELETE berkas)
function lupakanHash(fotoId) {
  const id = Number(fotoId);
  hashFoto.delete(id);
  sedangHitung.delete(id);
}

/* ── Lapisan 4 — skor & status akhir ─────────────────────────────────────── */
function skorAkhir(flags, skorAi) {
  const kurang = kode => (cfg.flag[kode] ? cfg.flag[kode].bobot : 0);
  const l1 = clamp(100 - cfg.flagLapis1.reduce((a, k) => a + (flags.includes(k) ? kurang(k) : 0), 0), 0, 100);
  const l2 = clamp(100 - cfg.flagLapis2.reduce((a, k) => a + (flags.includes(k) ? kurang(k) : 0), 0), 0, 100);
  const total = cfg.bobot.ai * skorAi + cfg.bobot.lapis1 * l1 + cfg.bobot.lapis2 * l2;
  return { skor: Math.round(total), l1, l2 };
}

function statusAkhir(skor, flags, adaAi) {
  if (!adaAi) return S.BELUM;
  const berat = flags.some(k => cfg.flag[k] && cfg.flag[k].berat);
  if (berat || skor < cfg.ambang.curiga) return S.CURIGA;
  if (skor >= cfg.ambang.valid && flags.length === 0) return S.VALID;
  return S.CEK;
}

function flagDariAi(hasil) {
  const f = [];
  if (!hasil) return f;
  if (hasil.indikasi_editan === 'ya') f.push('editan_tinggi');
  else if (hasil.indikasi_editan === 'kemungkinan') f.push('kualitas_ragu');
  if (hasil.latar_belakang === 'mencurigakan') f.push('latar_mencurigakan');
  if (hasil.isi_sesuai === false) f.push('isi_tidak_sesuai');
  const jenis = String(hasil.jenis_foto || '').toLowerCase();
  if (/screenshot|tangkapan layar|foto layar/.test(jenis)) f.push('foto_layar');
  if (hasil.stempel_terbaca === false) f.push('stempel_tak_terbaca');
  return f;
}

/* ── Variasi kamera dalam satu kategori (butuh pemindaian saudara) ───────── */
async function variasiKamera(db, lokasiId, kategori, kecualiFotoId) {
  try {
    const saudara = await db.all(
      'SELECT id FROM berkas WHERE lokasi_id = ? AND kategori = ?',
      [lokasiId, kategori]
    );
    const kamera = new Set();
    for (const s of saudara) {
      if (s.id === kecualiFotoId) continue;
      const info = await ambilFoto(db, s.id);
      if (!info || !GAMBAR.test(info.berkas.file_name)) continue;
      const exif = await bacaExif(info.filePath);
      const k = exif && (exif.Make || exif.Model)
        ? String((exif.Make || '') + ' ' + (exif.Model || '')).trim().toLowerCase()
        : null;
      if (k) kamera.add(k);
    }
    return kamera.size >= 2;
  } catch (e) {
    return false;
  }
}

/* ── Verifikasi satu foto ────────────────────────────────────────────────── */
async function verifikasiFoto(db, fotoId, opsi) {
  const o = Object.assign({ user: null, force: false, variasi: null, acuan: null }, opsi || {});
  const info = await ambilFoto(db, fotoId);
  if (!info) throw Object.assign(new Error('Foto tidak ditemukan.'), { code: 'NOT_FOUND' });
  const { berkas, lokasi, filePath } = info;

  if (!GAMBAR.test(berkas.file_name)) throw Object.assign(new Error('Berkas bukan gambar — verifikasi hanya untuk foto.'), { code: 'BUKAN_GAMBAR' });
  if (!fs.existsSync(filePath)) throw Object.assign(new Error('File foto tidak ada di server.'), { code: 'FILE_HILANG' });

  const acuan = o.acuan || acuanDariDetail(lokasi.detail_tambahan);

  /* Lapisan 1 — integritas & metadata */
  const hash = await hitungHash(filePath);
  const exif = await bacaExif(filePath);
  let dimensiFile = null;
  try {
    const m = await sharp(await fs.promises.readFile(filePath)).metadata();
    if (m.width && m.height) dimensiFile = { w: m.width, h: m.height };
  } catch (e) { /* dimensi tidak terbaca — cek EXIF dilewati */ }
  const variasi = o.variasi === null ? await variasiKamera(db, berkas.lokasi_id, berkas.kategori, fotoId) : !!o.variasi;

  // Duplikat: SHA-256 identik ATAU jarak hamming ≤ ambang ("kemiripan 99%")
  // terhadap SEMUA foto yang pernah di-input (bukan hanya yang pernah diverifikasi)
  let duplikat = null;
  let duplikatDekat = null;
  await indexHashFoto(db);
  const maksHamming = cfg.duplikat.maksHamming;
  for (const [id, k] of hashFoto) {
    if (id === fotoId) continue;
    const samaSha = k.sha256 && k.sha256 === hash.sha256;
    const jarakP = k.phash ? hamming(k.phash, hash.phash) : 99;
    const jarakD = k.dhash ? hamming(k.dhash, hash.dhash) : 99;
    const jarak = Math.min(jarakP, jarakD);
    if (samaSha) {
      duplikat = { fotoId: id, lokasiId: k.lokasiId, cara: 'sha256', hamming: 0, kemiripan: 100 };
      break;
    }
    if (!duplikatDekat && jarak <= maksHamming) {
      duplikatDekat = {
        fotoId: id,
        lokasiId: k.lokasiId,
        cara: jarakP <= jarakD ? 'phash' : 'dhash',
        hamming: jarak,
        kemiripan: Math.round(((64 - jarak) / 64) * 1000) / 10
      };
    }
  }
  if (!duplikat) duplikat = duplikatDekat;

  const flagAwal = [];
  if (duplikat) flagAwal.push('duplikat_lokasi_lain');
  const cek1 = periksa({ exif, acuan, dimensiFile, flagLain: flagAwal.concat(variasi ? ['kamera_bervariasi'] : []) });

  /* Lapisan 1b — tag lokasi: reverse geocoding GPS vs nama lokasi */
  let geo = null;
  try {
    geo = await cekLokasi(db, { lat: cek1.rincian.lat, lng: cek1.rincian.lng, namaLokasi: lokasi.nama_lokasi });
  } catch (e) {
    geo = { hasil: 'netral', alamat: null, tokenWilayah: [], tokenCocok: [], teks: 'Geocoding gagal: ' + (e.message || e) };
  }
  if (geo.hasil === 'tidak_cocok') cek1.flags.push('lokasi_tidak_cocok');

  /* Lapisan 2 — ELA + kualitas gambar */
  const hasilEla = await ela(filePath);
  const kualitas = await cekKualitas(filePath, berkas.file_name);
  hasilEla.flags = Array.from(new Set((hasilEla.flags || []).concat(kualitas.flags || [])));
  hasilEla.alasan = (hasilEla.alasan || []).concat(kualitas.alasan || []);
  hasilEla.kualitas = kualitas.detail;

  /* Lapisan 3 — AI (cache per SHA-256) */
  let ai = null;
  let aiError = null;
  const cache = o.force ? null : await db.get(
    'SELECT ai_json FROM verifikasi_eviden WHERE sha256 = ? AND ai_json IS NOT NULL AND foto_id != ? LIMIT 1',
    [hash.sha256, fotoId]
  ) || await db.get(
    'SELECT ai_json FROM verifikasi_eviden WHERE sha256 = ? AND ai_json IS NOT NULL AND foto_id = ? LIMIT 1',
    [hash.sha256, fotoId]
  );
  if (cache && cache.ai_json) {
    try {
      const tersimpan = JSON.parse(cache.ai_json);
      // Hanya hasil AI yang benar-benar sukses yang boleh dipakai ulang.
      // Hasil lama tanpa skema baru (latar_belakang/isi_sesuai) dianggap basi:
      // ditolak dari cache sehingga dianalisis ulang sekali dengan prompt baru.
      const skemaBaru = tersimpan
        && Object.prototype.hasOwnProperty.call(tersimpan, 'isi_sesuai')
        && Object.prototype.hasOwnProperty.call(tersimpan, 'latar_belakang');
      if (skemaBaru && Number.isFinite(Number(tersimpan.skor_keaslian))) {
        ai = tersimpan;
        ai.cached = true;
      }
    } catch (e) { ai = null; }
  }
  if (!ai) {
    try {
      ai = await analisisFoto({
        filePath,
        namaFile: berkas.file_name,
        konteks: {
          kategori: berkas.kategori,
          subKategori: berkas.sub_kategori,
          namaLokasi: lokasi.nama_lokasi,
          tematik: lokasi.tematik
        }
      });
    } catch (e) {
      aiError = e.message || String(e);
    }
  }

  /* Lapisan 4 — skor & status */
  const flags = Array.from(new Set(cek1.flags.concat(flagDariAi(ai)).concat(hasilEla.flags)));
  const adaAi = !!ai;
  const { skor } = skorAkhir(flags, ai ? ai.skor_keaslian : 0);
  const status = statusAkhir(skor, flags, adaAi);

  const alasan = []
    .concat(cek1.rincian.alasan || [])
    .concat(geo && geo.hasil !== 'cocok' && geo.teks ? ['Lokasi: ' + geo.teks] : [])
    .concat(hasilEla.alasan)
    .concat(ai && ai.alasan_singkat ? ['AI: ' + ai.alasan_singkat] : [])
    .concat(aiError ? ['AI gagal: ' + aiError] : []);
  if (adaAi && ai.cached) alasan.push('Hasil AI dipakai ulang dari berkas identik (cache SHA-256).');

  const lapis1 = {
    ada_exif: cek1.rincian.ada,
    software: cek1.rincian.software,
    kamera: cek1.rincian.kamera,
    tanggal: cek1.rincian.tanggal,
    lat: cek1.rincian.lat,
    lng: cek1.rincian.lng,
    jarak_meter: cek1.rincian.jarakMeter,
    lokasi: geo ? {
      hasil: geo.hasil,
      alamat: geo.alamat || null,
      token_wilayah: geo.tokenWilayah || [],
      token_cocok: geo.tokenCocok || [],
      teks: geo.teks || ''
    } : null,
    alasan: cek1.rincian.alasan
  };
  const lapis2 = {
    skor: hasilEla.skor,
    anomali: Math.round(hasilEla.anomali * 1000) / 1000,
    alasan: hasilEla.alasan,
    heatmap: hasilEla.heatmapPath,
    blok: hasilEla.blok || null,
    kualitas: hasilEla.kualitas || null
  };

  const duplikatInfo = duplikat
    ? {
        foto_id: duplikat.fotoId,
        lokasi_id: duplikat.lokasiId,
        cara: duplikat.cara,
        hamming: duplikat.hamming,
        kemiripan: duplikat.kemiripan,
        di_lokasi_lain: duplikat.lokasiId !== berkas.lokasi_id
      }
    : null;

  // Hasil AI dinormalkan sebelum disimpan (dipakai ulang oleh modal detail)
  if (ai) {
    ai.provider = cfg.ai.provider;
    ai.jenis_foto = ai.jenis_foto || '';
    ai.alasan_singkat = ai.alasan_singkat || '';
    ai.detail = ai.detail || '';
    if (ai.stempel_terbaca === undefined) ai.stempel_terbaca = null;
    if (!ai.indikasi_editan) ai.indikasi_editan = null;
    if (!['wajar', 'mencurigakan'].includes(ai.latar_belakang)) ai.latar_belakang = null;
    if (typeof ai.isi_sesuai !== 'boolean') ai.isi_sesuai = null;
  } else {
    ai = { error: aiError || 'tanpa hasil', provider: cfg.ai.provider };
  }

  const baris = {
    foto_id: fotoId,
    lokasi_id: berkas.lokasi_id,
    checklist: berkas.kategori,
    sha256: hash.sha256,
    dhash: hash.dhash,
    phash: hash.phash,
    exif_json: JSON.stringify({ lapis1, duplikat: duplikatInfo }),
    ela_json: JSON.stringify(lapis2),
    ela_score: hasilEla.skor,
    gps_lat: lapis1.lat,
    gps_lng: lapis1.lng,
    jarak_meter: lapis1.jarak_meter,
    software: lapis1.software,
    ai_json: JSON.stringify(ai),
    flags_json: JSON.stringify(flags),
    skor: skor,
    status: status,
    alasan: alasan.join(' | '),
    diverifikasi_oleh: o.user || null
  };

  await db.run(
    `INSERT INTO verifikasi_eviden
      (foto_id, lokasi_id, checklist, sha256, dhash, phash, exif_json, ela_json, ela_score,
       gps_lat, gps_lng, jarak_meter, software, ai_json, flags_json, skor, status, alasan,
       diverifikasi_oleh, diverifikasi_at, updated_at)
     VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?, datetime('now','localtime'), datetime('now','localtime'))
     ON CONFLICT(foto_id) DO UPDATE SET
       lokasi_id=excluded.lokasi_id, checklist=excluded.checklist, sha256=excluded.sha256,
       dhash=excluded.dhash, phash=excluded.phash, exif_json=excluded.exif_json,
       ela_json=excluded.ela_json, ela_score=excluded.ela_score, gps_lat=excluded.gps_lat,
       gps_lng=excluded.gps_lng, jarak_meter=excluded.jarak_meter, software=excluded.software,
       ai_json=excluded.ai_json, flags_json=excluded.flags_json, skor=excluded.skor,
       status=excluded.status, alasan=excluded.alasan, diverifikasi_oleh=excluded.diverifikasi_oleh,
       diverifikasi_at=datetime('now','localtime'), updated_at=datetime('now','localtime')`,
    [baris.foto_id, baris.lokasi_id, baris.checklist, baris.sha256, baris.dhash, baris.phash,
     baris.exif_json, baris.ela_json, baris.ela_score, baris.gps_lat, baris.gps_lng,
     baris.jarak_meter, baris.software, baris.ai_json, baris.flags_json, baris.skor,
     baris.status, baris.alasan, baris.diverifikasi_oleh]
  );

  return detailVerifikasi(db, fotoId);
}

/* ── Detail untuk UI ─────────────────────────────────────────────────────── */
function parseJsonSafe(s, bawaan) {
  try { return JSON.parse(s); } catch (e) { return bawaan; }
}

async function detailVerifikasi(db, fotoId) {
  const v = await db.get('SELECT * FROM verifikasi_eviden WHERE foto_id = ?', [fotoId]);
  if (!v) return null;
  const info = await ambilFoto(db, fotoId);
  const exifJson = parseJsonSafe(v.exif_json, {});
  const elaJson = parseJsonSafe(v.ela_json, {});
  const flags = parseJsonSafe(v.flags_json, []);
  let ai = parseJsonSafe(v.ai_json, null);
  if (ai && ai.error !== undefined && ai.skor_keaslian === undefined) ai = { error: ai.error, provider: ai.provider };
  return {
    foto_id: v.foto_id,
    lokasi_id: v.lokasi_id,
    checklist: v.checklist,
    file_name: info ? info.berkas.file_name : null,
    url: info ? `/uploads/${(info.lokasi.tematik || 'Umum').replace(/[^a-zA-Z0-9_-]/g, '_')}/${String(info.lokasi.nama_lokasi).replace(/[^a-zA-Z0-9_-]/g, '_')}/${String(info.berkas.kategori).replace(/[^a-zA-Z0-9_-]/g, '_')}/${info.berkas.file_name}` : null,
    status: v.status,
    status_efektif: v.override_status || v.status,
    override_status: v.override_status,
    override_catatan: v.override_catatan,
    skor: v.skor,
    alasan: v.alasan ? v.alasan.split(' | ') : [],
    flags: flags.map(k => ({ kode: k, label: (cfg.flag[k] || {}).label || k, berat: !!(cfg.flag[k] || {}).berat })),
    hash: { sha256: v.sha256, dhash: v.dhash, phash: v.phash },
    lapis1: exifJson.lapis1 || null,
    duplikat: exifJson.duplikat || null,
    lapis2: elaJson,
    lapis3: ai,
    diverifikasi_oleh: v.diverifikasi_oleh,
    diverifikasi_at: v.diverifikasi_at
  };
}

/* ── Ringkasan 7 check per foto untuk UI galeri (strip alasan + popover) ── */
const DAFTAR_CHECK = [
  ['duplikat', '1. Duplikat 99%'],
  ['lokasi', '2. Tag lokasi (GPS)'],
  ['exif', '3. Metadata EXIF'],
  ['kualitas', '4. Kualitas gambar'],
  ['forensik', '5. Forensik piksel (ELA)'],
  ['isi', '6. Isi vs kategori'],
  ['ai', '7. AI keaslian']
];

function ringkasanVerifikasi(v) {
  const exifJson = parseJsonSafe(v.exif_json, {});
  const elaJson = parseJsonSafe(v.ela_json, {});
  const flags = parseJsonSafe(v.flags_json, []);
  const ai = parseJsonSafe(v.ai_json, null) || {};
  const l1 = exifJson.lapis1 || {};
  const dup = exifJson.duplikat || null;
  const l2 = elaJson || {};
  const punya = k => flags.includes(k);
  const labelFlag = k => (cfg.flag[k] || {}).label || k;
  const adaHasil = !!v.sha256;
  const checks = [];

  const belumCek = (kode, nama, alasan) => ({ kode, nama, hasil: 'netral', teks: alasan || 'Belum dicek — jalankan verifikasi.' });

  if (!adaHasil) {
    for (const [kode, nama] of DAFTAR_CHECK) checks.push(belumCek(kode, nama, 'Belum diverifikasi.'));
  } else {
    /* 1. Duplikat */
    checks.push(dup
      ? {
          kode: 'duplikat', nama: '1. Duplikat 99%', hasil: 'tolak',
          teks: `Mirip ${dup.kemiripan != null ? dup.kemiripan + '%' : '≥99%'} dengan foto #${dup.foto_id}${dup.di_lokasi_lain ? ' di lokasi lain' : ' lain di titik ini'} (${dup.cara}).`
        }
      : { kode: 'duplikat', nama: '1. Duplikat 99%', hasil: 'ok', teks: 'Foto unik — tidak ada foto lain dengan kemiripan ≥99%.' });

    /* 2. Tag lokasi (GPS) */
    const lok = l1.lokasi;
    if (lok && lok.hasil) {
      checks.push({
        kode: 'lokasi', nama: '2. Tag lokasi (GPS)',
        hasil: lok.hasil === 'cocok' ? 'ok' : lok.hasil === 'tidak_cocok' ? 'tolak' : 'netral',
        teks: lok.teks || '—'
      });
    } else if (Number.isFinite(l1.lat)) {
      checks.push({ kode: 'lokasi', nama: '2. Tag lokasi (GPS)', hasil: 'netral', teks: `GPS ${Number(l1.lat).toFixed(5)}, ${Number(l1.lng).toFixed(5)} — hasil lama, jalankan verifikasi ulang untuk cek kecocokan wilayah.` });
    } else {
      checks.push({ kode: 'lokasi', nama: '2. Tag lokasi (GPS)', hasil: 'netral', teks: 'Tidak ada GPS pada foto — verifikasi lokasi dinilai netral.' });
    }

    /* 3. Metadata EXIF */
    if (punya('software_editor')) {
      checks.push({ kode: 'exif', nama: '3. Metadata EXIF', hasil: 'tolak', teks: `Software editor terdeteksi: ${l1.software || '?'}.` });
    } else if (punya('metadata_dismatch')) {
      checks.push({ kode: 'exif', nama: '3. Metadata EXIF', hasil: 'tolak', teks: labelFlag('metadata_dismatch') + '.' });
    } else if (!l1.ada_exif) {
      checks.push({ kode: 'exif', nama: '3. Metadata EXIF', hasil: 'warn', teks: 'Metadata EXIF tidak ada pada berkas.' });
    } else {
      const dasar = [l1.kamera, l1.tanggal].filter(Boolean).join(' · ') || 'EXIF ada';
      const catatan = [];
      if (punya('tanggal_luar')) catatan.push('tanggal di luar masa kerja');
      if (punya('kamera_bervariasi')) catatan.push('kamera berbeda dalam satu kategori');
      checks.push({
        kode: 'exif', nama: '3. Metadata EXIF',
        hasil: catatan.length ? 'warn' : 'ok',
        teks: dasar + ' · tanpa software editor' + (catatan.length ? ' — ' + catatan.join(', ') : '')
      });
    }

    /* 4. Kualitas gambar */
    const kq = l2.kualitas;
    if (!kq) {
      checks.push(belumCek('kualitas', '4. Kualitas gambar', 'Belum dicek — jalankan verifikasi ulang.'));
    } else if (punya('format_tidak_sesuai')) {
      checks.push({ kode: 'kualitas', nama: '4. Kualitas gambar', hasil: 'tolak', teks: labelFlag('format_tidak_sesuai') + '.' });
    } else {
      const temuan = (cfg.checkFlag.kualitas || []).filter(punya);
      const ukuran = `${kq.ketajaman != null ? 'ketajaman ' + kq.ketajaman : ''}${kq.luminans != null ? ' · luminans ' + kq.luminans : ''}${kq.lebar ? ` · ${kq.lebar}×${kq.tinggi} px` : ''}`;
      checks.push(temuan.length
        ? { kode: 'kualitas', nama: '4. Kualitas gambar', hasil: 'warn', teks: temuan.map(labelFlag).join('; ') + ` (${ukuran}).` }
        : { kode: 'kualitas', nama: '4. Kualitas gambar', hasil: 'ok', teks: `Tajam · terang · ${ukuran}${kq.jpegKualitas != null ? ' · JPEG q±' + kq.jpegKualitas : ''}.` });
    }

    /* 5. Forensik piksel (ELA) */
    if (l2.skor === undefined || l2.skor === null) {
      checks.push(belumCek('forensik', '5. Forensik piksel (ELA)'));
    } else {
      const anomali = l2.anomali != null ? Math.round(l2.anomali * 100) : 0;
      if (punya('ela_tinggi')) {
        checks.push({ kode: 'forensik', nama: '5. Forensik piksel (ELA)', hasil: 'tolak', teks: `Anomali ELA ${anomali}% dari area foto — ada area dicurigai diedit.` });
      } else if (punya('blok_artefak')) {
        checks.push({ kode: 'forensik', nama: '5. Forensik piksel (ELA)', hasil: 'warn', teks: `Anomali ELA ${anomali}% · artefak blok 8×8 terdeteksi — bekas kompresi ulang/tempelan.` });
      } else {
        checks.push({ kode: 'forensik', nama: '5. Forensik piksel (ELA)', hasil: 'ok', teks: `Anomali ELA ${anomali}% — normal.` });
      }
    }

    /* 6. Isi vs kategori */
    if (punya('isi_tidak_sesuai')) {
      checks.push({ kode: 'isi', nama: '6. Isi vs kategori', hasil: 'tolak', teks: `AI menilai isi foto tidak sesuai kategori "${v.checklist}".` });
    } else if (Number.isFinite(Number(ai.skor_keaslian))) {
      checks.push(ai.isi_sesuai === true
        ? { kode: 'isi', nama: '6. Isi vs kategori', hasil: 'ok', teks: `AI: isi foto sesuai kategori "${v.checklist}".` }
        : { kode: 'isi', nama: '6. Isi vs kategori', hasil: 'netral', teks: 'AI tidak menilai kesesuaian isi (hasil lama) — jalankan verifikasi ulang.' });
    } else {
      checks.push({ kode: 'isi', nama: '6. Isi vs kategori', hasil: 'netral', teks: 'Belum ada hasil AI.' });
    }

    /* 7. AI keaslian */
    if (!Number.isFinite(Number(ai.skor_keaslian))) {
      checks.push({ kode: 'ai', nama: '7. AI keaslian', hasil: 'netral', teks: ai.error ? `AI gagal: ${ai.error}` : 'Belum ada hasil AI.' });
    } else {
      const skorAi = Number(ai.skor_keaslian);
      let hasil7 = 'ok';
      if (punya('editan_tinggi') || punya('foto_layar') || skorAi < 50) hasil7 = 'tolak';
      else if (punya('latar_mencurigakan') || punya('kualitas_ragu') || skorAi < 80) hasil7 = 'warn';
      let teks7 = `Skor AI ${skorAi}/100 · indikasi editan: ${ai.indikasi_editan || '—'} · latar: ${ai.latar_belakang || '—'}`;
      if (ai.alasan_singkat) teks7 += ` — ${ai.alasan_singkat}`;
      checks.push({ kode: 'ai', nama: '7. AI keaslian', hasil: hasil7, teks: teks7 });
    }
  }

  /* Kesimpulan & strip 1 baris */
  const stAkhir = v.override_status || v.status || 'belum';
  const berat = flags.filter(k => cfg.flag[k] && cfg.flag[k].berat);
  const labelList = daftar => {
    const potong = daftar.slice(0, 3).map(labelFlag);
    return potong.join('; ') + (daftar.length > 3 ? `; +${daftar.length - 3} temuan lain` : '');
  };
  let kesimpulan;
  if (stAkhir === 'belum') {
    kesimpulan = 'Belum diverifikasi — jalankan Verifikasi AI.';
  } else if (stAkhir === 'valid') {
    kesimpulan = `Lolos — skor ${Math.round(v.skor || 0)}, tidak ada temuan.`;
  } else if (flags.length) {
    kesimpulan = `${stAkhir === 'curiga' ? 'Tidak lolos' : 'Perlu dicek'} — skor ${Math.round(v.skor || 0)} · temuan${berat.length ? ' berat' : ''}: ${labelList(berat.length ? berat : flags)}.`;
  } else {
    kesimpulan = `${stAkhir === 'curiga' ? 'Tidak lolos' : 'Perlu dicek'} — skor ${Math.round(v.skor || 0)} di bawah ambang tanpa temuan flag.`;
  }

  const ikon = { valid: '✅', cek: '⚠️', curiga: '❌', belum: '⏳' }[stAkhir] || '⏳';
  let ringkas;
  if (stAkhir === 'belum') ringkas = '⏳ Belum diverifikasi';
  else if (stAkhir === 'valid') ringkas = `✅ Lolos — skor ${Math.round(v.skor || 0)}, tanpa temuan`;
  else {
    const utama = berat.length ? berat : flags;
    const teksUtama = utama.length ? labelFlag(utama[0]) : `skor ${Math.round(v.skor || 0)} di bawah ambang`;
    const sisa = utama.length - 1;
    ringkas = `${ikon} ${stAkhir === 'curiga' ? 'Tidak lolos' : 'Perlu dicek'} — ${teksUtama}${sisa > 0 ? ` +${sisa} temuan lain` : ''}`;
  }
  if (v.override_status) {
    ringkas = `🛡️ Override admin → ${stAkhir} — ${v.override_catatan || 'tanpa catatan'}`;
    kesimpulan = `Di-override admin menjadi "${stAkhir}" — ${v.override_catatan || 'tanpa catatan'}. Keputusan manual menggantikan hasil otomatis.`;
  }

  return {
    status: stAkhir,
    skor: v.skor != null ? Math.round(v.skor) : null,
    override_status: v.override_status || null,
    checks,
    kesimpulan,
    ringkas
  };
}

/* ── Summary per kategori ────────────────────────────────────────────────── */
async function summary(db, lokasiId) {
  const rows = await db.all(
    `SELECT b.kategori AS kategori, COUNT(*) AS total,
            SUM(CASE WHEN COALESCE(v.override_status, v.status, 'belum') = 'valid' THEN 1 ELSE 0 END) AS valid,
            SUM(CASE WHEN COALESCE(v.override_status, v.status, 'belum') = 'cek' THEN 1 ELSE 0 END) AS cek,
            SUM(CASE WHEN COALESCE(v.override_status, v.status, 'belum') = 'curiga' THEN 1 ELSE 0 END) AS curiga,
            SUM(CASE WHEN COALESCE(v.override_status, v.status, 'belum') = 'belum' THEN 1 ELSE 0 END) AS belum
     FROM berkas b
     LEFT JOIN verifikasi_eviden v ON v.foto_id = b.id
     WHERE b.lokasi_id = ?
     GROUP BY b.kategori`,
    [lokasiId]
  );
  const perKategori = {};
  const total = { valid: 0, cek: 0, curiga: 0, belum: 0, total: 0 };
  for (const r of rows) {
    const entri = {
      valid: Number(r.valid || 0),
      cek: Number(r.cek || 0),
      curiga: Number(r.curiga || 0),
      belum: Number(r.belum || 0),
      total: Number(r.total || 0)
    };
    perKategori[r.kategori] = entri;
    for (const k of ['valid', 'cek', 'curiga', 'belum', 'total']) total[k] += entri[k];
  }
  return { perKategori, total };
}

/* ── Batch (dengan progres) ──────────────────────────────────────────────── */
const batches = new Map();

function progresBatch(batchId) {
  const b = batches.get(batchId);
  if (!b) return null;
  return {
    batchId,
    total: b.total,
    selesai: b.selesai,
    status: b.status,
    pesan: b.pesan,
    persen: b.total ? Math.round((b.selesai / b.total) * 100) : 100
  };
}

async function denganAntrean(fungsi, maks) {
  // antrean sederhana: jalankan maksimal `maks` tugas bersamaan
  return new Promise((resolve, reject) => {
    antreanUmum.push({ fungsi, resolve, reject });
    jalankanUmum(maks);
  });
}
const antreanUmum = [];
let berjalanUmum = 0;
function jalankanUmum(maks) {
  while (berjalanUmum < maks && antreanUmum.length) {
    const t = antreanUmum.shift();
    berjalanUmum++;
    Promise.resolve()
      .then(() => t.fungsi())
      .then(t.resolve, t.reject)
      .finally(() => { berjalanUmum--; jalankanUmum(maks); });
  }
}

async function verifikasiBatch(db, lokasiId, opsi) {
  const o = Object.assign({ checklist: null, user: null, force: false, fotoIds: null }, opsi || {});
  const lokasi = await db.get('SELECT * FROM lokasi WHERE id = ?', [lokasiId]);
  if (!lokasi) throw Object.assign(new Error('Lokasi tidak ditemukan.'), { code: 'NOT_FOUND' });

  let daftar;
  if (Array.isArray(o.fotoIds) && o.fotoIds.length) {
    const placeholders = o.fotoIds.map(() => '?').join(',');
    daftar = await db.all(`SELECT * FROM berkas WHERE id IN (${placeholders}) AND lokasi_id = ?`, o.fotoIds.concat([lokasiId]));
  } else {
    daftar = await db.all('SELECT * FROM berkas WHERE lokasi_id = ? ORDER BY id', [lokasiId]);
  }
  daftar = daftar.filter(b => GAMBAR.test(b.file_name || ''));
  if (o.checklist) daftar = daftar.filter(b => b.kategori === o.checklist);

  const batchId = 'b' + Date.now().toString(36) + Math.random().toString(36).slice(2, 7);
  batches.set(batchId, { total: daftar.length, selesai: 0, status: daftar.length ? 'jalan' : 'selesai', pesan: '' });
  if (batches.size > 50) {
    for (const [k, v] of batches) if (v.status !== 'jalan' && batches.size > 50) batches.delete(k);
  }

  (async () => {
    const st = batches.get(batchId);
    try {
      // Pra-pindai variasi kamera per kategori
      const variasiPerKategori = {};
      const kelompok = {};
      for (const b of daftar) (kelompok[b.kategori] = kelompok[b.kategori] || []).push(b);
      for (const [kategori, anggota] of Object.entries(kelompok)) {
        const kamera = new Set();
        for (const b of anggota) {
          const p = pathFoto(lokasi, b);
          if (!fs.existsSync(p)) continue;
          const exif = await bacaExif(p);
          const k = exif && (exif.Make || exif.Model)
            ? String((exif.Make || '') + ' ' + (exif.Model || '')).trim().toLowerCase()
            : null;
          if (k) kamera.add(k);
        }
        variasiPerKategori[kategori] = kamera.size >= 2;
      }

      await Promise.all(daftar.map(b =>
        denganAntrean(async () => {
          try {
            await verifikasiFoto(db, b.id, {
              user: o.user,
              force: o.force,
              variasi: variasiPerKategori[b.kategori],
              acuan: acuanDariDetail(lokasi.detail_tambahan)
            });
          } catch (e) {
            st.pesan = `${b.file_name}: ${e.message || e}`;
          } finally {
            st.selesai++;
          }
        }, cfg.ai.maxConcurrent)
      ));
      st.status = 'selesai';
    } catch (e) {
      st.status = 'gagal';
      st.pesan = e.message || String(e);
    }
  })();

  return { batchId, total: daftar.length };
}

/* ── Override admin ──────────────────────────────────────────────────────── */
async function override(db, fotoId, opsi) {
  const o = Object.assign({ status: null, catatan: null, user: null }, opsi || {});
  if (!Object.values(S).includes(o.status)) {
    throw Object.assign(new Error('Status override tidak dikenal.'), { code: 'VALIDASI' });
  }
  const ada = await db.get('SELECT foto_id FROM verifikasi_eviden WHERE foto_id = ?', [fotoId]);
  if (!ada) throw Object.assign(new Error('Foto belum diverifikasi.'), { code: 'NOT_FOUND' });
  await db.run(
    `UPDATE verifikasi_eviden SET override_status = ?, override_catatan = ?,
       updated_at = datetime('now','localtime') WHERE foto_id = ?`,
    [o.status, o.catatan, fotoId]
  );
  return detailVerifikasi(db, fotoId);
}

/* ── Verifikasi otomatis saat upload ─────────────────────────────────────── */
async function verifikasiOtomatisUpload(db, lokasiId, fileNames, user) {
  if (!cfg.autoVerify || !fileNames || !fileNames.length) return null;
  try {
    const placeholders = fileNames.map(() => '?').join(',');
    const rows = await db.all(
      `SELECT id FROM berkas WHERE lokasi_id = ? AND file_name IN (${placeholders})`,
      [lokasiId].concat(fileNames)
    );
    if (!rows.length) return null;
    const { batchId } = await verifikasiBatch(db, lokasiId, {
      user: user,
      fotoIds: rows.map(r => r.id)
    });
    return batchId;
  } catch (e) {
    console.warn('Verifikasi otomatis upload:', e.message || e);
    return null;
  }
}

module.exports = {
  S,
  GAMBAR,
  initTabel,
  pathFoto,
  verifikasiFoto,
  verifikasiBatch,
  progresBatch,
  detailVerifikasi,
  ringkasanVerifikasi,
  summary,
  override,
  verifikasiOtomatisUpload,
  indexHashFoto,
  lupakanHash
};
