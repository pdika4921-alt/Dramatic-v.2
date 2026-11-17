// ─────────────────────────────────────────────────────────────────────────────
// MITRATEL — modul tambahan mandiri.
// Tidak mengubah alur dokumen/lokasi/eviden yang sudah ada.
// ─────────────────────────────────────────────────────────────────────────────
const fs = require('fs');
const path = require('path');
const os = require('os');
const { execFile } = require('child_process');
const multer = require('multer');
const db = require('../db-wrapper');
const rekapOpm = require('./rekapOpm');
// KML/KMZ → gambar peta (citra satelit Esri + overlay rute) — dipakai slot KMZ ATP.
const { bacaKml, parseKml, renderPeta } = require('../kml-renderer');

const ROOT = path.join(__dirname, '..');
const TEMPLATE_OTP = path.join(ROOT, 'templates', 'OTP MITRATEL TAMPLATE.xlsx');
const TEMPLATE_ATP = path.join(ROOT, 'templates', 'ATP TEMPLATE.xlsx');
const UPLOAD_ROOT = path.join(ROOT, 'uploads', 'MITRATEL');
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 30 * 1024 * 1024, files: 40 } });

function safeName(v) { return String(v || '').replace(/[^a-zA-Z0-9._-]/g, '_'); }
function q(v) { return "'" + String(v).replace(/'/g, "''") + "'"; }

// URL gambar yang boleh diambil server (dipakai saat drag gambar dari WhatsApp/Web
// yang tidak bisa di-fetch langsung oleh browser karena CORS). Tolak host lokal/privat.
function urlGambarAman(v) {
  try {
    const u = new URL(String(v || '').trim());
    if (!/^https?:$/.test(u.protocol)) return false;
    const h = u.hostname.toLowerCase();
    if (!h || h === 'localhost' || h.endsWith('.localhost') || h === '::1' || h === '[::1]') return false;
    if (h.startsWith('[')) return false; // literal IPv6 (termasuk ::1 / private)
    if (h.endsWith('.local') || h.endsWith('.internal')) return false;
    if (/^(127\.|0\.|10\.|169\.254\.|192\.168\.)/.test(h)) return false;
    if (/^172\.(1[6-9]|2\d|3[01])\./.test(h)) return false;
    if (h === 'metadata.google.internal') return false;
    return true;
  } catch (_) { return false; }
}
function ps(script, timeout = 240000) {
  return new Promise((resolve, reject) => {
    const tmpFile = path.join(os.tmpdir(), 'mitratel_ps_' + Date.now() + Math.random().toString(36).slice(2) + '.ps1');
    fs.writeFileSync(tmpFile, script, 'utf8');
    execFile('powershell.exe',
      ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', tmpFile],
      { timeout, maxBuffer: 16 * 1024 * 1024, windowsHide: true },
      (err, stdout, stderr) => {
        try { fs.rmSync(tmpFile, { force: true }); } catch (_) {}
        err ? reject(new Error(String(stderr || err.message).trim().slice(0, 1200))) : resolve(stdout);
      });
  });
}

async function siap() {
  await db.exec(`CREATE TABLE IF NOT EXISTS mitratel_data (
    lokasi_id INTEGER PRIMARY KEY,
    insert_new_ring TEXT DEFAULT '', id_span_existing TEXT DEFAULT '', nama_ring TEXT DEFAULT '', id_span TEXT DEFAULT '',
    id_site_ne TEXT DEFAULT '', nama_site_ne TEXT DEFAULT '', tikor_ne_lat TEXT DEFAULT '', tikor_ne_long TEXT DEFAULT '',
    id_site_fe TEXT DEFAULT '', nama_site_fe TEXT DEFAULT '', tikor_fe_lat TEXT DEFAULT '', tikor_fe_long TEXT DEFAULT '',
    panjang_kabel_n_to_n TEXT DEFAULT '', panjang_kabel_baru TEXT DEFAULT '', jumlah_core TEXT DEFAULT '2',
    tanggal_submit_atp TEXT DEFAULT '', optical_length_ne TEXT DEFAULT '', optical_length_fe TEXT DEFAULT '',
    updated_at TEXT DEFAULT CURRENT_TIMESTAMP
  )`);
  await db.exec(`CREATE TABLE IF NOT EXISTS mitratel_foto (
    id INTEGER PRIMARY KEY AUTOINCREMENT, lokasi_id INTEGER NOT NULL, slot TEXT NOT NULL, file_name TEXT NOT NULL,
    original_name TEXT DEFAULT '', hasil_scan TEXT DEFAULT '', created_at TEXT DEFAULT CURRENT_TIMESTAMP,
    UNIQUE(lokasi_id, slot)
  )`);
  try { await db.exec("ALTER TABLE mitratel_foto ADD COLUMN tag TEXT DEFAULT ''"); } catch (_) { /* sudah ada */ }
  await db.exec(`CREATE TABLE IF NOT EXISTS atp_foto (
    id INTEGER PRIMARY KEY AUTOINCREMENT, lokasi_id INTEGER NOT NULL, slot TEXT NOT NULL, file_name TEXT NOT NULL,
    original_name TEXT DEFAULT '', created_at TEXT DEFAULT CURRENT_TIMESTAMP,
    UNIQUE(lokasi_id, slot)
  )`);
  // Slot ABD (Lay Out) membawa SITE NAME + SITE ID per blok.
  try { await db.exec("ALTER TABLE atp_foto ADD COLUMN site_name TEXT DEFAULT ''"); } catch (_) { /* sudah ada */ }
  try { await db.exec("ALTER TABLE atp_foto ADD COLUMN site_id TEXT DEFAULT ''"); } catch (_) { /* sudah ada */ }
  try { await db.exec("ALTER TABLE atp_foto ADD COLUMN optical_length TEXT DEFAULT ''"); } catch (_) { /* sudah ada */ }
  try { await db.exec("ALTER TABLE atp_foto ADD COLUMN att_ab TEXT DEFAULT ''"); } catch (_) { /* sudah ada */ }
  try { await db.exec("ALTER TABLE atp_foto ADD COLUMN att_ba TEXT DEFAULT ''"); } catch (_) { /* sudah ada */ }
  // Nilai OPTICAL LENGTH + REDAMAN per core (1–14) untuk sheet "Test Result Core (FLP)".
  // Dipisah dari atp_foto supaya bisa diisi tanpa harus mengunggah foto dulu.
  await db.exec(`CREATE TABLE IF NOT EXISTS atp_meta (
    lokasi_id INTEGER NOT NULL, core INTEGER NOT NULL,
    optical_length TEXT DEFAULT '', att_ab TEXT DEFAULT '', att_ba TEXT DEFAULT '',
    updated_at TEXT DEFAULT CURRENT_TIMESTAMP,
    PRIMARY KEY (lokasi_id, core)
  )`);
  await sinkronFotoDariDisk();
  await sinkronAtpDariDisk();
}

function normalizeData(b = {}) {
  const fields = ['insert_new_ring','id_span_existing','nama_ring','id_span','id_site_ne','nama_site_ne','tikor_ne_lat','tikor_ne_long',
    'id_site_fe','nama_site_fe','tikor_fe_lat','tikor_fe_long','panjang_kabel_n_to_n','panjang_kabel_baru','jumlah_core','tanggal_submit_atp',
    'optical_length_ne','optical_length_fe'];
  const out = {};
  fields.forEach(k => out[k] = String(b[k] == null ? '' : b[k]).trim());
  if (!['2', '4'].includes(out.jumlah_core)) out.jumlah_core = '2';
  return out;
}

async function getData(lokasiId) {
  await siap();
  const lokasi = await db.get('SELECT id,nama_lokasi,pid,tematik FROM lokasi WHERE id=?', [lokasiId]);
  if (!lokasi) return null;
  const data = await db.get('SELECT * FROM mitratel_data WHERE lokasi_id=?', [lokasiId]) || { lokasi_id: Number(lokasiId), jumlah_core: '2' };
  const fotosRaw = await db.all('SELECT slot,file_name,original_name,hasil_scan,tag,created_at FROM mitratel_foto WHERE lokasi_id=? ORDER BY slot', [lokasiId]);
  // Selaraskan dengan isi folder: file yang barisnya hilang ikut disertakan,
  // dan rujukan file yang sudah tidak ada diganti berkas terbaru pada slot itu.
  const disk = fotoDariDisk(lokasiId);
  const folder = path.join(UPLOAD_ROOT, safeName(lokasiId));
  for (const f of fotosRaw) {
    if (!fs.existsSync(path.join(folder, f.file_name)) && disk[f.slot]) f.file_name = disk[f.slot].file_name;
  }
  for (const [slot, info] of Object.entries(disk)) {
    if (!fotosRaw.some(f => f.slot === slot)) {
      fotosRaw.push({ slot, file_name: info.file_name, original_name: '', hasil_scan: '', tag: '', created_at: '' });
    }
  }
  fotosRaw.sort((a, b) => String(a.slot).localeCompare(String(b.slot), 'en', { numeric: true }));
  const fotos = [];
  for (const f of fotosRaw) {
    let tag = String(f.tag || '');
    const prefix = String(f.slot || '').startsWith('ne_') ? 'NE' : (String(f.slot || '').startsWith('fe_') ? 'FE' : '');
    if (prefix) {
      const dataRow = await db.get('SELECT * FROM mitratel_data WHERE lokasi_id=?', [lokasiId]);
      const coreNum = String(f.slot || '').split('_')[1] || '';
      const namaSite = prefix === 'NE' ? (dataRow ? dataRow.nama_site_ne : '') : (dataRow ? dataRow.nama_site_fe : '');
      const lat = prefix === 'NE' ? (dataRow ? dataRow.tikor_ne_lat : '') : (dataRow ? dataRow.tikor_fe_lat : '');
      const lon = prefix === 'NE' ? (dataRow ? dataRow.tikor_ne_long : '') : (dataRow ? dataRow.tikor_fe_long : '');
      const auto = prefix + ' CORE ' + coreNum + ' - ' + (namaSite || ('SITE ' + prefix)) + ' - LAT ' + (lat || '-') + ' / LON ' + (lon || '-');
      if (tag === '') tag = auto;
    }
    fotos.push({ slot: f.slot, file_name: f.file_name, original_name: f.original_name, hasil_scan: f.hasil_scan, tag: tag, created_at: f.created_at });
  }
  return { lokasi, data, fotos };
}

async function saveData(lokasiId, body) {
  await siap();
  const clean = normalizeData(body);
  const keys = Object.keys(clean);
  const exists = await db.get('SELECT lokasi_id FROM mitratel_data WHERE lokasi_id=?', [lokasiId]);
  if (exists) {
    await db.run('UPDATE mitratel_data SET ' + keys.map(k => k + '=?').join(',') + ',updated_at=CURRENT_TIMESTAMP WHERE lokasi_id=?',
      [...keys.map(k => clean[k]), lokasiId]);
  } else {
    await db.run('INSERT INTO mitratel_data (lokasi_id,' + keys.join(',') + ') VALUES (?,' + keys.map(() => '?').join(',') + ')',
      [lokasiId, ...keys.map(k => clean[k])]);
  }
  return getData(lokasiId);
}

function validSlot(slot) {
  return /^(kalibrasi_[1-3]|ne_(?:[1-9]|1[0-4])|fe_(?:[1-9]|1[0-4]))$/.test(String(slot || ''));
}

// ── Berkas foto di folder vs baris di tabel mitratel_foto ───────────────────
// Kadang file berhasil ter-upload tapi baris DB-nya hilang (mis. database pernah
// tertimpa instance server lain). Supaya eviden yang sudah di-upload tidak
// hilang saat generate, kita baca juga isi folder dan selaraskan keduanya.
const RE_FILE_FOTO = /^(kalibrasi_[1-3]|ne_(?:[1-9]|1[0-4])|fe_(?:[1-9]|1[0-4]))_\d+\.[a-z0-9]+$/i;

function fotoDariDisk(lokasiId) {
  const dir = path.join(UPLOAD_ROOT, safeName(lokasiId));
  const hasil = {};
  if (!fs.existsSync(dir)) return hasil;
  for (const ent of fs.readdirSync(dir, { withFileTypes: true })) {
    if (!ent.isFile()) continue;
    const m = RE_FILE_FOTO.exec(ent.name);
    if (!m) continue;
    const slot = m[1].toLowerCase();
    const full = path.join(dir, ent.name);
    // per slot: ambil yang paling baru
    let mtime = 0;
    try { mtime = fs.statSync(full).mtimeMs; } catch (_) { continue; }
    if (!hasil[slot] || mtime > hasil[slot].mtime) hasil[slot] = { file_name: ent.name, mtime };
  }
  return hasil;
}

let _backfillFoto = false;
async function sinkronFotoDariDisk() {
  if (_backfillFoto) return;
  _backfillFoto = true;
  try {
    if (!fs.existsSync(UPLOAD_ROOT)) return;
    for (const ent of fs.readdirSync(UPLOAD_ROOT, { withFileTypes: true })) {
      if (!ent.isDirectory()) continue;
      const id = Number(ent.name);
      if (!id) continue;
      const disk = fotoDariDisk(id);
      const baris = await db.all('SELECT id,slot,file_name FROM mitratel_foto WHERE lokasi_id=?', [id]);
      const perSlot = new Map(baris.map(b => [b.slot, b]));
      for (const [slot, info] of Object.entries(disk)) {
        const b = perSlot.get(slot);
        if (!b) {
          await db.run('INSERT INTO mitratel_foto (lokasi_id,slot,file_name,original_name,tag) VALUES (?,?,?,?,?)',
            [id, slot, info.file_name, '', '']);
          console.log(`[mitratel] foto di folder tanpa baris DB → dikembalikan: lokasi ${id} / ${slot} = ${info.file_name}`);
        } else if (b.file_name !== info.file_name &&
                   !fs.existsSync(path.join(UPLOAD_ROOT, ent.name, b.file_name))) {
          await db.run('UPDATE mitratel_foto SET file_name=? WHERE id=?', [info.file_name, b.id]);
          console.log(`[mitratel] rujukan foto diperbaiki: lokasi ${id} / ${slot} = ${info.file_name}`);
        }
      }
      // baris DB yang file-nya sudah tidak ada → pakai berkas terbaru milik slot itu
      for (const b of baris) {
        if (fs.existsSync(path.join(UPLOAD_ROOT, ent.name, b.file_name))) continue;
        if (disk[b.slot]) await db.run('UPDATE mitratel_foto SET file_name=? WHERE id=?', [disk[b.slot].file_name, b.id]);
      }
    }
  } catch (e) {
    console.warn('[mitratel] sinkron foto gagal:', e.message);
  }
}
function imageExt(name, mime) {
  const ext = path.extname(name || '').toLowerCase();
  if (['.jpg','.jpeg','.png','.webp','.bmp'].includes(ext)) return ext === '.jpeg' ? '.jpg' : ext;
  if (/png/.test(mime || '')) return '.png';
  return '.jpg';
}

async function savePhoto(lokasiId, slot, file) {
  await siap();
  if (!validSlot(slot)) throw new Error('Slot foto tidak valid.');
  if (!file || !file.buffer) throw new Error('Foto belum dipilih.');
  if (!/^image\//i.test(file.mimetype || '')) throw new Error('Berkas harus berupa gambar JPG, PNG, WEBP, atau BMP.');
  const old = await db.get('SELECT file_name FROM mitratel_foto WHERE lokasi_id=? AND slot=?', [lokasiId, slot]);
  const dir = path.join(UPLOAD_ROOT, safeName(lokasiId));
  fs.mkdirSync(dir, { recursive: true });
  const filename = safeName(slot) + '_' + Date.now() + imageExt(file.originalname, file.mimetype);
  fs.writeFileSync(path.join(dir, filename), file.buffer);
  if (old && old.file_name) { try { fs.rmSync(path.join(dir, old.file_name), { force: true }); } catch (_) {} }
  // Bersihkan SEMUA berkas lama milik slot ini (bukan hanya yang direferensikan
  // baris DB): sisa upload lama bisa "hidup lagi" lewat sinkronisasi folder.
  try {
    const reSlot = new RegExp('^' + slot + '_\\d+\\.[a-z0-9]+$', 'i');
    for (const n of fs.readdirSync(dir)) if (n !== filename && reSlot.test(n)) fs.rmSync(path.join(dir, n), { force: true });
  } catch (_) { /* tetap lanjut — berkas baru sudah tertulis */ }
  const dataRow = await db.get('SELECT * FROM mitratel_data WHERE lokasi_id=?', [lokasiId]);
  const prefix = String(slot || '').startsWith('ne_') ? 'NE' : (String(slot || '').startsWith('fe_') ? 'FE' : '');
  let tag = '';
  if (prefix) {
    const coreNum = String(slot || '').split('_')[1] || '';
    const namaSite = prefix === 'NE' ? (dataRow ? dataRow.nama_site_ne : '') : (dataRow ? dataRow.nama_site_fe : '');
    const lat = prefix === 'NE' ? (dataRow ? dataRow.tikor_ne_lat : '') : (dataRow ? dataRow.tikor_fe_lat : '');
    const lon = prefix === 'NE' ? (dataRow ? dataRow.tikor_ne_long : '') : (dataRow ? dataRow.tikor_fe_long : '');
    tag = prefix + ' CORE ' + coreNum + ' - ' + (namaSite || ('SITE ' + prefix)) + ' - LAT ' + (lat || '-') + ' / LON ' + (lon || '-');
  }
  if (old) await db.run('UPDATE mitratel_foto SET file_name=?,original_name=?,hasil_scan=?,tag=? WHERE lokasi_id=? AND slot=?', [filename, file.originalname || '', '', tag, lokasiId, slot]);
  else await db.run('INSERT INTO mitratel_foto (lokasi_id,slot,file_name,original_name,tag) VALUES (?,?,?,?,?)', [lokasiId, slot, filename, file.originalname || '', tag]);
  return { slot, file_name: filename, url: '/api/mitratel/foto/' + lokasiId + '/' + encodeURIComponent(filename) };
}

async function getTag(lokasiId, slot) {
  await siap();
  const row = await db.get('SELECT slot,file_name,tag FROM mitratel_foto WHERE lokasi_id=? AND slot=?', [lokasiId, slot]);
  if (!row) return '';
  const dataRow = await db.get('SELECT * FROM mitratel_data WHERE lokasi_id=?', [lokasiId]);
  const prefix = String(slot || '').startsWith('ne_') ? 'NE' : (String(slot || '').startsWith('fe_') ? 'FE' : '');
  if (prefix) {
    const coreNum = String(slot || '').split('_')[1] || '';
    const namaSite = prefix === 'NE' ? (dataRow ? dataRow.nama_site_ne : '') : (dataRow ? dataRow.nama_site_fe : '');
    const lat = prefix === 'NE' ? (dataRow ? dataRow.tikor_ne_lat : '') : (dataRow ? dataRow.tikor_fe_lat : '');
    const lon = prefix === 'NE' ? (dataRow ? dataRow.tikor_ne_long : '') : (dataRow ? dataRow.tikor_fe_long : '');
    const auto = prefix + ' CORE ' + coreNum + ' - ' + (namaSite || ('SITE ' + prefix)) + ' - LAT ' + (lat || '-') + ' / LON ' + (lon || '-');
    if (String(row.tag || '') === '') return auto;
  }
  return String(row.tag || '');
}

async function updateTag(lokasiId, slot, tag) {
  await siap();
  if (!validSlot(slot)) throw new Error('Slot foto tidak valid.');
  const clean = String(tag || '').trim();
  await db.run('UPDATE mitratel_foto SET tag=? WHERE lokasi_id=? AND slot=?', [clean, lokasiId, slot]);
  const next = await getTag(lokasiId, slot);
  return next;
}

async function deletePhoto(lokasiId, slot) {
  await siap();
  if (!validSlot(slot)) throw new Error('Slot foto tidak valid.');
  const old = await db.get('SELECT file_name FROM mitratel_foto WHERE lokasi_id=? AND slot=?', [lokasiId, slot]);
  if (!old) throw new Error('Foto tidak ditemukan.');
  await db.run('DELETE FROM mitratel_foto WHERE lokasi_id=? AND slot=?', [lokasiId, slot]);
  // Hapus SEMUA berkas milik slot ini — termasuk sisa upload lama yang tidak
  // direferensikan — supaya tidak "hidup lagi" lewat sinkronisasi folder.
  try {
    const dir = path.join(UPLOAD_ROOT, safeName(lokasiId));
    const re = new RegExp('^' + slot + '_\\d+\\.[a-z0-9]+$', 'i');
    for (const n of fs.readdirSync(dir)) if (re.test(n)) fs.rmSync(path.join(dir, n), { force: true });
  } catch (_) { /* data DB sudah terhapus */ }
  return true;
}

async function scanCalibration(lokasiId) {
  await siap();
  const rows = await db.all("SELECT * FROM mitratel_foto WHERE lokasi_id=? AND slot LIKE 'kalibrasi_%' ORDER BY slot", [lokasiId]);
  const hasil = {};   // slot → nilai efektif (hasil scan baru, atau nilai lama bila gagal)
  const gagal = {};   // slot → pesan kesalahan (bila pembacaan kali ini tidak berhasil)
  // Scan paralel per foto + tangkap error per foto: satu foto gagal (AI 503/hang)
  // tidak boleh membuat seluruh request mati dan hasil foto lain hilang.
  await Promise.all(rows.map(async row => {
    const file = path.join(UPLOAD_ROOT, safeName(lokasiId), row.file_name);
    const lama = String(row.hasil_scan || '');
    if (!fs.existsSync(file)) {
      hasil[row.slot] = lama;
      gagal[row.slot] = 'File foto tidak ditemukan.';
      return;
    }
    try {
      const angka = await rekapOpm.bacaAngkaOpm(file, { jenis: 'kalibrasi', odp: 'MITRATEL', port: row.slot });
      if (angka) {
        await db.run('UPDATE mitratel_foto SET hasil_scan=? WHERE lokasi_id=? AND slot=?', [angka, lokasiId, row.slot]);
        hasil[row.slot] = angka;
      } else {
        // Tidak ada angka terbaca — nilai scan lama TIDAK ditimpa (fotonya sama),
        // dan pengguna diberi tahu agar foto memuat layar alat.
        hasil[row.slot] = lama;
        gagal[row.slot] = 'Angka tidak terbaca — pastikan foto memuat layar/tampilan alat (dB).';
      }
    } catch (e) {
      hasil[row.slot] = lama;
      gagal[row.slot] = String(e.message || e).slice(0, 200);
    }
  }));
  return { hasil, gagal };
}

async function scanNeFe(lokasiId) {
  await siap();
  const rows = await db.all("SELECT * FROM mitratel_foto WHERE lokasi_id=? AND (slot LIKE 'ne_%' OR slot LIKE 'fe_%') ORDER BY slot", [lokasiId]);
  const hasil = {};
  const gagal = {};
  await Promise.all(rows.map(async row => {
    const file = path.join(UPLOAD_ROOT, safeName(lokasiId), row.file_name);
    const lama = String(row.hasil_scan || '');
    if (!fs.existsSync(file)) { hasil[row.slot] = lama; gagal[row.slot] = 'File foto tidak ditemukan.'; return; }
    try {
      const angka = await rekapOpm.bacaAngkaOpm(file, { jenis: 'opm_test', odp: 'MITRATEL', port: row.slot });
      if (angka) {
        await db.run('UPDATE mitratel_foto SET hasil_scan=? WHERE lokasi_id=? AND slot=?', [angka, lokasiId, row.slot]);
        hasil[row.slot] = angka;
      } else {
        hasil[row.slot] = lama;
        gagal[row.slot] = 'Angka tidak terbaca — pastikan foto memuat layar/tampilan alat (dB).';
      }
    } catch (e) {
      hasil[row.slot] = lama;
      gagal[row.slot] = String(e.message || e).slice(0, 200);
    }
  }));
  return { hasil, gagal };
}


function tags(d) {
  return {
    '{insert/new ring}': d.insert_new_ring, '{id span existing}': d.id_span_existing, '{nama ring}': d.nama_ring, '{id span}': d.id_span,
    '{id site ne}': d.id_site_ne, '{nama site ne}': d.nama_site_ne, '{tikor ne lat}': d.tikor_ne_lat, '{tikor ne long}': d.tikor_ne_long,
    '{id site fe}': d.id_site_fe, '{nama site fe}': d.nama_site_fe, '{tikor fe lat}': d.tikor_fe_lat, '{tikor fe long}': d.tikor_fe_long,
    '{pangjang kabel n to n}': d.panjang_kabel_n_to_n, '{panjang kabel  baru}': d.panjang_kabel_baru, '{jumlah core}': d.jumlah_core
  };
}

async function generateOtp(lokasiId) {
  await siap();
  if (!fs.existsSync(TEMPLATE_OTP)) throw new Error('Template report OPM MITRATEL tidak ditemukan.');
  const all = await getData(lokasiId);
  if (!all) throw new Error('Lokasi tidak ditemukan.');
  if (String(all.lokasi.tematik || '').toLowerCase() !== 'm-tel') throw new Error('Report ini hanya untuk lokasi dengan tematik M-tel.');
  const d = normalizeData(all.data);
  const photoMap = {};
  const scans = {};
  for (const f of all.fotos) {
    photoMap[f.slot] = path.join(UPLOAD_ROOT, safeName(lokasiId), f.file_name);
    scans[f.slot] = f.hasil_scan || '';
  }
  const outDir = path.join(UPLOAD_ROOT, safeName(lokasiId), 'report');
  fs.mkdirSync(outDir, { recursive: true });
  const base = 'OPM_MITRATEL_' + safeName(all.lokasi.nama_lokasi).slice(0, 80) + '_' + Date.now();
  const outXlsx = path.join(outDir, base + '.xlsx');
  const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'mitratel-otp-'));
  const cfgPath = path.join(tmp, 'config.json');
  // Nilai numerik utk EQUIPMENT CALIBRATION. JANGAN kirim string ke Value2:
  // dengan locale id-ID (titik = pemisah waktu), WPS/Excel membaca "0.31"
  // sebagai WAKTU 00:31 → tersimpan 0.0215, bukan 0.31.
  const scansNum = {};
  for (const k of ['kalibrasi_1', 'kalibrasi_2', 'kalibrasi_3']) {
    const s = String(scans[k] || '').trim();
    const n = Number(s.replace(',', '.'));
    scansNum[k] = (s !== '' && Number.isFinite(n)) ? n : null;
  }
  const scansNumNe = {};
  const scansNumFe = {};
  for (let i = 1; i <= 14; i++) {
    const kn = 'ne_' + i;
    const kf = 'fe_' + i;
    const sn = String(scans[kn] || '').trim();
    const nn = Number(sn.replace(',', '.'));
    scansNumNe[kn] = (sn !== '' && Number.isFinite(nn)) ? nn : null;
    const sf = String(scans[kf] || '').trim();
    const nf = Number(sf.replace(',', '.'));
    scansNumFe[kf] = (sf !== '' && Number.isFinite(nf)) ? nf : null;
  }
  const opticalLengthRaw = String(d.panjang_kabel_n_to_n || '').trim().replace(',', '.');
  const opticalLengthNum = Number(opticalLengthRaw);
  const opticalLength = Number.isFinite(opticalLengthNum) ? opticalLengthNum : null;
  fs.writeFileSync(cfgPath, JSON.stringify({ template: TEMPLATE_OTP, outXlsx, tags: tags(d), date: d.tanggal_submit_atp, photos: photoMap, scans, scansNum, scansNumNe, scansNumFe, opticalLength }), 'utf8');

  const script = [
    "$ErrorActionPreference='Stop'", "$cfg=Get-Content -Raw -LiteralPath " + q(cfgPath) + " | ConvertFrom-Json",
    "$xl=New-Object -ComObject Excel.Application", "$xl.Visible=$false", "$xl.DisplayAlerts=$false", "$wb=$null",
    "try {", "  Copy-Item -LiteralPath $cfg.template -Destination $cfg.outXlsx -Force", "  $wb=$xl.Workbooks.Open($cfg.outXlsx)",
    "  foreach($ws in $wb.Worksheets){ foreach($p in $cfg.tags.PSObject.Properties){ [void]$ws.UsedRange.Replace([string]$p.Name,[string]$p.Value) } }",
    "  $sheet3=$wb.Worksheets.Item('Sheet3'); $sheet3.Range('E14').Value2=[string]$cfg.date",
    "  $flp=$wb.Worksheets.Item('Test Result Core (FLP) ')",
    // EQUIPMENT CALIBRATION (First/Second/Third dB) — tulis sebagai ANGKA (Value2
    // menerima [double], bukan string) agar salah baca locale tidak terjadi.
    "  $celMap=@{ kalibrasi_1='I21'; kalibrasi_2='K21'; kalibrasi_3='O21' }; foreach($k in $celMap.Keys){ $cell=$flp.Range([string]$celMap[$k]); $n=$cfg.scansNum.$k; if($null -ne $n -and \"$n\" -ne ''){ $cell.Value2=[double]$n } else { $cell.Value2='' } }",
    "  for($i=1;$i -le 14;$i++){ $cell=$flp.Range('I'+(30+$i)); $n=$cfg.scansNumNe.('ne_'+$i); if($null -ne $n -and \"$n\" -ne ''){ $cell.Value2=[double]$n } else { $cell.Value2='' }; $cell=$flp.Range('P'+(30+$i)); $n=$cfg.scansNumFe.('fe_'+$i); if($null -ne $n -and \"$n\" -ne ''){ $cell.Value2=[double]$n } else { $cell.Value2='' } }",
    "  for($i=1;$i -le 14;$i++){ $has=$false; $n=$cfg.scansNumNe.('ne_'+$i); if($null -ne $n -and \"$n\" -ne ''){ $has=$true }; $n=$cfg.scansNumFe.('fe_'+$i); if($null -ne $n -and \"$n\" -ne ''){ $has=$true }; if($has -and $null -ne $cfg.opticalLength -and \"$($cfg.opticalLength)\" -ne ''){ $cell=$flp.Range('H'+(30+$i)); $cell.Value2=[double]$cfg.opticalLength } }",
    "  $opm=$wb.Worksheets.Item('OPM Measurement_Photos')",
    "  $map=@{ kalibrasi_1='C2'; kalibrasi_2='D2'; kalibrasi_3='E2' }",
    "  for($i=1;$i -le 14;$i++){ $map['ne_'+$i]=[char](66+$i)+'4'; $map['fe_'+$i]=[char](66+$i)+'6' }",
    "  foreach($p in $map.GetEnumerator()){ $slot=[string]$p.Key; $cell=$opm.Range([string]$p.Value); $tag=if($slot -like 'kalibrasi_*'){'{OPM Measurement_Photos'+$slot.Replace('kalibrasi_','')+'}'}elseif($slot -like 'ne_*'){'{KEGIATAN NE CORE '+$slot.Replace('ne_','')+'}'}else{'{KEGIATAN FE CORE '+$slot.Replace('fe_','')+'}'}; $cell.Value2=''; $img=$cfg.photos.$slot; if($img -and (Test-Path ([string]$img))){ $pad=3; $w=[Math]::Max(10,$cell.Width-2*$pad); $h=[Math]::Max(10,$cell.Height-2*$pad); [void]$opm.Shapes.AddPicture([string]$img,$false,$true,$cell.Left+$pad,$cell.Top+$pad,$w,$h) } }",
    "  foreach($ws in $wb.Worksheets){ [void]$ws.UsedRange.Replace('{First (dB)}',[string]$cfg.scans.kalibrasi_1); [void]$ws.UsedRange.Replace('{Second (dB)}',[string]$cfg.scans.kalibrasi_2); [void]$ws.UsedRange.Replace('{Third (dB)}',[string]$cfg.scans.kalibrasi_3) }",
    "  $xl.CalculateFull()",
    "  $opm.Activate()",
    "  $wb.Save()", "} finally { if($wb){$wb.Close($true)}; $xl.Quit(); [Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null }",
    // Pasca simpan: paksa tampilan Normal (bukan Page Layout/kolom cetak) dan
    // buka file di sheet foto. COM yang aktif di mesin ini adalah milik WPS Office
    // sehingga Window.View ditolak -- jadi sunting XML workbook-nya langsung.
    "try {",
    "  Add-Type -AssemblyName System.IO.Compression.FileSystem",
    "  Add-Type -AssemblyName System.IO.Compression",
    "  $zip=[System.IO.Compression.ZipFile]::Open($cfg.outXlsx,[System.IO.Compression.ZipArchiveMode]::Update)",
    "  $entWb=$zip.GetEntry('xl/workbook.xml')",
    "  if($entWb){ $sr=New-Object System.IO.StreamReader($entWb.Open()); $txt=$sr.ReadToEnd(); $sr.Dispose()",
    "    if($txt -match 'activeTab='){ $txt=$txt -replace 'activeTab=\"\\d+\"','activeTab=\"2\"' } else { $txt=$txt -replace '<workbookView ','<workbookView activeTab=\"2\" ' }",
    "    $entWb.Delete(); $newWb=$zip.CreateEntry('xl/workbook.xml',[System.IO.Compression.CompressionLevel]::Optimal)",
    "    $sw=New-Object System.IO.StreamWriter($newWb.Open()); $sw.Write($txt); $sw.Dispose() }",
    "  foreach($ent in @($zip.Entries | Where-Object { $_.FullName -like 'xl/worksheets/sheet*.xml' })){",
    "    $name=$ent.FullName",
    "    $sr=New-Object System.IO.StreamReader($ent.Open()); $txt=$sr.ReadToEnd(); $sr.Dispose()",
    "    $txt=$txt -replace ' view=\"pageLayout\"','' -replace ' view=\"pageBreakPreview\"','' -replace ' tabSelected=\"1\"',''",
    "    if($name -eq 'xl/worksheets/sheet3.xml'){ $txt=$txt -replace '<sheetView(\\s|>)','<sheetView tabSelected=\"1\"$1'; $txt=$txt -replace ' topLeftCell=\"[^\"]*\"','' -replace 'zoomScale=\"\\d+\"','zoomScale=\"85\"' -replace 'zoomScaleNormal=\"\\d+\"','zoomScaleNormal=\"100\"' -replace '<selection[^>]*/>','<selection activeCell=\"A1\" sqref=\"A1\"/>' }",
    "    $ent.Delete(); $newE=$zip.CreateEntry($name,[System.IO.Compression.CompressionLevel]::Optimal)",
    "    $sw=New-Object System.IO.StreamWriter($newE.Open()); $sw.Write($txt); $sw.Dispose() }",
    "  $zip.Dispose()",
    "} catch { Write-Warning ('pasca-proses tampilan gagal: ' + $_.Exception.Message) }"
  ].join('\n');
  try { await ps(script, 300000); } finally { try { fs.rmSync(tmp, { recursive: true, force: true }); } catch (_) {} }
  if (!fs.existsSync(outXlsx)) throw new Error('Excel gagal dibuat. Pastikan Microsoft Excel terpasang.');
  // PDF sudah tidak dibuat lagi — hanya XLSX (permintaan pengguna).
  return { xlsx: path.basename(outXlsx), pdf: '', baseUrl: '/api/mitratel/download/' + lokasiId + '/' };
}

// ─────────────────────────────────────────────────────────────────────────────
// ATP MITRATEL — eviden foto per placeholder pada template "ATP TEMPLATE.xlsx".
// Setiap slot punya sel tujuan (sheet + sel referensi) supaya foto masuk ke
// posisi yang benar, termasuk placeholder yang teksnya sama berulang.
// ─────────────────────────────────────────────────────────────────────────────
function colLetter(n) { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }

function bangunSlotAtp() {
  const slots = {};
  const add = (slot, sheet, cell, label, grup, blok, extra) => { slots[slot] = { sheet, cell, label, grup, blok, ...(extra || {}) }; };
  // Sheet "OTDR": baris 2 = Near End, 17 = Far End (FA), 32 = Insert Span (NEW).
  // Pola kolom per core: surrounding di kolom 5+6(n-1), photo di 8+6(n-1).
  const barisOtdr = [['ne', 2, 14], ['fa', 17, 14], ['new', 32, 4]];
  for (const [blok, row, max] of barisOtdr) {
    const nama = blok === 'ne' ? 'Near End' : (blok === 'fa' ? 'Far End' : 'Insert Span');
    for (let n = 1; n <= max; n++) {
      add(`otdr_${blok}_${n}_surround`, 'OTDR', colLetter(5 + 6 * (n - 1)) + row, `OTDR PHOTO & SURROUNDING ${n} — ${nama}`, 'otdr', blok);
      add(`otdr_${blok}_${n}_photo`, 'OTDR', colLetter(8 + 6 * (n - 1)) + row, `OTDR PHOTO ${n} — ${nama}`, 'otdr', blok);
    }
  }
  // Sheet "OTB, ODP & LABELLING"
  add('otb_ne_photo', 'OTB, ODP & LABELLING', 'E2', 'NE ODP / OTB PHOTO', 'otb', 'ne');
  add('otb_ne_label', 'OTB, ODP & LABELLING', 'H2', 'NE LABEL ODP / OTB', 'otb', 'ne');
  add('otb_fe_photo', 'OTB, ODP & LABELLING', 'E25', 'FE ODP / OTB PHOTO', 'otb', 'fe');
  add('otb_fe_label', 'OTB, ODP & LABELLING', 'H25', 'FE LABEL ODP / OTB', 'otb', 'fe');
  add('otb_ins_photo', 'OTB, ODP & LABELLING', 'E51', 'INSERT SITE ODP / OTB PHOTO', 'otb', 'ins');
  add('otb_ins_label', 'OTB, ODP & LABELLING', 'H51', 'INSERT SITE LABEL ODP / OTB', 'otb', 'ins');
  // Sheet "Cable Route In" — NE baris 5..59, FE baris 65..97
  add('cri_ne_siteview1', 'Cable Route In', 'B5', 'SITE VIEW FROM OUT SITE 1', 'cri', 'ne');
  add('cri_ne_siteview2', 'Cable Route In', 'E5', 'SITE VIEW FROM OUT SITE 2', 'cri', 'ne');
  add('cri_ne_lastpole1', 'Cable Route In', 'G5', 'PHOTO ODP LAST POLE 1', 'cri', 'ne');
  add('cri_ne_lastpole2', 'Cable Route In', 'J5', 'PHOTO ODP LAST POLE 2', 'cri', 'ne');
  add('cri_ne_galian1', 'Cable Route In', 'B16', 'View Kabel FO keluar dari Galian/Handhole (1)', 'cri', 'ne');
  add('cri_ne_masuk1', 'Cable Route In', 'G16', 'View Kabel FO masuk ke dalam Site (1)', 'cri', 'ne');
  add('cri_ne_galian2', 'Cable Route In', 'B27', 'View Kabel FO keluar dari Galian/Handhole (2)', 'cri', 'ne');
  add('cri_ne_masuk2', 'Cable Route In', 'G27', 'View Kabel FO masuk ke dalam Site (2)', 'cri', 'ne');
  add('cri_ne_kwh', 'Cable Route In', 'B39', 'Kabel FO masuk melalui tiang KWH', 'cri', 'ne');
  add('cri_ne_conduit', 'Cable Route In', 'G39', 'Kabel FO menggunakan conduit ditanam menuju Rack OTB', 'cri', 'ne');
  add('cri_ne_tray', 'Cable Route In', 'B51', 'Route kabel mengikuti tray horizontal', 'cri', 'ne');
  add('cri_ne_tanam', 'Cable Route In', 'G51', 'Route kabel tanam menuju Rack OTB', 'cri', 'ne');
  add('cri_fe_galian', 'Cable Route In', 'B65', 'View Kabel FO keluar dari Galian/Handhole', 'cri', 'fe');
  add('cri_fe_masuk', 'Cable Route In', 'G65', 'View Kabel FO masuk ke dalam Site', 'cri', 'fe');
  add('cri_fe_kwh', 'Cable Route In', 'B77', 'Kabel FO masuk melalui tiang KWH', 'cri', 'fe');
  add('cri_fe_conduit', 'Cable Route In', 'G77', 'Kabel FO menggunakan conduit ditanam menuju Rack OTB', 'cri', 'fe');
  add('cri_fe_tray', 'Cable Route In', 'B89', 'Route kabel mengikuti tray horizontal', 'cri', 'fe');
  add('cri_fe_tanam', 'Cable Route In', 'G89', 'Route kabel tanam menuju Rack OTB', 'cri', 'fe');
  // Sheet "Cable Route Out" — NE: JC Petikan(2), Slack(2), Closure(2), TIANG(4);
  // FE: Slack(2), Closure(2), TIANG(4). Sel = placeholder {..} pada merge area.
  const croNe = [['jc1', 'B5', 'JC Petikan 1'], ['jc2', 'G5', 'JC Petikan 2'],
    ['slack1', 'B16', 'Slack 1'], ['slack2', 'G16', 'Slack 2'],
    ['closure1', 'B28', 'Closure 1'], ['closure2', 'G28', 'Closure 2'],
    ['tiang1', 'B39', 'TIANG 1'], ['tiang2', 'G39', 'TIANG 2'],
    ['tiang3', 'B50', 'TIANG 3'], ['tiang4', 'G50', 'TIANG 4']];
  croNe.forEach(([k, c, l]) => add('cro_ne_' + k, 'Cable Route Out', c, l + ' — NEAR END', 'cro', 'ne'));
  const croFe = [['slack1', 'B64', 'Slack 1'], ['slack2', 'G64', 'Slack 2'],
    ['closure1', 'B76', 'Closure 1'], ['closure2', 'G76', 'Closure 2'],
    ['tiang1', 'B87', 'TIANG 1'], ['tiang2', 'G87', 'TIANG 2'],
    ['tiang3', 'B98', 'TIANG 3'], ['tiang4', 'G98', 'TIANG 4']];
  croFe.forEach(([k, c, l]) => add('cro_fe_' + k, 'Cable Route Out', c, l + ' — FAR END', 'cro', 'fe'));
  // Sheet "KMZ" — unggahan KML/KMZ dirender jadi gambar peta dan ditempatkan di
  // area merge di bawah judulnya; nama file asli ditulis pada area "KMZ File".
  add('kmz_plan', 'KMZ', 'A4', 'SITE KMZ PLAN (Span End to End)', 'kmz', 'plan');
  add('kmz_actual', 'KMZ', 'N4', 'AS BUILT DRAWING/ACTUAL (Span End to End)', 'kmz', 'actual');
  add('kmz_plan_insert', 'KMZ', 'A39', 'PLAN New cable Insert (*if any)', 'kmz', 'plan_ins');
  add('kmz_actual_insert', 'KMZ', 'N39', 'ACTUAL New cable Insert (*if any)', 'kmz', 'act_ins');
  // Sheet "Lay Out ABD (NE & FE)" — unggahan ABD (PDF/gambar) dirender ke area
  // gambar tanpa merge (pakai range eksplisit, sesuai jangkar gambar contoh).
  add('abd_ne', 'Lay Out ABD (NE & FE)', 'B8', 'NEAR END (for New Span)', 'abd', 'ne', { range: 'B8:T43' });
  add('abd_fe', 'Lay Out ABD (NE & FE)', 'C56', 'FAR END (for New Span)', 'abd', 'fe', { range: 'C56:Q87' });
  add('abd_ins', 'Lay Out ABD (NE & FE)', 'C103', 'NEAR END Only (for Site Insert)', 'abd', 'ins', { range: 'C103:Q133' });
  // Sheet "FLP Schematic Core" — slot ABD schematic + input SPAN ID / SPAN NAME.
  add('flp_sc_abd', 'FLP Schematic Core', 'B2', 'FLP Schematic Core (ABD)', 'flp', 'flp', { range: 'B7:T55' });
  return slots;
}
const SLOT_ATP = bangunSlotAtp();

// Teks pendukung yang ditulis saat generate (mengganti contoh milik template):
// nama file asli pada area "KMZ File", SITE NAME/ID + tabel Site Id/Site Name/File PDF pada ABD.
const TEKS_ATP = [
  { slot: 'kmz_plan', sheet: 'KMZ', cell: 'A32', jenis: 'file' },
  { slot: 'kmz_actual', sheet: 'KMZ', cell: 'N32', jenis: 'file' },
  { slot: 'kmz_plan_insert', sheet: 'KMZ', cell: 'A66', jenis: 'file' },
  { slot: 'kmz_actual_insert', sheet: 'KMZ', cell: 'N66', jenis: 'file' },
  { slot: 'abd_ne', sheet: 'Lay Out ABD (NE & FE)', cell: 'B6', jenis: 'sitename', prefix: 'SITE NAME  : ' },
  { slot: 'abd_ne', sheet: 'Lay Out ABD (NE & FE)', cell: 'B7', jenis: 'siteid', prefix: 'SITE ID : ' },
  { slot: 'abd_ne', sheet: 'Lay Out ABD (NE & FE)', cell: 'C47', jenis: 'tid' },
  { slot: 'abd_ne', sheet: 'Lay Out ABD (NE & FE)', cell: 'D47', jenis: 'tnama' },
  { slot: 'abd_ne', sheet: 'Lay Out ABD (NE & FE)', cell: 'E47', jenis: 'tfile' },
  { slot: 'abd_fe', sheet: 'Lay Out ABD (NE & FE)', cell: 'B52', jenis: 'sitename', prefix: 'SITE NAME  : ' },
  { slot: 'abd_fe', sheet: 'Lay Out ABD (NE & FE)', cell: 'B53', jenis: 'siteid', prefix: 'SITE ID  : ' },
  { slot: 'abd_fe', sheet: 'Lay Out ABD (NE & FE)', cell: 'B54', jenis: 'literal', value: 'FAR END (for New Span)' },
  { slot: 'abd_fe', sheet: 'Lay Out ABD (NE & FE)', cell: 'C93', jenis: 'tid' },
  { slot: 'abd_fe', sheet: 'Lay Out ABD (NE & FE)', cell: 'D93', jenis: 'tnama' },
  { slot: 'abd_fe', sheet: 'Lay Out ABD (NE & FE)', cell: 'E93', jenis: 'tfile' },
  { slot: 'abd_ins', sheet: 'Lay Out ABD (NE & FE)', cell: 'B99', jenis: 'sitename', prefix: 'SITE NAME  : ' },
  { slot: 'abd_ins', sheet: 'Lay Out ABD (NE & FE)', cell: 'B100', jenis: 'siteid', prefix: 'SITE ID : ' },
  { slot: 'abd_ins', sheet: 'Lay Out ABD (NE & FE)', cell: 'B101', jenis: 'literal', value: 'NEAR END Only (for Site Insert)' },
  { slot: 'abd_ins', sheet: 'Lay Out ABD (NE & FE)', cell: 'C140', jenis: 'tid' },
  { slot: 'abd_ins', sheet: 'Lay Out ABD (NE & FE)', cell: 'D140', jenis: 'tnama' },
  { slot: 'abd_ins', sheet: 'Lay Out ABD (NE & FE)', cell: 'E140', jenis: 'tfile' },
  { slot: 'flp_sc_abd', sheet: 'FLP Schematic Core', cell: 'B2', jenis: 'flptitle' },
  { slot: 'flp_sc_abd', sheet: 'FLP Schematic Core', cell: 'B3', jenis: 'spanid', prefix: 'SPAN ID : ' },
  { slot: 'flp_sc_abd', sheet: 'FLP Schematic Core', cell: 'B4', jenis: 'spanname', prefix: 'SPAN NAME : ' },
  { slot: 'flp_sc_abd', sheet: 'FLP Schematic Core', cell: 'B7', jenis: 'literal', value: 'FLP Schematic Core' },
];

// Gambar contoh milik template pada sheet Lay Out ABD harus dihapus saat generate
// supaya contoh milik template tidak bocor ke report (dihapus walau slot kosong).
const HAPUS_GAMBAR = [
  { sheet: 'Lay Out ABD (NE & FE)', name: 'Picture 7' },   // contoh ABD NE
  { sheet: 'Lay Out ABD (NE & FE)', name: 'Picture 2' },   // contoh ABD FE
  { sheet: 'Lay Out ABD (NE & FE)', name: 'Picture 10' },  // contoh ABD insert
  { sheet: 'FLP Schematic Core', name: 'Picture 1' },     // contoh schematic FLP
];
const RE_ATP_FILE = new RegExp('^(' + Object.keys(SLOT_ATP).map(k => k.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')).join('|') + ')_\\d+\\.[a-z0-9]+$', 'i');

function atpDir(lokasiId) { return path.join(UPLOAD_ROOT, safeName(lokasiId), 'atp'); }
function validSlotAtp(slot) { return Object.prototype.hasOwnProperty.call(SLOT_ATP, String(slot || '')); }

function fotoAtpDariDisk(lokasiId) {
  const dir = atpDir(lokasiId);
  const hasil = {};
  if (!fs.existsSync(dir)) return hasil;
  for (const ent of fs.readdirSync(dir, { withFileTypes: true })) {
    if (!ent.isFile()) continue;
    const m = RE_ATP_FILE.exec(ent.name);
    if (!m) continue;
    const slot = m[1];
    const full = path.join(dir, ent.name);
    let mtime = 0;
    try { mtime = fs.statSync(full).mtimeMs; } catch (_) { continue; }
    if (!hasil[slot] || mtime > hasil[slot].mtime) hasil[slot] = { file_name: ent.name, mtime };
  }
  return hasil;
}

let _backfillAtp = false;
async function sinkronAtpDariDisk() {
  if (_backfillAtp) return;
  _backfillAtp = true;
  try {
    if (!fs.existsSync(UPLOAD_ROOT)) return;
    for (const ent of fs.readdirSync(UPLOAD_ROOT, { withFileTypes: true })) {
      if (!ent.isDirectory()) continue;
      const id = Number(ent.name);
      if (!id) continue;
      const disk = fotoAtpDariDisk(id);
      const baris = await db.all('SELECT id,slot,file_name FROM atp_foto WHERE lokasi_id=?', [id]);
      const perSlot = new Map(baris.map(b => [b.slot, b]));
      for (const [slot, info] of Object.entries(disk)) {
        const b = perSlot.get(slot);
        if (!b) {
          await db.run('INSERT INTO atp_foto (lokasi_id,slot,file_name,original_name) VALUES (?,?,?,?)', [id, slot, info.file_name, '']);
          console.log(`[mitratel] foto ATP di folder tanpa baris DB → dikembalikan: lokasi ${id} / ${slot} = ${info.file_name}`);
        } else if (b.file_name !== info.file_name && !fs.existsSync(path.join(atpDir(id), b.file_name))) {
          await db.run('UPDATE atp_foto SET file_name=? WHERE id=?', [info.file_name, b.id]);
        }
      }
      for (const b of baris) {
        if (fs.existsSync(path.join(atpDir(id), b.file_name))) continue;
        if (disk[b.slot]) await db.run('UPDATE atp_foto SET file_name=? WHERE id=?', [disk[b.slot].file_name, b.id]);
      }
    }
  } catch (e) {
    console.warn('[mitratel] sinkron foto ATP gagal:', e.message);
  }
}

// ── Nilai Optical Length + Redaman per core (sheet Test Result Core (FLP)) ──
const MAX_CORE = 14;
// Hanya angka (boleh koma/titik, boleh minus) atau kosong.
function angkaBersih(v) {
  const s = String(v == null ? '' : v).trim().replace(',', '.');
  if (s === '') return '';
  if (!/^-?\d*\.?\d+$/.test(s)) throw new Error('Nilai harus berupa angka (mis. 2.051 atau -7.14).');
  return s.slice(0, 20);
}
async function getAtpMeta(lokasiId) {
  const rows = await db.all('SELECT core,optical_length,att_ab,att_ba FROM atp_meta WHERE lokasi_id=? ORDER BY core', [lokasiId]);
  return (rows || []).map(r => ({
    core: Number(r.core),
    optical_length: String(r.optical_length || '').trim(),
    att_ab: String(r.att_ab || '').trim(),
    att_ba: String(r.att_ba || '').trim()
  }));
}
async function setAtpCoreMeta(lokasiId, body) {
  await siap();
  const core = Number(body && body.core);
  if (!Number.isInteger(core) || core < 1 || core > MAX_CORE) throw new Error('Core tidak valid (1–14).');
  const ol = angkaBersih(body && body.optical_length);
  const ab = angkaBersih(body && body.att_ab);
  const ba = angkaBersih(body && body.att_ba);
  if (!ol && !ab && !ba) {
    // Baris yang seluruh isinya kosong tidak perlu disimpan.
    await db.run('DELETE FROM atp_meta WHERE lokasi_id=? AND core=?', [lokasiId, core]);
    return getAtpMeta(lokasiId);
  }
  const lama = await db.get('SELECT core FROM atp_meta WHERE lokasi_id=? AND core=?', [lokasiId, core]);
  if (lama) {
    await db.run('UPDATE atp_meta SET optical_length=?,att_ab=?,att_ba=?,updated_at=datetime(\'now\',\'localtime\') WHERE lokasi_id=? AND core=?',
      [ol, ab, ba, lokasiId, core]);
  } else {
    await db.run('INSERT INTO atp_meta (lokasi_id,core,optical_length,att_ab,att_ba) VALUES (?,?,?,?,?)',
      [lokasiId, core, ol, ab, ba]);
  }
  return getAtpMeta(lokasiId);
}

async function getAtpData(lokasiId) {
  await siap();
  const lokasi = await db.get('SELECT id,nama_lokasi,pid,tematik FROM lokasi WHERE id=?', [lokasiId]);
  if (!lokasi) return null;
  const data = await db.get('SELECT * FROM mitratel_data WHERE lokasi_id=?', [lokasiId]) || { lokasi_id: Number(lokasiId), jumlah_core: '2' };
  const rows = await db.all('SELECT slot,file_name,original_name,site_name,site_id,optical_length,att_ab,att_ba,created_at FROM atp_foto WHERE lokasi_id=? ORDER BY slot', [lokasiId]);
  const disk = fotoAtpDariDisk(lokasiId);
  const folder = atpDir(lokasiId);
  for (const f of rows) {
    if (!fs.existsSync(path.join(folder, f.file_name)) && disk[f.slot]) f.file_name = disk[f.slot].file_name;
  }
  for (const [slot, info] of Object.entries(disk)) {
    if (!rows.some(f => f.slot === slot)) rows.push({ slot, file_name: info.file_name, original_name: '', site_name: '', site_id: '', optical_length: '', att_ab: '', att_ba: '', created_at: '' });
  }
  const fotos = rows.filter(f => validSlotAtp(f.slot) && fs.existsSync(path.join(folder, f.file_name)));
  return { lokasi, data, fotos, meta: await getAtpMeta(lokasiId), slots: Object.entries(SLOT_ATP).map(([slot, v]) => ({ slot, ...v })) };
}

async function saveAtpPhoto(lokasiId, slot, file, extra = {}) {
  await siap();
  if (!validSlotAtp(slot)) throw new Error('Slot foto ATP tidak valid.');
  if (!file || !file.buffer) throw new Error('Berkas belum dipilih.');
  const dir = atpDir(lokasiId);
  fs.mkdirSync(dir, { recursive: true });
  const tipe = tipeSlotAtp(slot);
  const namaAsli = String(file.originalname || '').trim();
  const ext = path.extname(namaAsli).toLowerCase();
  let filename = '';
  let asliBaru = '';   // nama berkas asli (KML/KMZ/PDF) pada subfolder asli/, bila ada

  if (tipe === 'kmz') {
    // KMZ slot sekarang menerima gambar (JPG/PNG/WEBP/BMP) atau PDF (dirender jadi gambar).
    const okPdf = ext === '.pdf' || /pdf/i.test(file.mimetype || '');
    const okImg = /^image\//i.test(file.mimetype || '') || ['.jpg', '.jpeg', '.png', '.webp', '.bmp'].includes(ext);
    if (!okPdf && !okImg) throw new Error('Slot KMZ mohon unggah berkas gambar (JPG/PNG/WEBP/BMP) atau PDF.');
    if (okPdf) {
      asliBaru = simpanBerkasAsli(dir, slot, '.pdf', file.buffer);
      if (typeof helper.renderPdf !== 'function') throw new Error('Konversi PDF ke gambar tidak tersedia di server.');
      const pngs = await helper.renderPdf(path.join(dir, 'asli', asliBaru));
      if (!pngs || !pngs.length) throw new Error('PDF KMZ gagal dirender menjadi gambar.');
      filename = slot + '_' + Date.now() + '.png';
      fs.copyFileSync(pngs[0], path.join(dir, filename));
    } else {
      filename = slot + '_' + Date.now() + imageExt(namaAsli, file.mimetype);
      fs.writeFileSync(path.join(dir, filename), file.buffer);
      if (!extra.keep_asli) asliBaru = simpanBerkasAsli(dir, slot, ext, file.buffer);
    }
  } else if (tipe === 'abd') {
    // ABD: PDF dirender menjadi gambar (halaman 1); gambar dipakai langsung.
    const okPdf = ext === '.pdf' || /pdf/i.test(file.mimetype || '');
    const okImg = /^image\//i.test(file.mimetype || '') || ['.jpg', '.jpeg', '.png', '.webp', '.bmp'].includes(ext);
    if (!okPdf && !okImg) throw new Error('ABD harus berupa berkas PDF atau gambar (JPG/PNG/WEBP/BMP).');
    if (okPdf) {
      asliBaru = simpanBerkasAsli(dir, slot, '.pdf', file.buffer);
      if (typeof helper.renderPdf !== 'function') throw new Error('Konversi PDF ke gambar tidak tersedia di server.');
      const pngs = await helper.renderPdf(path.join(dir, 'asli', asliBaru));
      if (!pngs || !pngs.length) throw new Error('PDF ABD gagal dirender menjadi gambar — pastikan PDF tidak terproteksi atau rusak.');
      filename = slot + '_' + Date.now() + '.png';
      fs.copyFileSync(pngs[0], path.join(dir, filename));
    } else {
      filename = slot + '_' + Date.now() + imageExt(namaAsli, file.mimetype);
      fs.writeFileSync(path.join(dir, filename), file.buffer);
      // Unggahan hasil edit/tag (keep_asli) bukan berkas sumber ABD — jangan
      // menimpa/mengganti PDF asli yang sudah tersimpan.
      if (!extra.keep_asli) asliBaru = simpanBerkasAsli(dir, slot, ext, file.buffer);
    }
  } else {
    // Foto biasa (OTDR / OTB / Cable Route In & Out).
    if (!/^image\//i.test(file.mimetype || '')) throw new Error('Berkas harus berupa gambar JPG, PNG, WEBP, atau BMP.');
    filename = slot + '_' + Date.now() + imageExt(namaAsli, file.mimetype);
    fs.writeFileSync(path.join(dir, filename), file.buffer);
  }

  // Bersihkan SEMUA berkas lama milik slot ini (cegah file yatim "hidup lagi")
  try {
    const re = new RegExp('^' + slot + '_\\d+\\.[a-z0-9]+$', 'i');
    for (const n of fs.readdirSync(dir)) if (n !== filename && re.test(n)) fs.rmSync(path.join(dir, n), { force: true });
  } catch (_) { /* berkas baru sudah tertulis */ }
  // Berkas asli (KML/KMZ/PDF) di subfolder asli/ ikut diganti agar tidak menumpuk.
  // Unggahan hasil edit/tag (keep_asli) tidak boleh menyentuh berkas sumber sama sekali.
  try {
    const asliDir = path.join(dir, 'asli');
    if (fs.existsSync(asliDir) && !extra.keep_asli) {
      const re = new RegExp('^' + slot + '_\\d+\\.', 'i');
      for (const n of fs.readdirSync(asliDir)) if (n !== asliBaru && re.test(n)) fs.rmSync(path.join(asliDir, n), { force: true });
    }
  } catch (_) { /* pembersihan asli bersifat sekunder */ }
  // ABD membawa SITE NAME + SITE ID; slot lain dikosongkan.
  const sn = tipe === 'abd' ? String(extra.site_name || '').trim().slice(0, 120) : '';
  const si = tipe === 'abd' ? String(extra.site_id || '').trim().slice(0, 60) : '';
  const old = await db.get('SELECT id FROM atp_foto WHERE lokasi_id=? AND slot=?', [lokasiId, slot]);
  if (old) await db.run('UPDATE atp_foto SET file_name=?,original_name=?,site_name=?,site_id=? WHERE id=?', [filename, namaAsli, sn, si, old.id]);
  else await db.run('INSERT INTO atp_foto (lokasi_id,slot,file_name,original_name,site_name,site_id) VALUES (?,?,?,?,?,?)', [lokasiId, slot, filename, namaAsli, sn, si]);
  return getAtpData(lokasiId);
}

function tipeSlotAtp(slot) {
  const g = (SLOT_ATP[String(slot || '')] || {}).grup;
  if (g === 'kmz') return 'kmz';
  if (g === 'flp') return 'abd'; // FLP Schematic Core: upload ABD (PDF/gambar) → gambar schematic
  if (g === 'abd') return 'abd';
  return 'foto';
}

// Simpan berkas asli (KML/KMZ/PDF) di subfolder atp/asli/ — di luar pola sinkron
// foto (RE_ATP_FILE) agar tidak pernah disisipkan langsung ke Excel.
function simpanBerkasAsli(dir, slot, ext, buffer) {
  if (!buffer) return '';
  const asliDir = path.join(dir, 'asli');
  fs.mkdirSync(asliDir, { recursive: true });
  const nama = slot + '_' + Date.now() + (ext || '');
  fs.writeFileSync(path.join(asliDir, nama), buffer);
  return nama;
}

// Fallback offline: gambar skematik rute di atas kanvas putih (tanpa tile satelit).
async function kmzSkematik(features, outPath) {
  const { createCanvas } = require('@napi-rs/canvas');
  const W = 1400, H = 900, pad = 70;
  const semua = [];
  features.forEach(f => f.points.forEach(p => semua.push(p)));
  if (!semua.length) throw new Error('KML/KMZ tidak memiliki koordinat.');
  let minLon = Infinity, maxLon = -Infinity, minLat = Infinity, maxLat = -Infinity;
  semua.forEach(p => { if (p[0] < minLon) minLon = p[0]; if (p[0] > maxLon) maxLon = p[0]; if (p[1] < minLat) minLat = p[1]; if (p[1] > maxLat) maxLat = p[1]; });
  const spanLon = Math.max(maxLon - minLon, 1e-6), spanLat = Math.max(maxLat - minLat, 1e-6);
  const sc = Math.min((W - 2 * pad) / spanLon, (H - 2 * pad - 40) / spanLat);
  const cx = (minLon + maxLon) / 2, cy = (minLat + maxLat) / 2;
  const X = (lon) => W / 2 + (lon - cx) * sc;
  const Y = (lat) => (H - 40) / 2 - (lat - cy) * sc;
  const cv = createCanvas(W, H);
  const c = cv.getContext('2d');
  c.fillStyle = '#ffffff'; c.fillRect(0, 0, W, H);
  c.strokeStyle = '#e2e8f0'; c.lineWidth = 1;
  for (let x = 0; x <= W; x += 70) { c.beginPath(); c.moveTo(x, 0); c.lineTo(x, H - 40); c.stroke(); }
  for (let y = 0; y <= H - 40; y += 70) { c.beginPath(); c.moveTo(0, y); c.lineTo(W, y); c.stroke(); }
  features.forEach(f => {
    if (f.type !== 'line' && f.type !== 'polygon') return;
    c.strokeStyle = '#1d4ed8'; c.lineWidth = 5; c.lineJoin = 'round'; c.lineCap = 'round';
    c.beginPath();
    f.points.forEach((p, i) => (i ? c.lineTo(X(p[0]), Y(p[1])) : c.moveTo(X(p[0]), Y(p[1]))));
    if (f.type === 'polygon') c.closePath();
    c.stroke();
  });
  features.forEach(f => {
    if (f.type !== 'point' || !f.points.length) return;
    const x = X(f.points[0][0]), y = Y(f.points[0][1]);
    c.fillStyle = '#e11d48'; c.beginPath(); c.arc(x, y, 8, 0, Math.PI * 2); c.fill();
    c.lineWidth = 2.5; c.strokeStyle = '#ffffff'; c.stroke();
    if (f.name) {
      c.font = 'bold 15px Arial'; c.lineWidth = 3.5; c.strokeStyle = '#0f172a';
      c.strokeText(f.name, x + 12, y - 8); c.fillStyle = '#ffffff'; c.fillText(f.name, x + 12, y - 8);
    }
  });
  c.fillStyle = '#475569'; c.font = '14px Arial';
  c.fillText('Render otomatis dari KML/KMZ (skematik — citra satelit tidak tersedia)', 14, H - 16);
  c.strokeStyle = '#cbd5e1'; c.lineWidth = 2; c.strokeRect(2, 2, W - 4, H - 44);
  fs.writeFileSync(outPath, cv.toBuffer('image/png'));
}

// Simpan/perbarui SITE NAME + SITE ID sebuah slot ABD tanpa mengganti berkas.
async function setAtpSite(lokasiId, slot, body) {
  await siap();
  if (!validSlotAtp(slot) || tipeSlotAtp(slot) !== 'abd') throw new Error('Slot ABD tidak valid.');
  const nama = String(body.site_name || '').trim().slice(0, 120);
  const sid = String(body.site_id || '').trim().slice(0, 60);
  const old = await db.get('SELECT id FROM atp_foto WHERE lokasi_id=? AND slot=?', [lokasiId, slot]);
  if (!old) throw new Error('Unggah berkas ABD terlebih dahulu.');
  await db.run('UPDATE atp_foto SET site_name=?,site_id=? WHERE id=?', [nama, sid, old.id]);
  return getAtpData(lokasiId);
}

async function deleteAtpPhoto(lokasiId, slot) {
  await siap();
  if (!validSlotAtp(slot)) throw new Error('Slot foto ATP tidak valid.');
  const old = await db.get('SELECT file_name FROM atp_foto WHERE lokasi_id=? AND slot=?', [lokasiId, slot]);
  if (!old) throw new Error('Foto tidak ditemukan.');
  await db.run('DELETE FROM atp_foto WHERE lokasi_id=? AND slot=?', [lokasiId, slot]);
  try {
    const dir = atpDir(lokasiId);
    const re = new RegExp('^' + slot + '_\\d+\\.[a-z0-9]+$', 'i');
    for (const n of fs.readdirSync(dir)) if (re.test(n)) fs.rmSync(path.join(dir, n), { force: true });
  } catch (_) { /* data DB sudah terhapus */ }
  return true;
}

async function generateAtp(lokasiId) {
  await siap();
  if (!fs.existsSync(TEMPLATE_ATP)) throw new Error('Template ATP MITRATEL tidak ditemukan.');
  const all = await getAtpData(lokasiId);
  if (!all) throw new Error('Lokasi tidak ditemukan.');
  if (String(all.lokasi.tematik || '').toLowerCase() !== 'm-tel') throw new Error('Report ini hanya untuk lokasi dengan tematik M-tel.');
  const d = normalizeData(all.data);
  const fotoMap = {};
  for (const f of all.fotos) fotoMap[f.slot] = path.join(atpDir(lokasiId), f.file_name);
  const bySlotFoto = {};
  for (const f of all.fotos) bySlotFoto[f.slot] = f;
  // Daftar sel tujuan: teks placeholder dikosongkan, lalu foto (bila ada) disisipkan
  // ke area merge sel tersebut (atau range eksplisit, mis. sheet Lay Out ABD).
  const pins = Object.entries(SLOT_ATP).map(([slot, v]) => ({ sheet: v.sheet, cell: v.cell, range: v.range || '', img: fotoMap[slot] || '' }));
  // Teks pendukung: nama file KMZ asli + SITE NAME/ID & tabel ABD (nilai kosong
  // tetap ditulis agar contoh milik template tidak ikut terbawa).
  const teks = TEKS_ATP.map(t => {
    const f = bySlotFoto[t.slot];
    const namaFile = f ? String(f.original_name || f.file_name || '') : '';
    let value = '';
    if (t.jenis === 'file') value = namaFile;
    else if (t.jenis === 'literal') value = String(t.value || '');
    else if (t.jenis === 'flptitle') value = 'CONECTION CORE TO CORE';
    else if (t.jenis === 'spanid') value = t.prefix + (f ? String(f.site_id || '') : '');
    else if (t.jenis === 'spanname') value = t.prefix + (f ? String(f.site_name || '') : '');
    else if (t.jenis === 'sitename') value = t.prefix + (f ? String(f.site_name || '') : '');
    else if (t.jenis === 'siteid') value = t.prefix + (f ? String(f.site_id || '') : '');
    else if (t.jenis === 'tid') value = f ? String(f.site_id || '') : '';
    else if (t.jenis === 'tnama') value = f ? String(f.site_name || '') : '';
    else if (t.jenis === 'tfile') value = namaFile;
    return { sheet: t.sheet, cell: t.cell, value };
  });
  const outDir = path.join(UPLOAD_ROOT, safeName(lokasiId), 'report');
  fs.mkdirSync(outDir, { recursive: true });
  const base = 'ATP_MITRATEL_' + safeName(all.lokasi.nama_lokasi).slice(0, 80) + '_' + Date.now();
  const outXlsx = path.join(outDir, base + '.xlsx');
  const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'mitratel-atp-'));
  const cfgPath = path.join(tmp, 'config.json');
    // Nilai per core: Optical Length (H), Att. A-B (L), Att. B-A (P) — hanya port terisi.
  const meta = {};
  for (const m of (all.meta || [])) {
    meta[String(m.core)] = {
      optical_length: String(m.optical_length || '').trim(),
      att_ab: String(m.att_ab || '').trim(),
      att_ba: String(m.att_ba || '').trim()
    };
  }
  fs.writeFileSync(cfgPath, JSON.stringify({ template: TEMPLATE_ATP, outXlsx, tags: tags(d), pins, teks, hapus: HAPUS_GAMBAR, meta }), 'utf8');

  const script = [
    "$ErrorActionPreference='Stop'", "$cfg=Get-Content -Raw -LiteralPath " + q(cfgPath) + " | ConvertFrom-Json",
    "$xl=New-Object -ComObject Excel.Application", "$xl.Visible=$false", "$xl.DisplayAlerts=$false", "$wb=$null",
    "try {", "  Copy-Item -LiteralPath $cfg.template -Destination $cfg.outXlsx -Force",
    "  $wb=$xl.Workbooks.Open($cfg.outXlsx)",
    // Identitas (Cover Isat + Test Result Core) — teks {..} diganti dari data lokasi
    "  foreach($ws in $wb.Worksheets){ foreach($p in $cfg.tags.PSObject.Properties){ [void]$ws.UsedRange.Replace([string]$p.Name,[string]$p.Value) } }",
    // Hapus gambar contoh milik template pada sheet Lay Out ABD (tidak boleh bocor)
    "  foreach($d in $cfg.hapus){ $ws=$wb.Worksheets.Item([string]$d.sheet); foreach($sh in @($ws.Shapes)){ if([string]$sh.Name -eq [string]$d.name){ [void]$sh.Delete() } } }",
    // Tiap slot foto: kosongkan placeholder, lalu sisipkan gambar pada area merge
    // (atau range eksplisit bila sel tujuan tidak merge — sheet Lay Out ABD).
    "  foreach($p in $cfg.pins){",
    "    $ws=$wb.Worksheets.Item([string]$p.sheet)",
    "    $rng=$ws.Range([string]$p.cell)",
    "    $area=$null",
    "    if($p.range){ $area=$ws.Range([string]$p.range) } else { $area=$rng.MergeArea }",
    "    $rng.Value2=''",
    "    $img=[string]$p.img",
    "    if($img -and (Test-Path -LiteralPath $img)){",
    "      $pad=3; $w=[Math]::Max(10,$area.Width-2*$pad); $h=[Math]::Max(10,$area.Height-2*$pad)",
    "      [void]$ws.Shapes.AddPicture($img,$false,$true,$area.Left+$pad,$area.Top+$pad,$w,$h)",
    "    }",
    "  }",
    // Tulis teks pendukung SETELAH clearing placeholder/pin: nama file KMZ asli,
    // SITE NAME/ID + tabel ABD, dan SPAN ID/NAME + judul FLP Schematic Core.
    "  foreach($t in $cfg.teks){ $ws=$wb.Worksheets.Item([string]$t.sheet); $ws.Range([string]$t.cell).Value2=[string]$t.value }",
    "  $xl.CalculateFull()",
    // ── Optical Length (H), Att. A-B (L), Att. B-A (P) per core ──────────────
    // Baris core pada template: 34 = core 1 … 47 = core 14 (header di baris 33).
    // Kosongkan dulu ketiga kolom, lalu tulis hanya port yang punya nilai.
    "  $flp=$null; foreach($w in $wb.Worksheets){ if([string]$w.Name.Trim() -eq 'Test Result Core (FLP)'){ $flp=$w; break } }",
    "  if($null -eq $flp){ throw 'Sheet Test Result Core (FLP) tidak ditemukan.' }",
    "  for($i=1;$i -le 14;$i++){ $r=33+$i; foreach($c in @('H','L','P')){ $flp.Range($c+$r).Value2='' } }",
    "  foreach($p in $cfg.meta.PSObject.Properties){ $i=0; if(-not [int]::TryParse([string]$p.Name,[ref]$i)){ continue }; if($i -lt 1 -or $i -gt 14){ continue }; $r=33+$i; $m=$p.Value; $isi=@{ 'H'=[string]$m.optical_length; 'L'=[string]$m.att_ab; 'P'=[string]$m.att_ba }; foreach($c in @('H','L','P')){ $raw=([string]$isi[$c]).Trim(); if($raw -eq ''){ continue }; $s=$raw.Replace(',','.'); $n=0.0; if([double]::TryParse($s,[Globalization.NumberStyles]::Float,[Globalization.CultureInfo]::InvariantCulture,[ref]$n)){ $flp.Range($c+$r).Value2=$n } else { $flp.Range($c+$r).Value2=$raw } } }",
    "  $wb.Worksheets.Item('OTDR').Activate()",
    "  $wb.Save()", "} finally { if($wb){$wb.Close($true)}; $xl.Quit(); [Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null }",
    // Pasca simpan: paksa tampilan Normal + buka di sheet OTDR (sama seperti OPM,
    // COM di mesin ini adalah WPS sehingga Window.View ditolak → sunting XML langsung).
    "try {",
    "  Add-Type -AssemblyName System.IO.Compression.FileSystem",
    "  Add-Type -AssemblyName System.IO.Compression",
    "  $zip=[System.IO.Compression.ZipFile]::Open($cfg.outXlsx,[System.IO.Compression.ZipArchiveMode]::Update)",
    "  $entWb=$zip.GetEntry('xl/workbook.xml')",
    "  if($entWb){ $sr=New-Object System.IO.StreamReader($entWb.Open()); $txt=$sr.ReadToEnd(); $sr.Dispose()",
    "    if($txt -match 'activeTab='){ $txt=$txt -replace 'activeTab=\"\\d+\"','activeTab=\"2\"' } else { $txt=$txt -replace '<workbookView ','<workbookView activeTab=\"2\" ' }",
    "    $entWb.Delete(); $newWb=$zip.CreateEntry('xl/workbook.xml',[System.IO.Compression.CompressionLevel]::Optimal)",
    "    $sw=New-Object System.IO.StreamWriter($newWb.Open()); $sw.Write($txt); $sw.Dispose() }",
    "  foreach($ent in @($zip.Entries | Where-Object { $_.FullName -like 'xl/worksheets/sheet*.xml' })){",
    "    $name=$ent.FullName",
    "    $sr=New-Object System.IO.StreamReader($ent.Open()); $txt=$sr.ReadToEnd(); $sr.Dispose()",
    "    $txt=$txt -replace ' view=\"pageLayout\"','' -replace ' view=\"pageBreakPreview\"','' -replace ' tabSelected=\"1\"',''",
    "    if($name -eq 'xl/worksheets/sheet3.xml'){ $txt=$txt -replace '<sheetView(\\s|>)','<sheetView tabSelected=\"1\"$1'; $txt=$txt -replace ' topLeftCell=\"[^\"]*\"','' -replace 'zoomScale=\"\\d+\"','zoomScale=\"85\"' -replace 'zoomScaleNormal=\"\\d+\"','zoomScaleNormal=\"100\"' -replace '<selection[^>]*/>','<selection activeCell=\"A1\" sqref=\"A1\"/>' }",
    "    $ent.Delete(); $newE=$zip.CreateEntry($name,[System.IO.Compression.CompressionLevel]::Optimal)",
    "    $sw=New-Object System.IO.StreamWriter($newE.Open()); $sw.Write($txt); $sw.Dispose() }",
    "  $zip.Dispose()",
    "} catch { Write-Warning ('pasca-proses tampilan gagal: ' + $_.Exception.Message) }"
  ].join('\n');
  try { await ps(script, 300000); } finally { try { fs.rmSync(tmp, { recursive: true, force: true }); } catch (_) {} }
  if (!fs.existsSync(outXlsx)) throw new Error('Excel ATP gagal dibuat. Pastikan Microsoft Excel terpasang.');
  const terisi = Object.keys(fotoMap).length;
  return { xlsx: path.basename(outXlsx), pdf: '', baseUrl: '/api/mitratel/download/' + lokasiId + '/', foto: terisi, total: Object.keys(SLOT_ATP).length };
}

// Helper dari server.js (mis. renderPdf untuk ABD) — dipasang lewat register().
let helper = {};

// ─────────────────────────────────────────────────────────────────────────────
// LAPORAN M-TEL — rekap lokasi yang sudah input OPM & ATP + nominal (nilai DRM)
// ─────────────────────────────────────────────────────────────────────────────
const TOTAL_SLOT_OPM = 31;                 // 3 kalibrasi + 14 core NE + 14 core FE
function detailLokasi(json) {
  try { const d = JSON.parse(json || '{}'); return (d && typeof d === 'object') ? d : {}; }
  catch (_) { return {}; }
}
function uang(v) {
  const n = Number(String(v == null ? '' : v).replace(/[^\d.-]/g, ''));
  return Number.isFinite(n) && n > 0 ? n : 0;
}
async function laporanMtel() {
  await siap();
  const daftar = await db.all(
    "SELECT id,pid,nama_lokasi,no_sp,status_pekerjaan,detail_tambahan FROM lokasi WHERE lower(tematik)=? ORDER BY lower(nama_lokasi)",
    ['m-tel']);
  const opmRows = await db.all(
    "SELECT lokasi_id, COUNT(*) foto, SUM(CASE WHEN TRIM(COALESCE(hasil_scan,''))<>'' THEN 1 ELSE 0 END) scan FROM mitratel_foto GROUP BY lokasi_id");
  const atpRows = await db.all('SELECT lokasi_id, COUNT(*) foto FROM atp_foto GROUP BY lokasi_id');
  const coreRows = await db.all('SELECT lokasi_id, COUNT(*) core FROM atp_meta GROUP BY lokasi_id');
  const formRows = await db.all('SELECT lokasi_id FROM mitratel_data');
  const opm = new Map(opmRows.map(r => [Number(r.lokasi_id), { foto: Number(r.foto || 0), scan: Number(r.scan || 0) }]));
  const atp = new Map(atpRows.map(r => [Number(r.lokasi_id), Number(r.foto || 0)]));
  const core = new Map(coreRows.map(r => [Number(r.lokasi_id), Number(r.core || 0)]));
  const form = new Set(formRows.map(r => Number(r.lokasi_id)));
  const totalSlotAtp = Object.keys(SLOT_ATP).length;

  const rows = daftar.map(l => {
    const id = Number(l.id);
    const d = detailLokasi(l.detail_tambahan);
    const o = opm.get(id) || { foto: 0, scan: 0 };
    const fotoAtp = atp.get(id) || 0;
    const coreAtp = core.get(id) || 0;
    const adaForm = form.has(id);
    const sudahOpm = o.foto > 0 || adaForm || o.scan > 0;
    const sudahAtp = fotoAtp > 0 || coreAtp > 0;
    return {
      id,
      pid: String(l.pid || ''),
      nama: String(l.nama_lokasi || ''),
      sp: String(l.no_sp || d.no_sp || ''),
      status_pekerjaan: String(l.status_pekerjaan || ''),
      status_berkas: String(d.status_berkas || ''),
      witel: String(d.branch || ''),
      program: String(d.program || ''),
      pm: String(d.pm || ''),
      nominal: uang(d.nilai_total_drm),
      opm: { foto: o.foto, scan: o.scan, form: adaForm ? 1 : 0, sudah: sudahOpm ? 1 : 0 },
      atp: { foto: fotoAtp, core: coreAtp, sudah: sudahAtp ? 1 : 0 }
    };
  });

  // Rekap per Witel (branch)
  const perWitel = {};
  for (const r of rows) {
    const w = r.witel || 'TANPA WITEL';
    if (!perWitel[w]) perWitel[w] = { witel: w, lokasi: 0, opm: 0, atp: 0, lengkap: 0, nominal: 0, nominal_opm: 0, nominal_atp: 0 };
    const x = perWitel[w];
    x.lokasi++;
    x.nominal += r.nominal;
    if (r.opm.sudah) { x.opm++; x.nominal_opm += r.nominal; }
    if (r.atp.sudah) { x.atp++; x.nominal_atp += r.nominal; }
    if (r.opm.sudah && r.atp.sudah) x.lengkap++;
  }

  const sudahOpm = rows.filter(r => r.opm.sudah).length;
  const sudahAtp = rows.filter(r => r.atp.sudah).length;
  const lengkap = rows.filter(r => r.opm.sudah && r.atp.sudah).length;
  const belum = rows.length - rows.filter(r => r.opm.sudah || r.atp.sudah).length;
  const hitung = (pred) => rows.filter(pred).reduce((a, r) => a + r.nominal, 0);

  const ringkas = {
    total: rows.length,
    opm: sudahOpm,
    atp: sudahAtp,
    lengkap,
    belum,
    total_slot_opm: TOTAL_SLOT_OPM,
    total_slot_atp: totalSlotAtp,
    nominal_total: rows.reduce((a, r) => a + r.nominal, 0),
    nominal_opm: hitung(r => r.opm.sudah),
    nominal_atp: hitung(r => r.atp.sudah),
    nominal_lengkap: hitung(r => r.opm.sudah && r.atp.sudah),
    nominal_belum: hitung(r => !r.opm.sudah && !r.atp.sudah)
  };
  const witel = Object.values(perWitel).sort((a, b) => b.lokasi - a.lokasi);
  return { rows, ringkas, witel };
}

function register(app, checkAuth, h) {
  helper = h || {};
  const admin = checkAuth('admin');
  app.get('/mitratel', admin, (req, res) => res.sendFile(path.join(ROOT, 'public', 'mitratel.html')));
  app.get('/mitratel/atp', admin, (req, res) => res.sendFile(path.join(ROOT, 'public', 'mitratel_atp.html')));
  app.get('/mitratel/report', admin, (req, res) => res.sendFile(path.join(ROOT, 'public', 'mitratel_report.html')));
  app.get('/api/mitratel/report', admin, async (req, res) => {
    try { res.json(await laporanMtel()); } catch (e) { res.status(500).json({ message: e.message }); }
  });
  app.get('/api/mitratel/lokasi', admin, async (req, res) => {
    try { await siap(); const qv = String(req.query.q || '').trim(); const rows = await db.all("SELECT l.id,l.nama_lokasi,l.pid,l.tematik,md.updated_at FROM lokasi l LEFT JOIN mitratel_data md ON md.lokasi_id=l.id WHERE lower(l.tematik)=? " + (qv ? "AND lower(l.nama_lokasi) LIKE ? " : '') + 'ORDER BY l.nama_lokasi LIMIT 500', qv ? ['m-tel', '%' + qv.toLowerCase() + '%'] : ['m-tel']); res.json({ rows }); } catch(e) { res.status(500).json({ message:e.message }); }
  });
  app.get('/api/mitratel/data/:lokasiId', admin, async (req, res) => { const a=await getData(req.params.lokasiId); if(!a)return res.status(404).json({message:'Lokasi tidak ditemukan.'}); res.json(a); });
  app.post('/api/mitratel/data/:lokasiId', admin, async (req,res) => { try { res.json({success:true, ...(await saveData(req.params.lokasiId, req.body))}); } catch(e) { res.status(500).json({message:e.message}); } });
  app.post('/api/mitratel/upload/:lokasiId/:slot', admin, upload.single('file'), async (req,res) => { try { res.json({success:true, ...(await savePhoto(req.params.lokasiId,req.params.slot,req.file))}); } catch(e) { res.status(400).json({message:e.message}); } });
  app.post('/api/mitratel/tag/:lokasiId/:slot', admin, async (req,res) => { try { const t=await updateTag(req.params.lokasiId,req.params.slot,req.body.tag); res.json({success:true, tag:t}); } catch(e) { res.status(400).json({message:e.message}); } });
  app.delete('/api/mitratel/foto/:lokasiId/:slot', admin, async (req,res) => { try { await deletePhoto(req.params.lokasiId, req.params.slot); res.json({success:true}); } catch(e) { res.status(400).json({message:e.message}); } });
  app.post('/api/mitratel/scan-kalibrasi/:lokasiId', admin, async (req,res) => { try { const { hasil, gagal } = await scanCalibration(req.params.lokasiId); res.json({success:true, hasil, gagal}); } catch(e) { res.status(500).json({message:e.message}); } });
  app.post('/api/mitratel/scan-opm/:lokasiId', admin, async (req,res) => { try { const { hasil, gagal } = await scanNeFe(req.params.lokasiId);
  app.post('/api/mitratel/update-scan/:lokasiId/:slot', admin, async (req,res) => { try { await siap(); const slot=req.params.slot; if(!validSlot(slot)) throw new Error('Slot tidak valid.'); const val=String((req.body&&req.body.value)||'').trim(); await db.run('UPDATE mitratel_foto SET hasil_scan=? WHERE lokasi_id=? AND slot=?', [val, req.params.lokasiId, slot]); res.json({success:true}); } catch(e){ res.status(500).json({message:e.message}); } }); res.json({success:true, hasil, gagal}); } catch(e) { res.status(500).json({message:e.message}); } });
  app.post('/api/mitratel/generate/otp/:lokasiId', admin, async (req,res) => { try { res.json({success:true, ...(await generateOtp(req.params.lokasiId))}); } catch(e) { res.status(500).json({message:e.message}); } });
  app.get('/api/mitratel/foto/:lokasiId/:file', admin, (req,res) => res.sendFile(path.join(UPLOAD_ROOT,safeName(req.params.lokasiId),safeName(req.params.file))));
  app.get('/api/mitratel/download/:lokasiId/:file', admin, (req,res) => res.download(path.join(UPLOAD_ROOT,safeName(req.params.lokasiId),'report',safeName(req.params.file))));
  // ── ATP MITRATEL ────────────────────────────────────────────────────────────
  app.get('/api/mitratel/atp/slots', admin, (req,res) => res.json({ slots: Object.entries(SLOT_ATP).map(([slot,v])=>({slot,...v})) }));
  app.get('/api/mitratel/atp/data/:lokasiId', admin, async (req,res) => { const a=await getAtpData(req.params.lokasiId); if(!a)return res.status(404).json({message:'Lokasi tidak ditemukan.'}); res.json(a); });
  app.post('/api/mitratel/atp/upload/:lokasiId/:slot', admin, upload.single('file'), async (req,res) => { try { res.json({success:true, ...(await saveAtpPhoto(req.params.lokasiId,req.params.slot,req.file,{site_name:req.body.site_name,site_id:req.body.site_id,keep_asli:String(req.body.keep_asli||'')==='1'}))}); } catch(e) { res.status(400).json({message:e.message}); } });
  app.post('/api/mitratel/atp/site/:lokasiId/:slot', admin, async (req,res) => { try { res.json({success:true, ...(await setAtpSite(req.params.lokasiId,req.params.slot,req.body||{}))}); } catch(e) { res.status(400).json({message:e.message}); } });
  app.post('/api/mitratel/atp/meta/:lokasiId/:slot', admin, async (req,res) => { try { await siap(); const slot=req.params.slot; if(!validSlotAtp(slot)) throw new Error('Slot tidak valid.'); const ol=String((req.body&&req.body.optical_length)||'').trim(); const a=String((req.body&&req.body.att_ab)||'').trim(); const b=String((req.body&&req.body.att_ba)||'').trim(); const old=await db.get('SELECT id FROM atp_foto WHERE lokasi_id=? AND slot=?',[req.params.lokasiId,slot]); if(!old) throw new Error('Unggah foto terlebih dahulu.'); await db.run('UPDATE atp_foto SET optical_length=?,att_ab=?,att_ba=? WHERE id=?',[ol,a,b,old.id]); res.json({success:true, ...(await getAtpData(req.params.lokasiId))}); } catch(e){ res.status(400).json({message:e.message}); } });
  // Nilai Optical Length + redaman per core (tab "Foto OTDR") — tidak butuh foto.
  app.post('/api/mitratel/atp/core-meta/:lokasiId', admin, async (req,res) => { try { await siap(); const meta = await setAtpCoreMeta(req.params.lokasiId, req.body||{}); res.json({success:true, meta, ...(await getAtpData(req.params.lokasiId))}); } catch(e){ res.status(400).json({message:e.message}); } });
 // Ambil gambar dari URL (drag gambar dari WhatsApp/Web yang diblokir CORS).
 app.post('/api/mitratel/atp/fetch-url', admin, async (req,res) => {
   try {
     const u = String((req.body||{}).url || '').trim();
     if (!urlGambarAman(u)) throw new Error('URL gambar tidak diizinkan.');
     const ctrl = new AbortController();
     const to = setTimeout(() => ctrl.abort(), 15000);
     let r;
     try {
       r = await fetch(u, { signal: ctrl.signal, redirect: 'follow', headers: { 'User-Agent': 'Mozilla/5.0 (compatible; ATP-Mitratel/1.0)' } });
     } finally { clearTimeout(to); }
     if (!r.ok) throw new Error('Gagal mengambil gambar (HTTP ' + r.status + ').');
     if (!urlGambarAman(r.url)) throw new Error('URL akhir gambar tidak diizinkan.');
     const ct = String(r.headers.get('content-type') || '').split(';')[0].trim().toLowerCase();
     if (!/^image\//.test(ct)) throw new Error('Tautan bukan gambar (' + (ct || 'tanpa tipe') + ').');
     const buf = Buffer.from(await r.arrayBuffer());
     if (!buf.length) throw new Error('Gambar kosong.');
     if (buf.length > 15 * 1024 * 1024) throw new Error('Gambar lebih dari 15 MB.');
     res.set('Content-Type', ct);
     res.set('Cache-Control', 'no-store');
     res.send(buf);
   } catch (e) {
     res.status(400).json({ message: e && e.name === 'AbortError' ? 'Pengambilan gambar melebihi 15 detik.' : (e.message || 'Gagal mengambil gambar.') });
   }
 });
  app.delete('/api/mitratel/atp/foto/:lokasiId/:slot', admin, async (req,res) => { try { await deleteAtpPhoto(req.params.lokasiId, req.params.slot); res.json({success:true}); } catch(e) { res.status(400).json({message:e.message}); } });
  app.post('/api/mitratel/atp/generate/:lokasiId', admin, async (req,res) => { try { res.json({success:true, ...(await generateAtp(req.params.lokasiId))}); } catch(e) { res.status(500).json({message:e.message}); } });
  app.get('/api/mitratel/atp/foto/:lokasiId/:file', admin, (req,res) => res.sendFile(path.join(atpDir(req.params.lokasiId),safeName(req.params.file))));
}
module.exports = { register };
