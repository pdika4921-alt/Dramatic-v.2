const db = require('./db-wrapper.js');

// ─────────────────────────────────────────────────────────────────────────────
// SUMBER DATA MASTER
// Data diambil dari spreadsheet publik "Semesta OnHand & Revenue 2026 Area-2",
// tab "Data Master" (gid 0):
//   • kolom E  = Rev                    → Internal / Eksternal (dipakai di
//                 widget Rekap Status)
//   • kolom G  = PROGRAM                → tematik project
//   • kolom AF = Status Lapangan PMO-2  → status lapangan; baris berstatus
//                 "FINISH INSTALASI / OA" (dan "GO LIVE") dipakai sebagai
//                 indikator proyek yang masuk tahap evidentiary.
// Header sheet berada di baris 4, kolom E/G/AF dibaca lewat nama header
// sehingga aman bila kolom digeser.
// ─────────────────────────────────────────────────────────────────────────────
const PUBLISHED_ID = '2PACX-1vRX2voGHf6Sptp4zoh0XBMRtXTgTdUS1pksC7B4bbsU_m9MNphX7-LmKoIVYk173oaWiia2ZUc1_Veg';
const PUB_BASE = 'https://docs.google.com/spreadsheets/d/e/' + PUBLISHED_ID;
const DATA_SHEET_NAME = 'Data Master';
const SHEETS = [{ nama: DATA_SHEET_NAME, gid: 0 }];

// Tab "Dasboard 2026" pada menu Report masih memakai spreadsheet lama.
const SPREADSHEET_ID = '1wwJ4jBCpEQWKV5J9T7jF3HqCZa-QD5jNezRVsMKQhw8';

// Jadwal sinkronisasi. Sheet Data Master ~3 MB CSV, jadi default 5 menit
// (bisa diubah lewat env GODC_SYNC_INTERVAL_MS).
const SYNC_INTERVAL_MS = Number(process.env.GODC_SYNC_INTERVAL_MS) || 300000;

// Pembersihan lokasi lama hanya berjalan bila baris yang terbaca utuh
// (file CSV yang terpotong akibat sinkronisasi akan menghapus data master).
// Bypass manual: POST /api/godc/cleanup?force=1
const MIN_ROWS_FOR_CLEANUP = Number(process.env.GODC_MIN_ROWS_CLEANUP) || 5000;
const AUTO_CLEANUP = process.env.GODC_AUTO_CLEANUP !== '0';
// Berisiko menghapus baris, jadi bisa dimatikan dengan GODC_SWEEP_CHILDREN=0
const SWEEP_CHILDREN = process.env.GODC_SWEEP_CHILDREN !== '0';

const FIELDS = [
  'no_urut', 'sumber_data', 'sumber_rev', 'nama_lop', 'tematik_project', 'project',
  'ket_project', 'nama_mitra', 'branch', 'id_sap', 'id_pid', 'kat_status',
  'status_fisik', 'status_progress', 'nilai_total_drm', 'po_non_po', 'no_sp',
  'status_berkas', 'pic_berkas', 'keterangan',
  // Kolom tidak ada di sheet ini — disimpan kosong agar schema & query lama tetap jalan.
  'waspang', 'material_drm', 'jasa_drm'
];

const HEADER_MAP = {
  no_urut: 'NO',
  sumber_data: 'SOURCE DATA',
  sumber_rev: 'REV',
  nama_lop: 'LOKASI',
  tematik_project: 'PROGRAM',
  project: 'PROJECT',
  ket_project: 'SUBPROJECT',
  nama_mitra: 'PM',
  branch: 'BRANCH',
  id_sap: 'ID',
  id_pid: 'PID',
  kat_status: 'KAT STATUS',
  status_fisik: 'STATUS FISIK',
  status_progress: 'STATUS LAPANGAN PMO-2',
  nilai_total_drm: 'NILAI TELKOM',
  po_non_po: 'PO/WO',
  no_sp: 'NO SP TELKOM',
  status_berkas: 'STATUS BERKAS PMO-2',
  pic_berkas: 'PIC BERKAS PMO-2',
  keterangan: 'KETERANGAN',
  waspang: 'WASPANG',
  material_drm: 'MATERIAL DRM',
  jasa_drm: 'JASA DRM'
};

// Tematik PROGRAM (kolom G) → tematik aplikasi
const TEMATIK_APP_MAP = [
  { pola: /HEM/i, hasil: 'HEM' },
  { pola: /QEREL|RELOKASI|RELOLASI/i, hasil: 'Relokasi' },
  { pola: /MICRO\s*DEMAND/i, hasil: 'MICRODEMAND' },
  { pola: /NODE\s*B/i, hasil: 'Node B' },
  { pola: /MY\s*REP|MYREPUBLIK/i, hasil: 'MyRepublik' },
  { pola: /MITRATEL|MORATEL|M\s*-?\s*TEL/i, hasil: 'M-tel' },
  { pola: /(^|\W)JLM(\W|$)/i, hasil: 'Jlm' },
  { pola: /(^|\W)OLO(\W|$)/i, hasil: 'OLO' },
  { pola: /UNSC|(^|\W)PT\s*2(\W|$)/i, hasil: 'PT2 UNSC' },
  { pola: /IFORTE|FOR3E/i, hasil: 'IFORTE' }
];

// Status AF (Status Lapangan PMO-2) yang memicu "perlu input berkas eviden".
// Selain status final "FINISH INSTALASI / OA" dan "GO LIVE", tahap pengerjaan
// 03. MATDEV / 04. INSTALASI serta tahap closure 07. TC / 08. UT / 09. REKON
// juga dianggap butuh eviden; "00. DROP" dan "10. BAST" tidak.
// Pola tahap dipakukan ke awal teks supaya kata TC/UT di tengah kalimat
// tidak ikut cocok.
const TRIGGER_STATUS_RE = /(FINISH INSTALASI|GO\s*LIVE|^0?3\s*\.?\s*MATDEV\b|^0?4\s*\.?\s*INSTALASI\b|^0?7\s*\.?\s*TC\b|^0?8\s*\.?\s*UT\b|^0?9\s*\.?\s*REKON\b)/i;

// Nilai No. SP yang berarti "tidak ada nomor SP"
const SP_EMPTY_RE = /^(-{1,3}|n\/?a|tbd|belum|o?kp|kosong|0)$/i;

const state = {
  lastSync: null, lastError: null, syncing: false,
  total: 0, totalDrm: 0, rowsRead: 0, uniqueLokasi: 0, triggerCount: 0,
  lastHash: null, lastCreatedLokasi: 0, lastCreatedNames: [],
  createdAtLokasi: null, pidBackfilled: 0, cleanup: null,
  autoCreateDiagnostic: null
};

function clean(v) {
  return String(v == null ? '' : v).replace(/^\uFEFF/, '').replace(/\s+/g, ' ').trim();
}

function normLop(s) {
  return clean(s).toLowerCase();
}

// Normalisasi nama lokasi untuk pencocokan dedupe (huruf kecil, tanpa tanda baca)
function normName(s) {
  return clean(s).toLowerCase()
    .replace(/[\s.,;()&'"/\\-]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

// Normalisasi BRANCH. Sheet Data Master tidak konsisten kapital
// ("Bandung" vs "BANDUNG", "jakbar" vs "JAKBAR") dan memakai singkatan
// jakarta (JAKSEL) selain nama panjang (JAKARTA SELATAN). Tanpa ini filter
// BRANCH memecah satu cabang jadi beberapa pilihan.
const BRANCH_CANON = {
  JAKSEL: 'JAKARTA SELATAN',
  JAKTIM: 'JAKARTA TIMUR',
  JAKBAR: 'JAKARTA BARAT',
  JAKPUS: 'JAKARTA PUSAT',
  JAKUT: 'JAKARTA UTARA'
};

function normBranch(s) {
  const v = clean(s).toUpperCase().replace(/\s+/g, ' ').trim();
  if (!v) return '';
  return BRANCH_CANON[v] || v;
}

// Samakan bentuk singkatan jakarta juga pada sisi input filter, supaya
// ?branch=jaksel dan ?branch=JAKARTA SELATAN memilih cabang yang sama.
function normBranchKey(s) {
  return normBranch(s).replace(/\s+/g, '-');
}

// Peta tematik GODC → tematik aplikasi
function mapTematik(t) {
  const v = clean(t);
  for (const m of TEMATIK_APP_MAP) if (m.pola.test(v)) return m.hasil;
  return v;
}

// Apakah status proyek memicu "perlu input berkas eviden"
function isTriggerStatus(st) {
  return TRIGGER_STATUS_RE.test(clean(st));
}

// Tahap status (angka di depan, mis. "05. FINISH INSTALASI / OA" → 5)
function stageOf(st) {
  const m = clean(st).match(/^(\d+)\s*\./);
  return m ? parseInt(m[1], 10) : -1;
}

// Kelas PO / NON PO untuk pivot & daftar SP.
// Kolom "PO/WO": WO (Work Order) tetap dihitung PO karena aplikasi hanya
// punya 2 kategori (PO / NON PO).
function poClass(v) {
  const s = clean(v).toUpperCase();
  if (s === 'PO' || s === 'WO') return 'PO';
  if (s === 'NON PO' || s === 'NONPO' || s === 'NON-PO') return 'NON PO';
  return '—';
}

// Buang placeholder ("-", "N/A", dst) dari nomor SP
function cleanSp(v) {
  const s = clean(v);
  if (!s) return '';
  if (SP_EMPTY_RE.test(s.replace(/\s+/g, ''))) return '';
  return s;
}

function numFromDrm(v) {
  const s = String(v == null ? '' : v).replace(/[^\d.,-]/g, '').replace(/[,.]/g, '');
  const n = parseFloat(s);
  return isNaN(n) ? 0 : n;
}

function csvParse(text) {
  const rows = [];
  let row = [], field = '', inQ = false;
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (inQ) {
      if (c === '"') {
        if (text[i + 1] === '"') { field += '"'; i++; }
        else inQ = false;
      } else field += c;
    } else if (c === '"') {
      inQ = true;
    } else if (c === ',') {
      row.push(field); field = '';
    } else if (c === '\r') {
      // abaikan (dialihkan ke \n)
    } else if (c === '\n') {
      row.push(field); field = '';
      rows.push(row); row = [];
    } else {
      field += c;
    }
  }
  if (field.length > 0 || row.length > 0) { row.push(field); rows.push(row); }
  return rows;
}

function findHeader(rows) {
  for (let i = 0; i < rows.length; i++) {
    if (rows[i].some(c => clean(c).toUpperCase() === 'LOKASI')) return i;
  }
  throw new Error('Header "LOKASI" tidak ditemukan di sheet ' + DATA_SHEET_NAME);
}

// Peta kolom: nama field → indeks kolom pada baris header (case-insensitive)
function colMap(headerRow) {
  const upper = headerRow.map(c => clean(c).toUpperCase());
  const map = {};
  for (const key of FIELDS) {
    const label = HEADER_MAP[key];
    const idx = label ? upper.indexOf(label) : -1;
    map[key] = idx >= 0 ? idx : null;
  }
  return map;
}

function rowsToProjects(rows, headerIdx, map) {
  const out = [];
  for (let i = headerIdx + 1; i < rows.length; i++) {
    const r = rows[i];
    const namaLop = clean(r[map.nama_lop]);
    if (!namaLop) continue;
    const p = { nama_lop: namaLop, _lopKey: normLop(namaLop), _nameKey: normName(namaLop) };
    for (const key of FIELDS) {
      const idx = map[key];
      p[key] = (idx != null && r[idx] != null) ? clean(r[idx]) : '';
    }
    p.sumber = p.sumber_data || 'DATA MASTER';
    // Kolom E "Rev" berisi "Internal"/"Eksternal" (ejaan sheet;"Sheets juga
    // pernah menulis "External"), dinormalkan agar konsisten di Rekap Status.
    const rev = String(p.sumber_rev || '').trim().toUpperCase();
    p.sumber_rev = /^INT/.test(rev) ? 'INTERNAL' : (/^E/.test(rev) ? 'EKSTERNAL' : '');
    p.no_sp = cleanSp(p.no_sp);
    p._stage = stageOf(p.status_progress);
    p.butuh_eviden = triggerFlag(p.status_progress);
    p._drm = numFromDrm(p.nilai_total_drm);
    out.push(p);
  }
  return out;
}

// Nama lokasi bisa muncul berkali-kali di Data Master (satu lokasi, beberapa
// sub-proyek/SP). Baris tetap disimpan semua supaya data master utuh, hanya
// penyajiannya yang digabung per lokasi (lihat uniqByLokasi).
function triggerFlag(status) {
  return isTriggerStatus(status) ? 'YA' : 'TIDAK';
}

// Gabungkan baris kembar jadi satu unit kerja: baris berstatus paling maju
// dipakai untuk ditampilkan, tetapi flag "butuh eviden" di-OR sehingga lokasi
// yang salah satu barisnya FINISH INSTALASI tetap masuk daftar evidentiary.
function uniqByLokasi(items) {
  const byKey = new Map();
  for (const it of items) {
    const key = normName(it.nama_lop);
    const prev = byKey.get(key);
    if (!prev) {
      it.butuh_eviden = it.butuh_eviden || triggerFlag(it.status_progress);
      byKey.set(key, it);
      continue;
    }
    const win = stageOf(it.status_progress) > stageOf(prev.status_progress) ? it : prev;
    win.butuh_eviden = (it.butuh_eviden === 'YA' || prev.butuh_eviden === 'YA') ? 'YA' : triggerFlag(win.status_progress);
    byKey.set(key, win);
  }
  return Array.from(byKey.values());
}

function mergeLists(lists) {
  const merged = [];
  for (const l of lists) merged.push(...l);
  return merged;
}

// CSV publik: /pub?gid=..&single=true&output=csv (endpoint /export tidak bisa
// dipakai untuk spreadsheet yang hanya dibagikan lewat "publish to web").
function csvUrl(gid) {
  return `${PUB_BASE}/pub?gid=${Number(gid) || 0}&single=true&output=csv`;
}

function sheetUrl(gid) {
  return `${PUB_BASE}/pubhtml/sheet?headers=false&gid=${Number(gid) || 0}`;
}

async function readSheets(fetcher) {
  const fn = fetcher || (gid => fetch(csvUrl(gid)));
  const lists = [];
  let rowsRead = 0;
  for (const s of SHEETS) {
    const res = await fn(s.gid);
    if (!res.ok) {
      const snippet = (await res.text()).substring(0, 200);
      throw new Error(`Gagal fetch sheet ${s.nama} (gid=${s.gid}): HTTP ${res.status} — ${snippet}`);
    }
    const ct = res.headers && res.headers.get ? (res.headers.get('content-type') || '') : '';
    const text = typeof res === 'string' ? res : await res.text();
    if (ct.includes('text/html') || text.trim().startsWith('<!DOCTYPE') || text.trim().startsWith('<html')) {
      throw new Error(`Sheet ${s.nama} mengembalikan HTML (kemungkinan redirect login). Pastikan spreadsheet di-share publik.`);
    }
    const rows = csvParse(text);
    const headerIdx = findHeader(rows);
    const map = colMap(rows[headerIdx]);
    const projects = rowsToProjects(rows, headerIdx, map);
    rowsRead += projects.length;
    lists.push(projects);
  }
  const merged = mergeLists(lists);
  const totalDrm = merged.reduce((a, p) => a + p._drm, 0);
  const triggers = merged.filter(p => p.butuh_eviden === 'YA');
  return {
    rowsRead: merged.length,
    total: merged.length,
    uniqueLokasi: new Set(merged.map(p => p._nameKey)).size,
    totalDrm,
    triggerCount: triggers.length,
    triggerLokasi: new Set(triggers.map(p => p._nameKey)).size,
    merged
  };
}

// Report dashboard (menu Report) — spreadsheet lama, tab "Dasboard 2026".
async function readReport(gid) {
  const res = await fetch(`https://docs.google.com/spreadsheets/d/${SPREADSHEET_ID}/export?format=csv&gid=${Number(gid) || 0}`);
  const text = await res.text();
  return csvParse(text).map(r => r.map(c => String(c)));
}

const GODC_COLUMNS = [
  'no_urut', 'sumber', 'sumber_data', 'sumber_rev', 'nama_lop', 'tematik_project', 'project',
  'ket_project', 'nama_mitra', 'branch', 'id_sap', 'pid', 'kat_status',
  'status_fisik', 'status_progress', 'butuh_eviden', 'nilai_total_drm',
  'po_non_po', 'no_sp', 'status_berkas', 'pic_berkas', 'keterangan',
  'waspang', 'material_drm', 'jasa_drm', 'drm_nilai', 'updated_at'
];

async function ensureSchema() {
  await db.run(`CREATE TABLE IF NOT EXISTS godc_project (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    sumber TEXT,
    nama_lop TEXT,
    tematik_project TEXT,
    waspang TEXT,
    status_progress TEXT,
    ket_project TEXT,
    nama_mitra TEXT,
    id_sap TEXT,
    pid TEXT,
    material_drm TEXT,
    jasa_drm TEXT,
    nilai_total_drm TEXT,
    po_non_po TEXT,
    drm_nilai REAL DEFAULT 0,
    updated_at TEXT
  )`);
  const extra = {
    no_urut: 'TEXT', sumber_data: 'TEXT', sumber_rev: 'TEXT', project: 'TEXT', branch: 'TEXT',
    kat_status: 'TEXT', status_fisik: 'TEXT', no_sp: 'TEXT',
    status_berkas: 'TEXT', pic_berkas: 'TEXT', keterangan: 'TEXT', pid: 'TEXT',
    butuh_eviden: 'TEXT'
  };
  for (const col of Object.keys(extra)) {
    try { await db.run(`ALTER TABLE godc_project ADD COLUMN ${col} ${extra[col]}`); } catch (e) { /* kolom sudah ada */ }
  }
  await db.run(`CREATE TABLE IF NOT EXISTS godc_meta (key TEXT PRIMARY KEY, value TEXT)`);
  await db.exec(`CREATE TABLE IF NOT EXISTS lokasi (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    pid TEXT,
    nama_lokasi TEXT UNIQUE,
    tematik TEXT,
    detail_tambahan TEXT,
    status_pekerjaan TEXT
  )`);
  for (const col of ['status_pekerjaan', 'no_sp']) {
    try { await db.run(`ALTER TABLE lokasi ADD COLUMN ${col} TEXT`); } catch (e) { /* kolom sudah ada */ }
  }
  await db.run(`CREATE TABLE IF NOT EXISTS godc_sp (
    nama_lokasi TEXT PRIMARY KEY,
    no_sp TEXT,
    updated_at TEXT
  )`);
  try { await db.run(`ALTER TABLE godc_sp ADD COLUMN source TEXT`); } catch (e) { /* kolom sudah ada */ }
}

function autoCreateLokasi(merged) {
  const tgl = new Date().toISOString();
  return new Promise(async (resolve) => {
    try {
      const all = await db.all('SELECT id, nama_lokasi, pid FROM lokasi');
      const ada = new Set(all.map(r => normName(r.nama_lokasi)));
      // Satu lokasi = satu unit kerja: baris kembar hanya dibuat satu kali.
      const allTriggers = uniqByLokasi(merged.filter(p => p.butuh_eviden === 'YA'));
      const seen = new Set();
      const triggers = [];
      const skippedNames = [];
      for (const p of allTriggers) {
        const key = p._nameKey;
        if (ada.has(key) || seen.has(key)) { skippedNames.push(clean(p.nama_lop)); continue; }
        seen.add(key);
        triggers.push(p);
      }
      const made = [];
      await db.batch(async () => {
        for (const p of triggers) {
          const namaLok = clean(p.nama_lop);
          const tem = mapTematik(p.tematik_project);
          const detail = JSON.stringify({
            sumber: 'GODC', sheet: DATA_SHEET_NAME, sumber_godc: p.sumber,
            id_sap: p.id_sap, program: p.tematik_project, branch: p.branch,
            pm: p.nama_mitra, no_sp: p.no_sp, status_berkas: p.status_berkas,
            pic_berkas: p.pic_berkas, nilai_total_drm: p.nilai_total_drm,
            godc_status: p.status_progress, godc_updated: tgl
          });
          // status_pekerjaan diisi status AF saat itu — dipakai filter
          // "Status Pekerjaan" pada widget Rekap Status.
          await db.run('INSERT INTO lokasi (nama_lokasi, tematik, pid, detail_tambahan, status_pekerjaan, no_sp) VALUES (?,?,?,?,?,?)',
            [namaLok, tem, (p.id_pid || p.id_sap || ''), detail, clean(p.status_progress), p.no_sp || '']);
          made.push(namaLok);
        }
      });
      let backfilled = 0;
      await db.batch(async () => {
        const lokAll = await db.all('SELECT id, nama_lokasi, pid FROM lokasi');
        const byKey = new Map(merged.map(p => [p._nameKey, p]));
        for (const row of lokAll) {
          const p = byKey.get(normName(row.nama_lokasi));
          if (!p) continue;
          const newPid = p.id_pid || p.id_sap || '';
          if ((row.pid || '') !== newPid) {
            await db.run('UPDATE lokasi SET pid = ? WHERE id = ?', [newPid, row.id]);
            backfilled++;
          }
        }
      });
      state.lastCreatedLokasi = made.length;
      // Hanya sampel nama yang dikembalikan lewat API; daftar penuh bisa ratusan baris.
      state.lastCreatedNames = made.slice(0, 25);
      state.createdAtLokasi = made.length ? tgl : null;
      state.pidBackfilled = backfilled;
      // Dua angka berbeda sengaja dipisah: jumlah BARIS trigger (mentah) dan
      // jumlah LOKASI unik yang harus punya berkas.
      state.triggerCount = merged.filter(p => p.butuh_eviden === 'YA').length;
      state.triggerLokasi = allTriggers.length;
      state.autoCreateDiagnostic = {
        sheet: DATA_SHEET_NAME,
        totalGODC: merged.length,
        allTriggerCount: allTriggers.length,
        alreadyInLokasi: skippedNames.length,
        newlyCreated: made.length,
        skippedNames: skippedNames.slice(0, 25),
        createdNames: made.slice(0, 25)
      };
      resolve();
    } catch (e) {
      state.lastError = state.lastError || 'autoCreateLokasi: ' + (e.message || e);
      resolve();
    }
  });
}

// Simpan No. SP untuk sebuah proyek (disimpan di godc_sp + ditulis juga ke lokasi yang cocok)
async function saveSp(namaLokasi, noSp) {
  const name = clean(namaLokasi);
  const now = new Date().toISOString();
  await db.run(
    `INSERT INTO godc_sp (nama_lokasi, no_sp, updated_at, source)
     VALUES (?,?,?,'MANUAL')
     ON CONFLICT(nama_lokasi) DO UPDATE SET no_sp = excluded.no_sp, updated_at = excluded.updated_at, source = 'MANUAL'`,
    [name, noSp, now]
  );
  const all = await db.all('SELECT id, nama_lokasi FROM lokasi');
  const hit = all.find(r => normName(r.nama_lokasi) === normName(name));
  if (hit) await db.run('UPDATE lokasi SET no_sp = ? WHERE id = ?', [noSp, hit.id]);
  return { lokasi_id: hit ? hit.id : null };
}

// Tarik No. SP dari kolom "NO SP TELKOM" (kolom U). Input manual di aplikasi
// (source='MANUAL') selalu menang atas nilai dari sheet.
async function syncSpFromSheet(merged) {
  const now = new Date().toISOString();
  const rows = uniqByLokasi(merged.filter(p => p.butuh_eviden === 'YA')).filter(p => p.no_sp);
  let filled = 0;
  await db.batch(async () => {
    for (const p of rows) {
      const res = await db.run(
        `INSERT INTO godc_sp (nama_lokasi, no_sp, updated_at, source)
         VALUES (?,?,?,'SHEET')
         ON CONFLICT(nama_lokasi) DO UPDATE SET no_sp = excluded.no_sp, updated_at = excluded.updated_at
         WHERE godc_sp.source IS NULL OR godc_sp.source <> 'MANUAL'`,
        [clean(p.nama_lop), p.no_sp, now]
      );
      if (res && res.changes) filled++;
    }
  });
  return filled;
}

async function getSpMap() {
  const rows = await db.all('SELECT nama_lokasi, no_sp FROM godc_sp');
  const m = new Map();
  for (const r of rows) if (r.no_sp) m.set(normName(r.nama_lokasi), r.no_sp);
  return m;
}

// Hapus lokasi yang namanya tidak ada lagi di Data Master (beserta seluruh
// data turunannya: berkas/eviden, RAB, rekon, ODP/OPM, BOQ, dsb).
const CHUNK = 400;

async function cleanupOrphanLokasi(rows, opts) {
  const options = opts || {};
  const force = !!options.force;
  const dryRun = !!options.dryRun;
  const all = await db.all('SELECT id, nama_lokasi FROM lokasi');
  const keep = new Set(rows.map(r => r._nameKey || normName(r.nama_lop)));
  const orphans = all.filter(r => !keep.has(normName(r.nama_lokasi)));
  const info = {
    rowsRead: rows.length,
    totalLokasi: all.length,
    orphanCount: orphans.length,
    deleted: false,
    dryRun,
    reason: ''
  };
  if (!orphans.length) {
    state.cleanup = Object.assign(info, { deletedCount: 0 });
    return state.cleanup;
  }
  if (!force && rows.length < MIN_ROWS_FOR_CLEANUP) {
    info.reason = `Dilewati: baris terbaca (${rows.length}) < batas aman ${MIN_ROWS_FOR_CLEANUP}. Paksa via POST /api/godc/cleanup?force=1`;
    state.cleanup = info;
    return state.cleanup;
  }
  info.orphanNames = orphans.map(r => r.nama_lokasi).slice(0, 200);
  if (dryRun) { state.cleanup = info; return info; }

  const ids = orphans.map(r => r.id);
  const tables = [];
  await db.batch(async () => {
    const list = await db.all("SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%'");
    for (const t of list.map(x => x.name)) {
      if (t === 'lokasi' || t === 'godc_sp' || t === 'godc_project' || t === 'godc_meta') continue;
      const cols = await db.all(`PRAGMA table_info(${t})`);
      if (!cols.some(c => c.name === 'lokasi_id')) continue;
      let deleted = 0;
      for (let i = 0; i < ids.length; i += CHUNK) {
        const part = ids.slice(i, i + CHUNK);
        const res = await db.run(
          `DELETE FROM ${t} WHERE lokasi_id IN (${part.map(() => '?').join(',')})`, part);
        deleted += (res && res.changes) || 0;
      }
      if (deleted) tables.push({ table: t, deleted });
    }
    // Baris No. SP milik lokasi yang dihapus
    for (let i = 0; i < orphans.length; i += CHUNK) {
      const part = orphans.slice(i, i + CHUNK).map(r => r.nama_lokasi);
      await db.run(
        `DELETE FROM godc_sp WHERE nama_lokasi IN (${part.map(() => '?').join(',')})`, part);
    }
    let deletedLokasi = 0;
    for (let i = 0; i < ids.length; i += CHUNK) {
      const part = ids.slice(i, i + CHUNK);
      const res = await db.run(`DELETE FROM lokasi WHERE id IN (${part.map(() => '?').join(',')})`, part);
      deletedLokasi += (res && res.changes) || 0;
    }
    info.deletedCount = deletedLokasi;
    info.tables = tables;
  });
  info.deleted = true;
  info.totalRowsDeleted = tables.reduce((a, x) => a + x.deleted, 0);
  state.cleanup = info;
  return info;
}

// Bersihkan baris tabel anak yang menunjuk lokasi yang sudah tidak ada lagi
// (yatim). Baris ini tidak bisa dibuka dari UI karena induknya hilang.
async function sweepOrphanChildren() {
  const list = await db.all("SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%'");
  const out = [];
  await db.batch(async () => {
    for (const t of list.map(x => x.name)) {
      if (t === 'lokasi') continue;
      const cols = await db.all(`PRAGMA table_info(${t})`);
      if (!cols.some(c => c.name === 'lokasi_id')) continue;
      const res = await db.run(
        `DELETE FROM ${t} WHERE lokasi_id IS NOT NULL AND lokasi_id NOT IN (SELECT id FROM lokasi)`);
      if (res && res.changes) out.push({ table: t, deleted: res.changes });
    }
  });
  return out;
}

// Lihat saja (tanpa menghapus) baris tabel anak yang menunjuk lokasi yang
// sudah tidak ada lagi, supaya UI bisa memberi tahu jumlahannya.
async function countOrphanChildren() {
  const list = await db.all("SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%'");
  const out = [];
  for (const t of list.map(x => x.name)) {
    if (t === 'lokasi') continue;
    const cols = await db.all(`PRAGMA table_info(${t})`);
    if (!cols.some(c => c.name === 'lokasi_id')) continue;
    const r = await db.get(
      `SELECT COUNT(*) n FROM ${t} WHERE lokasi_id IS NOT NULL AND lokasi_id NOT IN (SELECT id FROM lokasi)`);
    if (r && r.n > 0) out.push({ table: t, rows: r.n });
  }
  return out;
}

async function persist(stats, merged) {
  const now = new Date().toISOString();
  await db.batch(async () => {
    await db.run('DELETE FROM godc_project');
    const cols = GODC_COLUMNS.join(', ');
    const qs = GODC_COLUMNS.map(() => '?').join(', ');
    const st = db.prepare(`INSERT INTO godc_project (${cols}) VALUES (${qs})`);
    for (const p of merged) {
      await st.run(
        String(p.no_urut || ''), p.sumber, p.sumber_data || '', p.sumber_rev || '', p.nama_lop, p.tematik_project, p.project || '',
        p.ket_project, p.nama_mitra, p.branch || '', p.id_sap, p.id_pid || '', p.kat_status || '',
        p.status_fisik || '', p.status_progress, p.butuh_eviden || 'TIDAK', p.nilai_total_drm, p.po_non_po,
        cleanSp(p.no_sp), p.status_berkas || '', p.pic_berkas || '', p.keterangan || '',
        p.waspang || '', p.material_drm || '', p.jasa_drm || '', p._drm || 0, now
      );
    }
    await st.finalize();
  });
  await db.run(`INSERT OR REPLACE INTO godc_meta (key, value) VALUES ('last_sync', ?)`, [now]);
  await db.run(`INSERT OR REPLACE INTO godc_meta (key, value) VALUES ('rows_read', ?)`, [String(stats.rowsRead)]);
  await db.run(`INSERT OR REPLACE INTO godc_meta (key, value) VALUES ('total', ?)`, [String(stats.total)]);
  await db.run(`INSERT OR REPLACE INTO godc_meta (key, value) VALUES ('total_drm', ?)`, [String(stats.totalDrm)]);
  await db.run(`INSERT OR REPLACE INTO godc_meta (key, value) VALUES ('trigger_rows', ?)`, [String(stats.triggerCount)]);
  await db.run(`INSERT OR REPLACE INTO godc_meta (key, value) VALUES ('trigger_lokasi', ?)`, [String(stats.triggerLokasi || 0)]);
  // Kunci meta dari sumber lama (INTERNAL/EXTERNAL) sudah tidak relevan.
  await db.run(`DELETE FROM godc_meta WHERE key IN ('internal','external')`);
}

function hashProjects(merged) {
  return JSON.stringify(merged.map(p => ({
    n: p.nama_lop, t: p.tematik_project, s: p.status_progress, o: p.sumber_data,
    pm: p.nama_mitra, br: p.branch, id: p.id_sap, pid: p.id_pid,
    nd: p.nilai_total_drm, pn: p.po_non_po, sp: p.no_sp,
    sb: p.status_berkas, pb: p.pic_berkas
  })));
}
async function runSync(opts) {
  const options = opts || {};
  if (state.syncing) return state;
  state.syncing = true;
  state.lastError = null;
  state.autoCreateDiagnostic = null;
  try {
    const stats = await readSheets();
    const hash = hashProjects(stats.merged);
    await ensureSchema();
    state.rowsRead = stats.rowsRead;
    state.uniqueLokasi = stats.uniqueLokasi;
    if (state.lastHash !== hash) {
      await persist(stats, stats.merged);
      state.lastHash = hash;
      state.total = stats.total;
      state.totalDrm = stats.totalDrm;
      state.triggerCount = stats.triggerCount;
    }
    await syncSpFromSheet(stats.merged);
    await autoCreateLokasi(stats.merged);
    if (options.cleanup || AUTO_CLEANUP) {
      await cleanupOrphanLokasi(stats.merged, { force: !!options.cleanup });
      // Sisa-sisa lama: baris tabel anak yang menunjuk lokasi yang sudah dihapus
      // jauh sebelum sinkronisasi ini. Tidak tampil di UI karena induknya hilang.
      if (SWEEP_CHILDREN) {
        const swept = await sweepOrphanChildren();
        if (swept.length) {
          state.cleanup = Object.assign({}, state.cleanup, { sweptChildren: swept });
        }
      }
    }
    await db.run(`INSERT OR REPLACE INTO godc_meta (key, value) VALUES ('created_lokasi', ?)`, [String(state.lastCreatedLokasi)]);
    state.lastSync = new Date().toISOString();
  } catch (e) {
    state.lastError = String(e && e.message || e);
  } finally {
    state.syncing = false;
  }
  return state;
}

let _intervalId = null;

function start(intervalMs) {
  if (_intervalId) return;
  runSync().catch(e => console.error('[GODC] Initial sync error:', e.message || e));
  _intervalId = setInterval(() => {
    runSync().catch(e => console.error('[GODC] Sync error:', e.message || e));
  }, intervalMs || SYNC_INTERVAL_MS);
}

function stop() {
  if (_intervalId) { clearInterval(_intervalId); _intervalId = null; }
}

function isRunning() { return _intervalId !== null; }

module.exports = {
  PUBLISHED_ID, PUB_BASE, DATA_SHEET_NAME, SHEETS, SPREADSHEET_ID, FIELDS,
  SYNC_INTERVAL_MS, MIN_ROWS_FOR_CLEANUP, AUTO_CLEANUP, SWEEP_CHILDREN,
  state, csvParse, clean, cleanSp, normLop, normName, normBranch, normBranchKey, mapTematik, isTriggerStatus,
  stageOf, triggerFlag, poClass, numFromDrm, findHeader, colMap, rowsToProjects,
  uniqByLokasi, csvUrl, sheetUrl, readSheets, readReport, ensureSchema, persist,
  autoCreateLokasi, syncSpFromSheet, saveSp, getSpMap, cleanupOrphanLokasi,
  sweepOrphanChildren, countOrphanChildren, runSync, start, stop, isRunning
};
