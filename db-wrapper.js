const initSqlJs = require('sql.js');
const fs = require('fs');
const path = require('path');

let _db = null;
const DB_PATH = process.env.DB_PATH || path.join(__dirname, 'database.db');

async function getDb() {
  if (_db) return _db;
  fs.mkdirSync(path.dirname(DB_PATH), { recursive: true });
  const SQL = await initSqlJs();
  if (fs.existsSync(DB_PATH)) {
    const fileBuffer = fs.readFileSync(DB_PATH);
    _db = new SQL.Database(fileBuffer);
  } else {
    _db = new SQL.Database();
  }
  return _db;
}

// Saat batch sedang berjalan, penulisan file ditunda sampai batch selesai
// (satu batch = satu kali tulis, bukan satu kali tulis per statement).
let deferDepth = 0;

let _pernahWarnUkuran = false;

function saveDb() {
  if (!_db || deferDepth > 0) return;
  const data = Buffer.from(_db.export());

  // Alarm: memori jauh lebih kecil dari file yang ada → kemungkinan salinan basi
  // (proses lain sedang memegang database). Satu instance server saja yang boleh jalan.
  try {
    if (fs.existsSync(DB_PATH)) {
      const lama = fs.statSync(DB_PATH).size;
      if (lama > 1048576 && data.length < lama * 0.5 && !_pernahWarnUkuran) {
        _pernahWarnUkuran = true;
        console.warn(`⚠ saveDb: database di memori (${data.length} B) jauh lebih kecil dari file (${lama} B).`
          + ' Kemungkinan ada instance server lain / salinan basi. Pastikan hanya SATU "node server.js" yang berjalan.');
      }
    }
  } catch (_) { /* abaikan */ }

  // Tulis ATOMIK: file sementara dengan nama UNIK (pid + timestamp) agar tidak
  // berkolisi dengan instance lain, lalu rename dengan retry bila target terkunci.
  const tmp = DB_PATH + '.tmp.' + process.pid + '.' + Date.now();
  try {
    fs.writeFileSync(tmp, data);
  } catch (e) {
    console.error('saveDb gagal menulis database (write tmp):', e.message);
    return;
  }

  let ok = false;
  for (let i = 0; i < 5 && !ok; i++) {
    try {
      fs.renameSync(tmp, DB_PATH);
      ok = true;
    } catch (e) {
      if (i < 4) {
        // tunggu singkat bila target terkunci
        const t0 = Date.now(); while (Date.now() - t0 < 50) { /* busy wait */ }
      } else {
        // fallback bila rename ditolak
        try {
          fs.copyFileSync(tmp, DB_PATH);
          ok = true;
        } catch (e2) {
          console.error('saveDb gagal menulis database:', e2.message);
        }
      }
    }
  }
  try { if (fs.existsSync(tmp)) fs.unlinkSync(tmp); } catch (_) { /* abaikan */ }
}

class AsyncDatabase {
  async exec(sql) {
    const db = await getDb();
    db.run(sql);
    saveDb();
  }
  async run(sql, params = []) {
    const db = await getDb();
    const p = Array.isArray(params) ? params : [params];
    db.run(sql, p);
    // getRowsModified() harus dibaca SEBELUM export(): ekspor file membatalkan
    // counter "baris terubah" sehingga nilainya selalu 0 bila dibaca sesudahnya.
    const changes = db.getRowsModified();
    saveDb();
    return { changes: changes };
  }
  async get(sql, params = []) {
    const db = await getDb();
    const p = Array.isArray(params) ? params : [params];
    const stmt = db.prepare(sql);
    stmt.bind(p);
    if (stmt.step()) {
      const row = stmt.getAsObject();
      stmt.free();
      return row;
    }
    stmt.free();
    return undefined;
  }
  async all(sql, params = []) {
    const db = await getDb();
    const p = Array.isArray(params) ? params : [params];
    const results = [];
    const stmt = db.prepare(sql);
    stmt.bind(p);
    while (stmt.step()) {
      results.push(stmt.getAsObject());
    }
    stmt.free();
    return results;
  }
  prepare(sql) {
    return {
      run: async (...p) => {
        const db = await getDb();
        db.run(sql, p.flat());
        const changes = db.getRowsModified();
        saveDb();
        return { changes: changes };
      },
      finalize: () => Promise.resolve()
    };
  }

  // Jalankan banyak语句 dalam satu transaksi: satu kali tulis file di akhir,
  // dan seluruh perubahan batal otomatis bila ada statement yang gagal.
  // Wajib untuk sinkronisasi ribuan baris (mis. 11.267 proyek Data Master).
  async batch(fn) {
    const db = await getDb();
    deferDepth++;
    let result;
    try {
      db.run('BEGIN');
      result = await fn();
      db.run('COMMIT');
    } catch (e) {
      try { db.run('ROLLBACK'); } catch (_) { /* abaikan */ }
      throw e;
    } finally {
      deferDepth--;
    }
    if (deferDepth === 0) saveDb();
    return result;
  }
}

module.exports = new AsyncDatabase();