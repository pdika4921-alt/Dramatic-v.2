// ─────────────────────────────────────────────────────────────────────────────
// REKAP OPM RELOKASI
//
// Dua tugas:
//   1. bacaAngkaOpm()  — baca otomatis angka hasil ukur (dBm) dari FOTO eviden
//      OPM memakai AI vision (Gemini/Claude, sama seperti verifikasi keaslian).
//   2. isiRekapExcel() — isi template Excel "REKAP OPM RELOKASI.xls" dari data
//      ODP + nilai tiap port. Blok ODP di-loop (digandakan bila ODP > 2) dan
//      tinggi blok mengikuti jumlah port per ODP (kelipatan 8 baris: 8 port →
//      8 baris, 16 port → 16 baris). RUMUS di dalam template DIPERTAHANKAN
//      (hanya sel placeholder {..} dan kolom nilai (V) yang diubah).
//      Hasilnya: satu .xlsx + satu PDF per ODP (dipakai server untuk membuat
//      gambar 1 halaman per ODP).
//
// Excel dijalankan lewat COM (Excel.Application) — mesin ini sudah memakai COM
// yang sama untuk MS Word (konversi PDF), jadi tidak ada dependensi baru.
// ─────────────────────────────────────────────────────────────────────────────
const fs = require('fs');
const os = require('os');
const path = require('path');
const { execFile } = require('child_process');
const sharp = require('sharp');

const cfg = require('./verifyConfig');

/* Antrean scan OPM: beberapa ODP/port boleh dipindai bersamaan, tetapi jumlah
   panggilan AI yang aktif dibatasi supaya tidak kena rate-limit API. */
let scanBerjalan = 0;
const antreanScan = [];
function tiketScan() {
  return new Promise(resolve => { antreanScan.push(resolve); lanjutScan(); });
}
function lanjutScan() {
  const maks = Math.max(1, Number(cfg.ai.maxConcurrent) || 3);
  while (scanBerjalan < maks && antreanScan.length) {
    scanBerjalan++;
    antreanScan.shift()();
  }
}
function lepasScan() { scanBerjalan--; lanjutScan(); }

const MASTER_XLS = path.join(__dirname, '..', 'templates', 'REKAP OPM RELOKASI.xls');

// Baris & kolom penting pada template REKAP OPM RELOKASI.xls (sheet DETAIL REKAP REDAMAN)
const BARIS_AWAL_BLOK = 21;   // baris pertama blok ODP ke-1
const TINGGI_BLOK     = 8;    // tinggi BLOK TEMPLATE (8 baris = 8 port);
                              // blok 16 port dibangun dari 2 potong blok template
const KOL_NAMA_ODP    = 6;    // kolom F
const KOL_NILAI       = 22;   // kolom V (Hasil Ukur Redaman ODP (OUT))
const KOL_AKHIR_CETAK = 'Y';  // kolom terakhir yang berisi data (B..Y)

/* ───────────────────────────── OCR angka OPM ───────────────────────────── */

function promptSistemOpm() {
  return [
    'Anda membaca satu FOTO hasil pengukuran redaman fiber optik (OPM/OTDR) di Indonesia.',
    'Angka hasil ukur biasanya tampil sebagai angka besar di LAYAR alat OPM (sering bergaya 7-segmen) atau di layar aplikasi/screenshot, bersatuan dBm atau dB — contoh -15.01, -12,34, atau 0.35.',
    'Periksa foto dengan TELITI sebelum menjawab: perhatikan bagian layar/display, bukan tulisan lain.',
    'Tugas: ambil HANYA satu angka hasil ukur (dB/dBm) yang paling jelas/dominan.',
    'Abaikan angka panjang/desimal tidak masuk akal, nomor seri, tanggal, waktu, atau satuan lain (nm, km, suhu).',
    'Abaikan SEMUA angka pada watermark GPS/koordinat, teks lokasi/alamat, peta, dan nama jalan — itu bukan hasil ukur.',
    'Jika foto sama sekali tidak memuat layar/tampilan alat ukur (mis. hanya gedung, tiang, ODP, kabel, atau foto biasa), kembalikan "angka" kosong.',
    'Balas HANYA dengan JSON valid tanpa teks lain, format: {"angka":"-15.01","keterangan":"<=15 kata"}',
    'Gunakan titik sebagai pemisah desimal di dalam field "angka"; jangan pakai tanda dBm.'
  ].join(' ');
}

function promptUserOpm(konteks) {
  const k = konteks || {};
  // Foto kalibrasi (MITRATEL): pembacaan alat untuk kolom EQUIPMENT CALIBRATION
  // First/Second/Third (dB) — bisa berupa power meter atau OPM.
  if (k.jenis === 'kalibrasi') {
    return [
      'Foto ini adalah pembacaan KALIBRASI alat untuk laporan OPM MITRATEL.',
      'Bacalah SATU nilai redaman/kalibrasi (bersatuan dB, boleh positif/negatif, mis. 0.35 atau -1.20)',
      'yang tertera pada layar/tampilan alat ukur (power meter / OPM).',
      'Slot foto: ' + (k.port || '-'),
      'Jika foto TIDAK memuat layar alat ukur (mis. hanya gedung, tiang, ODP, kabel, peta, atau watermark GPS),',
      'maka kembalikan "angka" kosong — jangan mengarang dan jangan membaca angka watermark/koordinat.',
      'Balas JSON sesuai format.'
    ].join('\n');
  }
  if (k.jenis === 'opm_test') {
    return [
      'Foto ini adalah pembacaan OPM untuk laporan OPM MITRATEL (Tabel OPM Test).',
      'Bacalah SATU nilai redaman (dB) yang tertera pada layar/tampilan alat ukur (OPM/power meter).',
      'Slot foto: ' + (k.port || '-') + ' (port ' + String(k.port || '').replace(/[^0-9]/g, '') + ')',
      'Jika foto TIDAK memuat layar alat ukur (mis. hanya gedung, tiang, ODP, kabel, peta, atau watermark GPS),',
      'maka kembalikan "angka" kosong — jangan mengarang dan jangan membaca angka watermark/koordinat.',
      'Balas JSON sesuai format.'
    ].join('\n');
  }
  const bagian = ['Bacakan angka hasil ukur redaman (dBm) dari foto OPM ini.'];
  if (k.odp) bagian.push('Titik/ODP: ' + k.odp);
  if (k.port) bagian.push('Port: ' + k.port);
  bagian.push('Balas JSON sesuai format.');
  return bagian.join('\n');
}

async function gambarBase64(filePath) {
  const buf = await sharp(filePath)
    .resize(cfg.ai.maxSisi, cfg.ai.maxSisi, { fit: 'inside', withoutEnlargement: true })
    .jpeg({ quality: 85 })
    .toBuffer();
  return buf.toString('base64');
}

async function panggilClaude(base64, sistem, teks, apiKey, model) {
  const r = await fetch('https://api.anthropic.com/v1/messages', {
    method: 'POST',
    headers: {
      'content-type': 'application/json',
      'x-api-key': apiKey,
      'anthropic-version': '2023-06-01'
    },
    body: JSON.stringify({
      model: model || 'claude-3-5-sonnet-latest',
      max_tokens: 300,
      system: sistem,
      messages: [{
        role: 'user',
        content: [
          { type: 'image', source: { type: 'base64', media_type: 'image/jpeg', data: base64 } },
          { type: 'text', text: teks }
        ]
      }]
    }),
    // Batas waktu per panggilan: tanpa ini, koneksi yang menggantung membuat
    // scan foto bisa hang tanpa batas (retry di bawah tidak pernah tercapai).
    signal: AbortSignal.timeout(cfg.ai.timeoutMs || 45000)
  });
  if (!r.ok) {
    const badan = await r.text().catch(() => '');
    const err = new Error('Claude HTTP ' + r.status);
    if (r.status === 429 || r.status === 529 || r.status >= 500) err.retryAfterMs = 8000;
    throw err;
  }
  const d = await r.json();
  return (d.content || []).map(c => c.text || '').join('');
}

async function panggilGemini(base64, sistem, teks, apiKey, model) {
  const namaModel = model || 'gemini-3.8-flash';
  const url = 'https://generativelanguage.googleapis.com/v1beta/models/' +
    encodeURIComponent(namaModel) + ':generateContent';
  const r = await fetch(url, {
    method: 'POST',
    headers: { 'content-type': 'application/json', 'x-goog-api-key': apiKey },
    body: JSON.stringify({
      system_instruction: { parts: [{ text: sistem }] },
      contents: [{
        role: 'user',
        parts: [
          { inline_data: { mime_type: 'image/jpeg', data: base64 } },
          { text: teks }
        ]
      }],
      generationConfig: { maxOutputTokens: 300, temperature: 0.1, responseMimeType: 'application/json' }
    }),
    // Batas waktu per panggilan (lihat catatan di panggilClaude).
    signal: AbortSignal.timeout(cfg.ai.timeoutMs || 45000)
  });
  if (!r.ok) {
    const badan = await r.text().catch(() => '');
    let pesan = '';
    try { pesan = (JSON.parse(badan).error || {}).message || ''; } catch (e) { /* bukan JSON */ }
    const err = new Error('Gemini HTTP ' + r.status + (pesan ? ': ' + pesan.slice(0, 140) : ''));
    const m = /"retryDelay":\s*"(\d+(?:\.\d+)?)s"/.exec(badan);
    if (m) err.retryAfterMs = Math.round(parseFloat(m[1]) * 1000);
    else if (r.status === 429 || r.status >= 500) err.retryAfterMs = 8000;
    throw err;
  }
  const d = await r.json();
  const bagian = (d.candidates && d.candidates[0] && d.candidates[0].content && d.candidates[0].content.parts) || [];
  return bagian.map(p => p.text || '').join('');
}

function parseJsonLonggar(teks) {
  const potong = String(teks || '').replace(/```(?:json)?/gi, '').trim();
  try { return JSON.parse(potong); } catch (e) { /* lanjut */ }
  const buka = potong.indexOf('{');
  const tutup = potong.lastIndexOf('}');
  if (buka >= 0 && tutup > buka) {
    try { return JSON.parse(potong.slice(buka, tutup + 1)); } catch (e) { /* lanjut */ }
  }
  return null;
}

function normalisasiAngka(nilai) {
  if (nilai == null) return '';
  let s = String(nilai).trim();
  if (!s) return '';
  const m = /[+-]?\d+(?:[.,]\d+)?/.exec(s);
  if (!m) return '';
  s = m[0].replace(',', '.');
  const n = Number(s);
  if (!Number.isFinite(n)) return '';
  return String(n);
}

/**
 * Baca angka hasil ukur (dBm) dari satu foto eviden OPM.
 * @returns {Promise<string>} angka (mis. "-15.01") atau '' bila tidak terbaca.
 */
async function bacaAngkaOpmInti(filePath, konteks = {}) {
  const provider = cfg.ai.provider === 'anthropic' ? 'claude' : cfg.ai.provider;
  if (provider === 'mock') {
    return '';
  }
  if (!cfg.ai.apiKey) throw new Error('AI_API_KEY belum diisi di .env — tidak bisa scan angka OPM.');
  const base64 = await gambarBase64(filePath);
  const sistem = promptSistemOpm();
  const teks   = promptUserOpm(konteks);
  const daftarModel = String(cfg.ai.model || '').split(',').map(s => s.trim()).filter(Boolean);
  let galat = null;
  let pernahKosong = false;
  for (let percobaan = 0; percobaan <= cfg.ai.retry; percobaan++) {
    const m = daftarModel.length ? daftarModel[percobaan % daftarModel.length] : '';
    try {
      let teksHasil;
      if (provider === 'claude') {
        teksHasil = await panggilClaude(base64, sistem, teks, cfg.ai.apiKey, m);
      } else if (provider === 'gemini') {
        teksHasil = await panggilGemini(base64, sistem, teks, cfg.ai.apiKey, m);
      } else {
        throw new Error('AI_PROVIDER tidak dikenal: ' + provider);
      }
      const j = parseJsonLonggar(teksHasil) || {};
      const angka = normalisasiAngka(j.angka);
      if (angka) return angka;
      // AI menjawab "tidak ada angka" — model vision kadang kosong padahal
      // fotonya berisi angka, jadi coba ulang beberapa kali sebelum menyerah.
      pernahKosong = true;
      if (percobaan < cfg.ai.retry) await new Promise(r => setTimeout(r, 250));
    } catch (e) {
      galat = e;
      if (percobaan < cfg.ai.retry) {
        // Jeda retry dipangkas: dulu bisa sampai 30 detik per percobaan sehingga
        // scan 1 ODP terasa lama. Sekarang maksimal 4 detik, masih cukup untuk
        // menghormati retryAfter dari API bila memang diberikan.
        const jeda = Math.min(4000, Math.max(250 * (percobaan + 1), (e && e.retryAfterMs) || 0));
        await new Promise(r => setTimeout(r, jeda));
      }
    }
  }
  if (pernahKosong) return '';
  throw galat || new Error('Gagal membaca angka OPM');
}

/* Versi publik: ambil tiket antrean dulu, lepas setelah selesai (sukses/gagal). */
async function bacaAngkaOpm(filePath, konteks = {}) {
  await tiketScan();
  try {
    return await bacaAngkaOpmInti(filePath, konteks);
  } finally {
    lepasScan();
  }
}

/* ───────────────────── Isi Excel + render PDF per ODP ───────────────────── */

function jalankanPs(perintah, timeoutMs) {
  return new Promise((resolve, reject) => {
    execFile('powershell.exe',
      ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-Command', perintah],
      { timeout: timeoutMs || 240000, maxBuffer: 16 * 1024 * 1024, windowsHide: true },
      (err, stdout, stderr) => {
        if (err) reject(new Error(String(stderr || err.message || '').trim().slice(0, 500) || 'Proses Excel gagal.'));
        else resolve(stdout);
      });
  });
}

function q(p) { return "'" + String(p).replace(/'/g, "''") + "'"; }

/**
 * Isi template Excel lalu ekspor 1 PDF per ODP.
 *
 * @param {object} opsi
 * @param {{proyek,surat_pesanan,lokasi_smile,catuan}} opsi.header
 * @param {Array<{nama:string,nilai:string[]}>} opsi.odps  nilai[0] = port 1, dst.
 * @param {string} opsi.outDir  folder tujuan (dibuat bila belum ada)
 * @param {string} [opsi.prefix]
 * @returns {Promise<{xlsx:string, pdfs:Array<{nama:string,path:string}>}>}
 */
async function isiRekapExcel({ header, odps, outDir, prefix }) {
  if (!fs.existsSync(MASTER_XLS)) throw new Error('Master Excel tidak ditemukan: ' + MASTER_XLS);
  if (!Array.isArray(odps) || !odps.length) throw new Error('Belum ada ODP untuk direkap.');
  if (!fs.existsSync(outDir)) fs.mkdirSync(outDir, { recursive: true });

  const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'rekapopm-'));
  const base = String(prefix || 'REKAP_OPM').replace(/[^a-zA-Z0-9_-]/g, '_');
  const workXls  = path.join(tmp, 'work.xls');
  const outXlsx  = path.join(outDir, base + '.xlsx');
  const configPath = path.join(tmp, 'config.json');

  const daftarPdf = odps.map((o, i) => ({
    nama: String(o.nama || ('ODP' + (i + 1))),
    path: path.join(outDir, base + '_' + String(i + 1).padStart(2, '0') + '.pdf')
  }));

  const konfig = {
    master: MASTER_XLS,
    workXls,
    outXlsx,
    header: {
      proyek: (header && header.proyek) || '-',
      surat_pesanan: (header && header.surat_pesanan) || '-',
      lokasi_smile: (header && header.lokasi_smile) || '-',
      catuan: (header && header.catuan) || '-'
    },
    odps: odps.map(o => ({
      nama: String(o.nama || ''),
      nilai: Array.isArray(o.nilai) ? o.nilai.map(v => (v == null ? '' : String(v))) : []
    })),
    pdfs: daftarPdf.map(p => p.path)
  };
  fs.writeFileSync(configPath, JSON.stringify(konfig), 'utf8');

  const ps = [
    "$ErrorActionPreference='Stop'",
    "$cfg = Get-Content -Raw -LiteralPath " + q(configPath) + " | ConvertFrom-Json",
    "$xl = New-Object -ComObject Excel.Application",
    "$xl.Visible = $false",
    "$xl.DisplayAlerts = $false",
    "$wb = $null",
    "try {",
    "  Copy-Item -LiteralPath $cfg.master -Destination $cfg.workXls -Force",
    "  $wb = $xl.Workbooks.Open($cfg.workXls)",
    "  $ws = $wb.Worksheets.Item(1)",
    "  $ur = $ws.UsedRange",
    "  [void]$ur.Replace('{proyek}', [string]$cfg.header.proyek)",
    "  [void]$ur.Replace('{surat pesanan}', [string]$cfg.header.surat_pesanan)",
    "  [void]$ur.Replace('{lokasi smile}', [string]$cfg.header.lokasi_smile)",
    "  [void]$ur.Replace('{catuan}', [string]$cfg.header.catuan)",
    "  $baseRow = " + BARIS_AWAL_BLOK,
    "  $step = " + TINGGI_BLOK,
    "  $all = @($cfg.odps)",
    "  $n = $all.Count",
    "  # Hitung tinggi + posisi tiap blok ODP. Tinggi blok = kelipatan 8 baris",
    "  # (8 port → 8 baris, 16 port → 16 baris = 2 potong 8 baris template).",
    "  $heights = @()",
    "  $starts = @()",
    "  $cur = $baseRow",
    "  for ($i = 0; $i -lt $n; $i++) {",
    "    $len = @($all[$i].nilai).Count",
    "    if ($len -lt 1) { $len = 1 }",
    "    $h = [int]($step * [Math]::Ceiling($len / [double]$step))",
    "    $heights += $h",
    "    $starts += $cur",
    "    $cur = $cur + $h",
    "  }",
    "  # Bangun blok baru: baris template cuma 2 blok (baris 21..28 & 29..36).",
    "  # Potong yang jatuh DI LUAR baris template ditempel dari blok template pertama",
    "  # (termasuk tinggi barisnya), potong yang berada DI DALAM baris template dibiarkan.",
    "  for ($i = 0; $i -lt $n; $i++) {",
    "    $start = $starts[$i]",
    "    $h = $heights[$i]",
    "    for ($off = 0; $off -lt $h; $off += $step) {",
    "      $t = $start + $off",
    "      if ($t -eq $baseRow -or $t -eq ($baseRow + $step)) { continue }",
    "      $src = $ws.Range('A' + $baseRow + ':AW' + ($baseRow + $step - 1))",
    "      $dst = $ws.Range('A' + $t + ':AW' + ($t + $step - 1))",
    "      $src.Copy($dst) | Out-Null",
    "      for ($k = 0; $k -lt $step; $k++) { $ws.Rows.Item($t + $k).RowHeight = $ws.Rows.Item($baseRow + $k).RowHeight }",
    "    }",
    "    # Bentuk ulang merge kolom E (feeder) & F (nama ODP) untuk SATU blok penuh.",
    "    # Template sudah merge 8 baris; kalau bloknya 16 port atau baris salinan",
    "    # berasal dari sel yang kebetulan sedang ter-merge sebagian, merge-nya",
    "    # jadi hilang — Merge() pada rentang yang sudah pas = no-op, jadi aman.",
    "    [void]$ws.Range('E' + $start + ':E' + ($start + $h - 1)).Merge()",
    "    [void]$ws.Range('F' + $start + ':F' + ($start + $h - 1)).Merge()",
    "  }",
    "  for ($i = 0; $i -lt $n; $i++) {",
    "    $start = $starts[$i]",
    "    $h = $heights[$i]",
    "    $odp = $all[$i]",
    "    $ws.Cells.Item($start, " + KOL_NAMA_ODP + ").Value2 = [string]$odp.nama",
    "    $vals = @($odp.nilai)",
    "    for ($p = 0; $p -lt $vals.Count; $p++) {",
    "      $row = $start + $p",
    "      $raw = [string]$vals[$p]",
    "      $cell = $ws.Cells.Item($row, " + KOL_NILAI + ")",
    "      if ($raw.Trim() -eq '') { [void]$cell.ClearContents() }",
    "      elseif ($raw -match '^[+-]?[0-9]+([.,][0-9]+)?$') {",
    "        $cell.NumberFormat = '0.00'",
    "        $num = [double]::Parse($raw.Replace(',', '.'), [Globalization.CultureInfo]::InvariantCulture)",
    "        $cell.Value2 = $num",
    "      } else {",
    "        $cell.NumberFormat = '@'",
    "        $cell.Value2 = [string]$raw",
    "      }",
    "    }",
    "    # Bersihkan baris sisa dalam blok bila jumlah nilai < tinggi blok",
    "    # (supaya placeholder template {ODP1P9} dst tidak ikut tercetak).",
    "    for ($p = $vals.Count; $p -lt $h; $p++) { [void]$ws.Cells.Item($start + $p, " + KOL_NILAI + ").ClearContents() }",
    "  }",
    "  $wb.SaveAs($cfg.outXlsx, 51) | Out-Null",
    "  $ws.PageSetup.Orientation = 2",
    "  $ws.PageSetup.Zoom = $false",
    "  $ws.PageSetup.FitToPagesWide = 1",
    "  $ws.PageSetup.FitToPagesTall = 1",
    "  $ws.PageSetup.LeftMargin = $xl.InchesToPoints(0.2)",
    "  $ws.PageSetup.RightMargin = $xl.InchesToPoints(0.2)",
    "  $ws.PageSetup.TopMargin = $xl.InchesToPoints(0.25)",
    "  $ws.PageSetup.BottomMargin = $xl.InchesToPoints(0.25)",
    "  for ($i = 0; $i -lt $n; $i++) {",
    "    $start = $starts[$i]",
    "    $end = $start + $heights[$i] - 1",
    "    for ($j = 0; $j -lt $n; $j++) {",
    "      $s2 = $starts[$j]",
    "      $e2 = $s2 + $heights[$j] - 1",
    "      $ws.Range(($s2.ToString() + ':' + $e2.ToString())).EntireRow.Hidden = ($j -ne $i)",
    "    }",
    "    $ws.PageSetup.PrintArea = 'B6:" + KOL_AKHIR_CETAK + "' + $end",
    "    $ws.ExportAsFixedFormat(0, [string]$cfg.pdfs[$i])",
    "  }",
    "  for ($j = 0; $j -lt $n; $j++) {",
    "    $s2 = $starts[$j]",
    "    $e2 = $s2 + $heights[$j] - 1",
    "    $ws.Range(($s2.ToString() + ':' + $e2.ToString())).EntireRow.Hidden = $false",
    "  }",
    "  $wb.Close($false)",
    "  $wb = $null",
    "  Write-Output 'OK'",
    "} finally {",
    "  if ($wb) { try { $wb.Close($false) } catch {} }",
    "  try { $xl.Quit() } catch {}",
    "  try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl) } catch {}",
    "  try { Remove-Item -LiteralPath " + q(tmp) + " -Recurse -Force -ErrorAction SilentlyContinue } catch {}",
    "}"
  ].join('\n');

  await jalankanPs(ps, 240000);

  if (!fs.existsSync(outXlsx)) throw new Error('Excel tidak menghasilkan file .xlsx.');
  const pdfs = daftarPdf.filter(p => fs.existsSync(p.path) && fs.statSync(p.path).size > 0);
  if (!pdfs.length) throw new Error('Excel tidak menghasilkan PDF rekap.');
  return { xlsx: outXlsx, pdfs };
}

module.exports = { bacaAngkaOpm, isiRekapExcel, MASTER_XLS, KATEGORI_REKAP: 'INPUT REKAP OPM' };
