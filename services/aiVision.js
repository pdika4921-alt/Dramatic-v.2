// Lapisan 3 — verifikasi AI Vision (Claude / Gemini / mock) dengan antrean,
// batas waktu, dan percobaan ulang. Respons dipaksa berupa JSON murni.
const cfg = require('./verifyConfig');
const sharp = require('sharp');

/* ── Antrean konkurensi (maksimal n foto dianalisis bersamaan) ───────────── */
let berjalan = 0;
const antre = [];
function tiket() {
  return new Promise(resolve => {
    antre.push(resolve);
    lanjut();
  });
}
function lanjut() {
  while (berjalan < cfg.ai.maxConcurrent && antre.length) {
    berjalan++;
    antre.shift()();
  }
}
function lepas() {
  berjalan--;
  lanjut();
}

async function gambarBase64(filePath) {
  const sumber = Buffer.isBuffer(filePath) ? filePath : await require('fs').promises.readFile(filePath);
  const buf = await sharp(sumber)
    .resize(cfg.ai.maxSisi, cfg.ai.maxSisi, { fit: 'inside', withoutEnlargement: true })
    .jpeg({ quality: 85 })
    .toBuffer();
  return buf.toString('base64');
}

function promptSistem() {
  return [
    'Anda adalah auditor eviden lapangan telekomunikasi Indonesia (Telkom Akses).',
    'Tugas: menilai KEASLIHAN sebuah foto bukti pekerjaan hanya dari tampilan gambar.',
    'Perhatikan: apakah foto nyata di lokasi kerja (alat, tiang, kabel, galian, orang berseragam),',
    'apakah ada tanda stempel/tanggal terbaca, apakah ada tanda editan/layar/screenshot/potongan ganjil,',
    'dan apakah isi foto masuk akal untuk pekerjaan yang diklaim.',
    'ANALISA LATAR BELAKANG secara khusus: konsistensi arah cahaya dan bayangan antar objek,',
    'perspektif yang wajar, ketajaman/noise yang berbeda antara objek utama dan latar,',
    'tepi tempelan yang janggal, pola clone-stamp/repetisi identik, tanda region berbeda resolusi,',
    'batas layar/notifikasi yang menandakan screenshot, dan objek latar yang tidak masuk akal.',
    'Nilai juga apakah ISI foto SESUAI dengan kategori eviden yang disebut pada konteks.',
    'JANGAN menebak dari nama berkas.',
    'Balas HANYA dengan JSON valid tanpa teks lain, persis format:',
    '{"skor_keaslian":0-100,"jenis_foto":"","stempel_terbaca":true|false|null,',
    '"indikasi_editan":"tidak"|"kemungkinan"|"ya",',
    '"latar_belakang":"wajar"|"mencurigakan","isi_sesuai":true|false,',
    '"alasan_singkat":"≤20 kata","detail":"≤40 kata"}'
  ].join(' ');
}

function promptUser(konteks) {
  const k = konteks || {};
  const bagian = [];
  if (k.kategori) bagian.push(`Kategori eviden: ${k.kategori}`);
  if (k.subKategori) bagian.push(`Titik/ODP: ${k.subKategori}`);
  if (k.namaLokasi) bagian.push(`Lokasi pekerjaan: ${k.namaLokasi}`);
  if (k.tematik) bagian.push(`Tematik: ${k.tematik}`);
  bagian.push('Nilai foto ini: keaslian, latar belakang, serta kesesuaian isi dengan kategori eviden di atas, lalu balas JSON sesuai format yang diminta.');
  return bagian.join('\n');
}

/* ── Provider ────────────────────────────────────────────────────────────── */
async function denganBatasWaktu(promise, ms) {
  const ac = new AbortController();
  const timer = setTimeout(() => ac.abort(), ms);
  try {
    const res = await Promise.race([
      promise(ac.signal),
      new Promise((_, tolak) => setTimeout(() => tolak(new Error('Timeout AI setelah ' + ms + ' ms')), ms))
    ]);
    return res;
  } finally {
    clearTimeout(timer);
  }
}

async function panggilClaude(base64, teks, apiKey, model) {
  const r = await fetch('https://api.anthropic.com/v1/messages', {
    method: 'POST',
    headers: {
      'content-type': 'application/json',
      'x-api-key': apiKey,
      'anthropic-version': '2023-06-01'
    },
    body: JSON.stringify({
      model: model || 'claude-3-5-sonnet-latest',
      max_tokens: 700,
      system: promptSistem(),
      messages: [{
        role: 'user',
        content: [
          { type: 'image', source: { type: 'base64', media_type: 'image/jpeg', data: base64 } },
          { type: 'text', text: teks }
        ]
      }]
    }),
    signal: undefined
  });
  if (!r.ok) {
    const badan = await r.text().catch(() => '');
    const err = new Error('Claude HTTP ' + r.status + (/:/.test(badan) ? '' : ''));
    // 429/529/503 = batas laju / sibuk — tunggu sebentar sebelum percobaan ulang
    if (r.status === 429 || r.status === 529 || r.status >= 500) err.retryAfterMs = 8000;
    throw err;
  }
  const d = await r.json();
  const teksHasil = (d.content || []).map(c => c.text || '').join('');
  if (!teksHasil) throw new Error('Claude mengembalikan respons kosong');
  return teksHasil;
}

async function panggilGemini(base64, teks, apiKey, model) {
  const namaModel = model || 'gemini-3.8-flash';
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${encodeURIComponent(namaModel)}:generateContent`;
  const r = await fetch(url, {
    method: 'POST',
    headers: { 'content-type': 'application/json', 'x-goog-api-key': apiKey },
    body: JSON.stringify({
      system_instruction: { parts: [{ text: promptSistem() }] },
      contents: [{
        role: 'user',
        parts: [
          { inline_data: { mime_type: 'image/jpeg', data: base64 } },
          { text: teks }
        ]
      }],
      generationConfig: { maxOutputTokens: 700, temperature: 0.2 }
    })
  });
  if (!r.ok) {
    const badan = await r.text().catch(() => '');
    let pesan = '';
    try { pesan = (JSON.parse(badan).error || {}).message || ''; } catch (e) { /* bukan JSON */ }
    const err = new Error('Gemini HTTP ' + r.status + (pesan ? ': ' + pesan.slice(0, 140) : ''));
    // Google menyebut batas laju sebagai 429, sering juga 503; hormati retryDelay bila ada
    const m = /"retryDelay":\s*"(\d+(?:\.\d+)?)s"/.exec(badan);
    if (m) err.retryAfterMs = Math.round(parseFloat(m[1]) * 1000);
    else if (r.status === 429 || r.status >= 500) err.retryAfterMs = 8000;
    throw err;
  }
  const d = await r.json();
  const bagian = (d.candidates && d.candidates[0] && d.candidates[0].content && d.candidates[0].content.parts) || [];
  const teksHasil = bagian.map(p => p.text || '').join('');
  if (!teksHasil) throw new Error('Gemini mengembalikan respons kosong');
  return teksHasil;
}

let hitungMock = 0;
function panggilMock(namaFile, kategori) {
  const pemicu = (String(namaFile || '') + ' ' + String(kategori || '')).toLowerCase();
  if (pemicu.includes('gagal')) throw new Error('Simulasi kegagalan API AI (mock)');
  hitungMock++;
  const curiga = pemicu.includes('aisuspect');
  return JSON.stringify({
    skor_keaslian: curiga ? 35 : 88,
    jenis_foto: 'foto lapangan',
    stempel_terbaca: curiga ? false : true,
    indikasi_editan: curiga ? 'ya' : 'tidak',
    latar_belakang: curiga ? 'mencurigakan' : 'wajar',
    isi_sesuai: !curiga,
    alasan_singkat: curiga ? '(mock) terlihat hasil suntingan kuat' : '(mock) foto lapangan wajar',
    detail: 'Analisis otomatis mode mock — ganti AI_PROVIDER untuk hasil sesungguhnya.',
    n: hitungMock
  });
}

/* ── Parsing JSON longgar ────────────────────────────────────────────────── */
function parseJsonLonggar(teks) {
  const potong = String(teks || '').replace(/```(?:json)?/gi, '').trim();
  try { return JSON.parse(potong); } catch (e) { /* lanjut */ }
  const buka = potong.indexOf('{');
  const tutup = potong.lastIndexOf('}');
  if (buka >= 0 && tutup > buka) {
    try { return JSON.parse(potong.slice(buka, tutup + 1)); } catch (e) { /* lanjut */ }
  }
  throw new Error('Respons AI bukan JSON valid');
}

function validasi(hasil) {
  const skor = Number(hasil.skor_keaslian);
  if (!Number.isFinite(skor)) throw new Error('Respons AI tanpa skor_keaslian');
  hasil.skor_keaslian = Math.max(0, Math.min(100, Math.round(skor)));
  hasil.jenis_foto = hasil.jenis_foto ? String(hasil.jenis_foto) : '';
  hasil.alasan_singkat = hasil.alasan_singkat ? String(hasil.alasan_singkat) : '';
  hasil.detail = hasil.detail ? String(hasil.detail) : '';
  if (!['tidak', 'kemungkinan', 'ya'].includes(hasil.indikasi_editan)) hasil.indikasi_editan = null;
  if (!['wajar', 'mencurigakan'].includes(hasil.latar_belakang)) hasil.latar_belakang = null;
  if (typeof hasil.isi_sesuai !== 'boolean') hasil.isi_sesuai = null;
  return hasil;
}

/* ── API utama ───────────────────────────────────────────────────────────── */
async function analisisFoto({ filePath, konteks, namaFile }) {
  await tiket();
  try {
    const provider = cfg.ai.provider === 'anthropic' ? 'claude' : cfg.ai.provider;
    if (provider !== 'mock' && !cfg.ai.apiKey) {
      throw new Error(`Provider "${provider}" memerlukan AI_API_KEY di .env`);
    }

    let base64 = null;
    if (provider !== 'mock') base64 = await gambarBase64(filePath);
    const teks = promptUser(konteks);

    // AI_MODEL boleh berupa daftar model dipisah koma — percobaan berganti model
    // tiap kali, sehingga kuota harian tiap model (free tier: 20/hari/model)
    // terpakai bergiliran, bukan habis di satu model.
    const daftarModel = String(cfg.ai.model || '').split(',').map(s => s.trim()).filter(Boolean);

    let galat = null;
    for (let percobaan = 0; percobaan <= cfg.ai.retry; percobaan++) {
      const m = daftarModel.length ? daftarModel[percobaan % daftarModel.length] : '';
      let modelDipakai = '';
      try {
        let teksHasil;
        if (provider === 'mock') {
          teksHasil = panggilMock(namaFile, konteks && konteks.kategori);
        } else if (provider === 'claude') {
          teksHasil = await denganBatasWaktu(s => panggilClaude(base64, teks, cfg.ai.apiKey, m), cfg.ai.timeoutMs);
          modelDipakai = m || 'claude-3-5-sonnet-latest';
        } else if (provider === 'gemini') {
          teksHasil = await denganBatasWaktu(s => panggilGemini(base64, teks, cfg.ai.apiKey, m), cfg.ai.timeoutMs);
          modelDipakai = m || 'gemini-3.8-flash';
        } else {
          throw new Error('AI_PROVIDER tidak dikenal: ' + provider);
        }
        const hasil = validasi(parseJsonLonggar(teksHasil));
        if (modelDipakai) hasil.model = modelDipakai;
        return hasil;
      } catch (e) {
        galat = e;
        if (percobaan < cfg.ai.retry) {
          // Backoff eksponensial, minimal 500 ms, maksimal 30 detik;
          // permintaan lanjut dari API (429/503) diprioritaskan.
          const jeda = Math.min(30000, Math.max(500 * (percobaan + 1), (e && e.retryAfterMs) || 0));
          await new Promise(r => setTimeout(r, jeda));
        }
      }
    }
    throw galat || new Error('Analisis AI gagal');
  } finally {
    lepas();
  }
}

module.exports = { analisisFoto };
