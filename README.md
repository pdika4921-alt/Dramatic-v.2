# Sistem Rekonsiliasi Konstruksi — PT Telkom Akses (Bogor)

Sistem web untuk manajemen **master lokasi**, **upload eviden lapangan**,
**audit kelengkapan berkas**, perhitungan **RAB** dan **rekonsiliasi DRM vs REKON**
serta generate dokumen (LACT, BA Prelim, BAUT, BA Tambah Kurang, dll).

## Struktur Folder

```
web-eviden/
├── public/                 # Static frontend (di-serve langsung)
│   ├── login.html          # Halaman login (role-based redirect)
│   ├── admin.html          # Dashboard admin: master lokasi + audit checker
│   ├── user.html           # Upload eviden lapangan (ODP, OPM, Cut Over, dst)
│   ├── rab.html            # Input RAB & rekonsiliasi DRM vs REKON
│   ├── 404.html            # Halaman error 404 ter-branding
│   ├── bctelkomakses.jpg   # Background foto
│   ├── images/             # Aset gambar statis
│   ├── assets/
│   │   ├── css/theme.css    # Design system bersama (token + komponen ds-)
│   │   ├── js/util.js       # Helper: fetch+CSRF wrapper, escapeHtml
│   │   ├── img/logo-ta.png      # Logo Telkom Akses (wordmark)
│   │   └── img/favicon-ta.png   # Favicon brand (simbol Telkom Akses)
│   └── uploads/            # Berkas unggahan (ter-generate per lokasi)
├── templates/             # Template Word (.docx) untuk generate dokumen
└── node_modules/          # Dependency server (Express, dll)
```

## Halaman & Endpoint API

| Halaman     | Fungsi utama                                                         |
|-------------|---------------------------------------------------------------------|
| login.html  | Autentikasi via `POST /api/login` (role: admin / user)              |
| admin.html  | CRUD master lokasi, audit checker, generate dokumen                 |
| user.html   | Upload eviden + input BA Cut Over (`/api/lokasi`, upload)           |
| rab.html    | Input RAB, rekonsiliasi, generate BA Tambah Kurang                  |

## Design System

Semua halaman menggunakan token yang didefinisikan di `public/assets/css/theme.css`
(variabel `--brand-*`, `--slate-*`, radius, shadow, font). Komponen bersama
menggunakan prefix `ds-` (mis. `.ds-btn`, `.ds-card`, `.ds-input`, `.ds-nav`)
agar tidak bentrok dengan utility Tailwind (admin.html).

## Catatan Keamanan (penting)

- **CSRF**: `assets/js/util.js` secara otomatis menyisipkan header
  `X-CSRF-Token` pada request mutating (`POST/PUT/PATCH/DELETE`) **jika**
  `<meta name="csrf-token">` diisi oleh server. Pastikan backend:
  1. Menerbitkan token CSRF per-session dan menyuntiknya ke `<meta>`.
  2. Memvalidasi header tersebut pada semua endpoint mutating.
- **Upload file**: validasi **tipe + ukuran** di sisi server untuk
  `image/*`, `application/pdf`, `.kml`, `.kmz` (jangan hanya mengandalkan
  atribut `accept` di client).
- **Auth**: pastikan verifikasi kredensial & otorisasi role dilakukan di server;
  jangan ada credential hardcode di JavaScript.
- **XSS**: untuk teks user yang disuntikkan via `innerHTML`, gunakan
  `window.escapeHtml()` (tersedia di `util.js`).

## Performa

- Font di-load via `<link rel="preconnect">` + stylesheet (bukan `@import`).
- `admin.html` menggunakan Tailwind CDN browser — untuk produksi, ganti dengan
  Tailwind hasil build agar lebih ringan & stabil.

## Cara Menjalankan

Lihat dokumentasi server (Node/Express) pada konfigurasi proyek ini.
Pastikan `public/` di-serve sebagai static root agar path `assets/...` dan
`/login.html`, `/admin.html`, `/user.html`, `/rab.html` ter-resolusi dengan benar.
