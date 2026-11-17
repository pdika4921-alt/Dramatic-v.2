// kml-renderer.js
// Modul untuk memproses file KML/KMZ menjadi gambar peta (PNG).
// - KML : XML geospasial (Google Earth)
// - KMZ : arsip ZIP berisi doc.kml
// Proses: parse geometri -> unduh tile citra satelit (Esri World Imagery,
//         gaya Google Earth) -> komposit tile via sharp -> overlay SVG geometri
//         (Point/Line/Polygon) -> PNG akhir. Cadangan bila satelit gagal: OSM.
const fs = require('fs');
const { DOMParser } = require('@xmldom/xmldom');
const PizZip = require('pizzip');
const sharp = require('sharp');

const TILE_SATELIT = 'https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}';
const TILE_URL = 'https://tile.openstreetmap.org/{z}/{x}/{y}.png';
const TILE_SIZE = 256;
const USER_AGENT = 'web-eviden/1.0 (aplikasi eviden konstruksi)';

function lonToX(lon, z) {
  return ((lon + 180) / 360) * Math.pow(2, z);
}

function latToY(lat, z) {
  const rad = (lat * Math.PI) / 180;
  return ((1 - Math.log(Math.tan(rad) + 1 / Math.cos(rad)) / Math.PI) / 2) * Math.pow(2, z);
}

function sleep(ms) {
  return new Promise((r) => setTimeout(r, ms));
}

// Cari elemen keturunan berdasarkan nama lokal (abaikan prefix namespace).
function cariElemen(el, localName) {
  const hasil = [];
  (function walk(node) {
    const cs = node.childNodes || [];
    for (let i = 0; i < cs.length; i++) {
      const c = cs[i];
      if (c.nodeType !== 1) continue;
      const ln = c.localName || (c.tagName || '').replace(/^.*:/, '');
      if (ln === localName) hasil.push(c);
      walk(c);
    }
  })(el);
  return hasil;
}

// Ambil text content elemen pertama dengan nama lokal tertentu.
function teksElemen(el, localName) {
  const found = cariElemen(el, localName);
  return found.length ? (found[0].textContent || '').trim() : '';
}

// Ubah string koordinat KML "lon,lat,alt lon,lat,alt" menjadi array [lon, lat].
function ambilKoordinat(text) {
  const pts = [];
  (text || '').split(/\s+/).forEach((tok) => {
    const m = /^(-?[\d.]+),(-?[\d.]+)/.exec(tok.trim());
    if (m) pts.push([parseFloat(m[1]), parseFloat(m[2])]);
  });
  return pts;
}

// Traverse geometri (rekursif utk MultiGeometry).
function walkGeometri(el, cb) {
  const cs = el.childNodes || [];
  for (let i = 0; i < cs.length; i++) {
    const c = cs[i];
    if (c.nodeType !== 1) continue;
    const ln = c.localName || (c.tagName || '').replace(/^.*:/, '');
    if (ln === 'MultiGeometry') walkGeometri(c, cb);
    else if (ln === 'Point' || ln === 'LineString' || ln === 'Polygon') cb(c, ln);
  }
}

// Ekstrak semua geometri dari dokumen KML.
// Mengembalikan array feature: { name, type: 'point'|'line'|'polygon', points:[[lon,lat],...] }
function parseKml(xml) {
  // Bersihkan BOM / spasi di awal supaya parser XML tidak error
  xml = String(xml || '').replace(/^\uFEFF/, '').replace(/^\s+/, '');
  const doc = new DOMParser().parseFromString(xml, 'text/xml');
  const features = [];
  if (!doc) return features;

  const placemarks = cariElemen(doc, 'Placemark');
  for (let i = 0; i < placemarks.length; i++) {
    const pm = placemarks[i];
    const nama = teksElemen(pm, 'name') || ('Lokasi ' + (i + 1));

    // gx:Track (rekaman GPS Google Earth)
    const coordsTrack = cariElemen(pm, 'coord');
    if (coordsTrack.length) {
      const pts = [];
      coordsTrack.forEach((c) => {
        const t = (c.textContent || '').trim();
        const m = /^(-?[\d.]+)[\s,]+(-?[\d.]+)/.exec(t);
        if (m) pts.push([parseFloat(m[1]), parseFloat(m[2])]);
      });
      if (pts.length >= 1) features.push({ name: nama, type: 'line', points: pts });
      continue;
    }

    walkGeometri(pm, (geom, tipe) => {
      const coordsEls = cariElemen(geom, 'coordinates');
      if (!coordsEls.length) return;
      if (tipe === 'Point') {
        const pts = ambilKoordinat(coordsEls[0].textContent);
        if (pts.length) features.push({ name: nama, type: 'point', points: [pts[0]] });
      } else {
        coordsEls.forEach((ce) => {
          const pts = ambilKoordinat(ce.textContent);
          if (pts.length >= 1) features.push({ name: nama, type: tipe === 'Polygon' ? 'polygon' : 'line', points: pts });
        });
      }
    });
  }
  return features;
}

// Baca buffer KML/KMZ menjadi teks XML.
function bacaKml(xml) {
  if (typeof xml === 'string' || Buffer.isBuffer(xml)) {
    const buf = Buffer.isBuffer(xml) ? xml : Buffer.from(xml, 'utf8');
    // Deteksi KMZ: mulai dengan magic bytes ZIP (PK)
    if (buf.length > 2 && buf[0] === 0x50 && buf[1] === 0x4b) {
      const zip = new PizZip(buf);
      const namaKml = cariDocKml(zip);
      if (!namaKml) throw new Error('Arsip KMZ tidak mengandung file .kml.');
      return zip.file(namaKml).asText();
    }
    return buf.toString('utf8');
  }
  throw new Error('Data KML tidak valid.');
}

function cariDocKml(zip) {
  const files = zip.files || {};
  const nama = Object.keys(files).filter((n) => /\.kml$/i.test(n));
  if (!nama.length) return null;
  return nama.sort((a, b) => {
    const pa = a.toLowerCase().split('/').length;
    const pb = b.toLowerCase().split('/').length;
    return pa - pb;
  })[0];
}

// ── RENDER PETA ──────────────────────────────────────────────────────────────

function tileBounds(bbox, z) {
  let [w, s, e, n] = bbox;
  if (e < w) e += 360;
  let xMin = Math.floor(lonToX(w, z));
  let xMax = Math.floor(lonToX(e, z));
  let yMin = Math.floor(latToY(n, z));
  let yMax = Math.floor(latToY(s, z));
  const maxXY = Math.pow(2, z) - 1;
  xMin = Math.max(0, Math.min(maxXY, xMin));
  xMax = Math.max(0, Math.min(maxXY, xMax));
  yMin = Math.max(0, Math.min(maxXY, yMin));
  yMax = Math.max(0, Math.min(maxXY, yMax));
  if (xMax < xMin) xMax = xMin;
  if (yMax < yMin) yMax = yMin;
  return { xMin, xMax, yMin, yMax };
}

function pilihZoom(bbox, maxPx, maxTiles) {
  for (let z = 1; z <= 19; z++) {
    const r = tileBounds(bbox, z);
    const wpx = (r.xMax - r.xMin + 1) * TILE_SIZE;
    const hpx = (r.yMax - r.yMin + 1) * TILE_SIZE;
    const jumlah = (r.xMax - r.xMin + 1) * (r.yMax - r.yMin + 1);
    if (wpx > maxPx || hpx > maxPx || jumlah > maxTiles) return Math.max(1, z - 1);
  }
  return 19;
}

async function unduhTileSatu(url, accept) {
  for (let percobaan = 0; percobaan < 3; percobaan++) {
    try {
      const ctl = new AbortController();
      const timer = setTimeout(() => ctl.abort(), 15000);
      const res = await fetch(url, {
        headers: { 'User-Agent': USER_AGENT, 'Accept': accept },
        signal: ctl.signal,
      });
      clearTimeout(timer);
      if (!res.ok) throw new Error('HTTP ' + res.status);
      return Buffer.from(await res.arrayBuffer());
    } catch (e) {
      await sleep(400);
    }
  }
  return null;
}

async function unduhTile(r, z) {
  const hasil = [];
  for (let x = r.xMin; x <= r.xMax; x++) {
    for (let y = r.yMin; y <= r.yMax; y++) {
      const urlSat = TILE_SATELIT.replace('{z}', z).replace('{x}', x).replace('{y}', y);
      const urlOsm = TILE_URL.replace('{z}', z).replace('{x}', x).replace('{y}', y);
      let buf = await unduhTileSatu(urlSat, 'image/jpeg,image/*;q=0.8,*/*;q=0.5');
      if (!buf) {
        // Cadangan: tile jalan (OSM) bila citra satelit gagal diunduh.
        buf = await unduhTileSatu(urlOsm, 'image/png,*/*');
      }
      if (!buf) throw new Error('Gagal mengunduh tile peta: ' + urlSat);
      hasil.push({ x, y, buf });
      await sleep(160);
    }
  }
  return hasil;
}

function escXml(s) {
  return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;')
    .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function buatSvgOverlay(features, ctx) {
  const { z, xMin, yMin, W, H } = ctx;
  const px = (lon) => (lonToX(lon, z) - xMin) * TILE_SIZE;
  const py = (lat) => (latToY(lat, z) - yMin) * TILE_SIZE;

  let svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}" viewBox="0 0 ${W} ${H}">`;
  svg += '<rect x="0" y="0" width="' + W + '" height="' + H + '" fill="none" stroke="#e5e7eb" stroke-width="2"/>';

  features.forEach((f) => {
    if (f.type === 'line' || f.type === 'polygon') {
      const closed = f.type === 'polygon';
      let d = '';
      f.points.forEach((p, i) => {
        d += (i === 0 ? 'M' : 'L') + px(p[0]).toFixed(1) + ' ' + py(p[1]).toFixed(1);
      });
      if (closed) d += 'Z';
      const fill = closed ? 'rgba(59,130,246,0.25)' : 'none';
      // Alur bergaya Google Earth: garis terang dengan bayangan gelap agar menonjol di atas citra satelit.
      svg += `<path d="${d}" fill="none" stroke="#0b1f4b" stroke-opacity="0.65" stroke-width="9" stroke-linejoin="round" stroke-linecap="round"/>`;
      svg += `<path d="${d}" fill="${fill}" stroke="#3b82f6" stroke-width="4.5" stroke-linejoin="round" stroke-linecap="round"/>`;
    }
  });

  features.forEach((f) => {
    if (f.type !== 'point' || !f.points.length) return;
    const cx = px(f.points[0][0]);
    const cy = py(f.points[0][1]);
    svg += `<circle cx="${cx.toFixed(1)}" cy="${cy.toFixed(1)}" r="7" fill="#e11d48" stroke="#ffffff" stroke-width="2.5"/>`;
    if (f.name) {
      svg += `<text x="${(cx + 10).toFixed(1)}" y="${(cy - 8).toFixed(1)}" font-size="14" font-family="Arial, Helvetica, sans-serif" font-weight="bold" fill="#ffffff" stroke="#0f172a" stroke-width="3.5" paint-order="stroke">${escXml(f.name)}</text>`;
    }
  });

  svg += `<text x="${W - 8}" y="${H - 8}" text-anchor="end" font-size="11" font-family="Arial, Helvetica, sans-serif" fill="#ffffff" fill-opacity="0.85" stroke="#000000" stroke-width="3" paint-order="stroke">Citra © Esri World Imagery</text>`;

  svg += '</svg>';
  return svg;
}

// Render semua feature ke file PNG.
async function renderPeta(features, outPath) {
  const semua = [];
  features.forEach((f) => f.points.forEach((p) => semua.push(p)));
  if (!semua.length) throw new Error('Tidak ada koordinat yang bisa dipetakan.');

  let minLon = Infinity, maxLon = -Infinity, minLat = Infinity, maxLat = -Infinity;
  semua.forEach((p) => {
    if (p[0] < minLon) minLon = p[0];
    if (p[0] > maxLon) maxLon = p[0];
    if (p[1] < minLat) minLat = p[1];
    if (p[1] > maxLat) maxLat = p[1];
  });

  const pad = Math.max(maxLon - minLon, maxLat - minLat) * 0.15 + 1e-6;
  const w = minLon - pad, e = maxLon + pad;
  const s = minLat - pad, n = maxLat + pad;

  const z = pilihZoom([w, s, e, n], 1200, 64);
  const r = tileBounds([w, s, e, n], z);
  const W = (r.xMax - r.xMin + 1) * TILE_SIZE;
  const H = (r.yMax - r.yMin + 1) * TILE_SIZE;

  const tiles = await unduhTile(r, z);

  const comps = tiles.map((t) => ({
    input: t.buf,
    left: (t.x - r.xMin) * TILE_SIZE,
    top: (t.y - r.yMin) * TILE_SIZE,
  }));

  const svg = buatSvgOverlay(features, { z, xMin: r.xMin, yMin: r.yMin, W, H });
  comps.push({ input: Buffer.from(svg), left: 0, top: 0 });

  let img = sharp({ create: { width: W, height: H, channels: 3, background: '#f2efe9' } });
  img = img.composite(comps);

  await img.png({ compressionLevel: 6 }).toFile(outPath);
}

module.exports = {
  parseKml,
  bacaKml,
  renderPeta,
  ambilKoordinat,
};
