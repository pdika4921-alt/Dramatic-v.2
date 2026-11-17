const PizZip = require('pizzip');
const fs = require('fs');

const SRC = 'templates/templateRelokasi.docx';
const OUT = ['templates/template_HEM.docx', 'templates/template_MICRODEMAND.docx'];

const zip = new PizZip(fs.readFileSync(SRC, 'binary'));
const docXml = zip.files['word/document.xml'].asText();

const bodyStart = docXml.indexOf('<w:body>');
const bodyEnd = docXml.indexOf('</w:body>');
if (bodyStart < 0 || bodyEnd < 0) throw new Error('Struktur body tidak ditemukan');
const sectPr = docXml.slice(bodyEnd - 500, bodyEnd).match(/<w:sectPr[\s\S]*?<\/w:sectPr>/);
if (!sectPr) throw new Error('sectPr tidak ditemukan');

// ── Helper XML ────────────────────────────────────────────────────────────────
const p = (content, opts = {}) =>
  `<w:p>${opts.pPr || ''}${content}</w:p>`;

const pStyle = (style, extra = '') =>
  `<w:pPr><w:pStyle w:val="${style}"/>${extra}</w:pPr>`;

const r = (text, rPr = '') =>
  `<w:r><w:rPr>${rPr}</w:rPr><w:t xml:space="preserve">${text}</w:t></w:r>`;

const cell = (inner, w = '2440') =>
  `<w:tc><w:tcPr><w:tcW w:w="${w}" w:type="dxa"/></w:tcPr>${inner}</w:tc>`;

const B = '<w:b/><w:bCs/>';
const I = '<w:i/>';
const sz16 = '<w:sz w:val="16"/><w:szCs w:val="16"/>';
const sz18 = '<w:sz w:val="18"/><w:szCs w:val="18"/>';
const sz20 = '<w:sz w:val="20"/><w:szCs w:val="20"/>';
const sz32 = '<w:sz w:val="32"/><w:szCs w:val="32"/>';

const tblPr = `<w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/><w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/></w:tblPr>`;
const tblGrid4 = `<w:tblGrid><w:gridCol w:w="2440"/><w:gridCol w:w="2440"/><w:gridCol w:w="2440"/><w:gridCol w:w="2440"/></w:tblGrid>`;
const tblGrid5 = `<w:tblGrid><w:gridCol w:w="1400"/><w:gridCol w:w="1800"/><w:gridCol w:w="1200"/><w:gridCol w:w="4000"/><w:gridCol w:w="1390"/></w:tblGrid>`;
const tblGrid7 = `<w:tblGrid><w:gridCol w:w="700"/><w:gridCol w:w="1700"/><w:gridCol w:w="2800"/><w:gridCol w:w="900"/><w:gridCol w:w="1200"/><w:gridCol w:w="1200"/><w:gridCol w:w="1200"/></w:tblGrid>`;
const tblGridPort = `<w:tblGrid><w:gridCol w:w="1300"/><w:gridCol w:w="1500"/><w:gridCol w:w="7000"/></w:tblGrid>`;

const spacePara = `<w:p><w:pPr><w:spacing w:before="120" w:after="60"/></w:pPr></w:p>`;

// Paragraf berisi SATU tag docxtemplater (section loop harus di dalam <w:t>)
const tagPara = (tag) =>
  `<w:p><w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr><w:r><w:t xml:space="preserve">${tag}</w:t></w:r></w:p>`;

// ── BODY BARU ────────────────────────────────────────────────────────────────
const body = `<w:body>

${p(r('LAPORAN EVIDEN BERKAS {tematik}', sz32), pStyle('Title', '<w:jc w:val="center"/>'))}
${p(r('Laporan Rekonsiliasi Konstruksi — PT Telkom Akses', sz18), pStyle('Subtitle', '<w:jc w:val="center"/>'))}

${p(r('INFORMASI PROYEK', sz20 + B), pStyle('Heading1'))}

<w:tbl>${tblPr}${tblGrid4}
  <w:tr>
    ${cell(`<w:p><w:pPr><w:spacing w:before="40" w:after="40"/></w:pPr>${r('PROYEK', B + sz16)}${r(': {proyek}', sz16)}</w:p>`, '2440')}
    ${cell(`<w:p><w:pPr><w:spacing w:before="40" w:after="40"/></w:pPr>${r('KONTRAK', B + sz16)}${r(': {kontrak}', sz16)}</w:p>`, '2440')}
    ${cell(`<w:p><w:pPr><w:spacing w:before="40" w:after="40"/></w:pPr>${r('SURAT PESANAN', B + sz16)}${r(': {surat_pesan}', sz16)}</w:p>`, '2440')}
    ${cell(`<w:p><w:pPr><w:spacing w:before="40" w:after="40"/></w:pPr>${r('DISTRICT', B + sz16)}${r(': {district}', sz16)}</w:p>`, '2440')}
  </w:tr>
  <w:tr>
    ${cell(`<w:p><w:pPr><w:spacing w:before="40" w:after="40"/></w:pPr>${r('NAMA LOKASI', B + sz16)}${r(': {lokasi_smile}', sz16)}</w:p>`, '2440')}
    ${cell(`<w:p><w:pPr><w:spacing w:before="40" w:after="40"/></w:pPr>${r('PELAKSANA', B + sz16)}${r(': {pelaksana}', sz16)}</w:p>`, '2440')}
    ${cell(`<w:p><w:pPr><w:spacing w:before="40" w:after="40"/></w:pPr>${r('PENGAWAS LAPANGAN', B + sz16)}${r(': {nama_waspang}', sz16)}</w:p>`, '2440')}
    ${cell(`<w:p><w:pPr><w:spacing w:before="40" w:after="40"/></w:pPr>${r('NIK', B + sz16)}${r(': {nik_waspang}', sz16)}</w:p>`, '2440')}
  </w:tr>
</w:tbl>

${spacePara}
${p(r('DAFTAR EVIDEN BERKAS', sz20 + B), pStyle('Heading1'))}
${spacePara}

${tagPara('{#kategori_list}')}
${p(r('KATEGORI {nomor}. {nama}', sz18 + B) + r(`  — {jumlah} foto`, sz16), pStyle('Heading2'))}
${tagPara('{#boq}')}
<w:tbl>${tblPr}${tblGrid7}
  <w:tr>
    ${cell(`<w:p><w:pPr><w:spacing w:before="40" w:after="40"/><w:jc w:val="center"/></w:pPr>${r('NO', B + sz16)}</w:p>`, '700')}
    ${cell(`<w:p><w:pPr><w:spacing w:before="40" w:after="40"/></w:pPr>${r('DESIGNATOR', B + sz16)}</w:p>`, '1700')}
    ${cell(`<w:p><w:pPr><w:spacing w:before="40" w:after="40"/></w:pPr>${r('URAIAN PEKERJAAN', B + sz16)}</w:p>`, '2800')}
    ${cell(`<w:p><w:pPr><w:spacing w:before="40" w:after="40"/></w:pPr>${r('SATUAN', B + sz16)}</w:p>`, '900')}
    ${cell(`<w:p><w:pPr><w:spacing w:before="40" w:after="40"/></w:pPr>${r('MATERIAL', B + sz16)}</w:p>`, '1200')}
    ${cell(`<w:p><w:pPr><w:spacing w:before="40" w:after="40"/></w:pPr>${r('JASA', B + sz16)}</w:p>`, '1200')}
    ${cell(`<w:p><w:pPr><w:spacing w:before="40" w:after="40"/></w:pPr>${r('VOLUME', B + sz16)}</w:p>`, '1200')}
  </w:tr>
  <w:tr>
    ${cell(`<w:p><w:pPr><w:spacing w:before="40" w:after="40"/><w:jc w:val="center"/></w:pPr>${r('{#boq_list}{nomor}', sz16)}</w:p>`, '700')}
    ${cell(`<w:p><w:pPr><w:spacing w:before="40" w:after="40"/></w:pPr>${r('{designator}', sz16)}</w:p>`, '1700')}
    ${cell(`<w:p><w:pPr><w:spacing w:before="40" w:after="40"/></w:pPr>${r('{uraian}', sz16)}</w:p>`, '2800')}
    ${cell(`<w:p><w:pPr><w:spacing w:before="40" w:after="40"/><w:jc w:val="center"/></w:pPr>${r('{satuan}', sz16)}</w:p>`, '900')}
    ${cell(`<w:p><w:pPr><w:spacing w:before="40" w:after="40"/><w:jc w:val="right"/></w:pPr>${r('{material}', sz16)}</w:p>`, '1200')}
    ${cell(`<w:p><w:pPr><w:spacing w:before="40" w:after="40"/><w:jc w:val="right"/></w:pPr>${r('{jasa}', sz16)}</w:p>`, '1200')}
    ${cell(`<w:p><w:pPr><w:spacing w:before="40" w:after="40"/><w:jc w:val="right"/></w:pPr>${r('{volume}{/boq_list}', sz16)}</w:p>`, '1200')}
  </w:tr>
  <w:tr>
    ${cell(`<w:p><w:pPr><w:spacing w:before="40" w:after="40"/></w:pPr>${r('', sz16)}</w:p>`, '700')}
    ${cell(`<w:p><w:pPr><w:spacing w:before="40" w:after="40"/></w:pPr>${r('', sz16)}</w:p>`, '1700')}
    ${cell(`<w:p><w:pPr><w:spacing w:before="40" w:after="40"/></w:pPr>${r('TOTAL', B + sz16)}</w:p>`, '2800')}
    ${cell(`<w:p><w:pPr><w:spacing w:before="40" w:after="40"/></w:pPr>${r('', sz16)}</w:p>`, '900')}
    ${cell(`<w:p><w:pPr><w:spacing w:before="40" w:after="40"/><w:jc w:val="right"/></w:pPr>${r('{boq_total_material}', B + sz16)}</w:p>`, '1200')}
    ${cell(`<w:p><w:pPr><w:spacing w:before="40" w:after="40"/><w:jc w:val="right"/></w:pPr>${r('{boq_total_jasa}', B + sz16)}</w:p>`, '1200')}
    ${cell(`<w:p><w:pPr><w:spacing w:before="40" w:after="40"/></w:pPr>${r('', sz16)}</w:p>`, '1200')}
  </w:tr>
</w:tbl>
${tagPara('{/boq}')}
${tagPara('{#boq_kosong}')}
${p(r('— Belum ada data BOQ Aktual —', I + sz16 + '<w:color w:val="808080"/>'), pStyle('BodyText'))}
${tagPara('{/boq_kosong}')}
${tagPara('{#bukan_boq}')}
${tagPara('{#kosong}')}
${p(r('— Belum ada foto untuk kategori ini —', I + sz16 + '<w:color w:val="808080"/>'), pStyle('BodyText'))}
${tagPara('{/kosong}')}
${tagPara('{#terlampir}')}
${p(r('Berkas terlampir: ', I + sz16 + '<w:color w:val="6b7280"/>') + r('{nama_file}', sz16 + I + '<w:color w:val="6b7280"/>'), pStyle('BodyText'))}
${tagPara('{/terlampir}')}
${tagPara('{#ada}')}
<w:tbl>${tblPr}${tblGrid4}
  <w:tr>
    ${cell(`<w:p><w:pPr><w:jc w:val="center"/></w:pPr>${r('{#foto_baris}', sz16)}</w:p><w:p><w:pPr><w:jc w:val="center"/><w:spacing w:before="60" w:after="60"/></w:pPr>${r('{%f1}', sz16)}</w:p>`, '2440')}
    ${cell(`<w:p><w:pPr><w:jc w:val="center"/><w:spacing w:before="60" w:after="60"/></w:pPr>${r('{%f2}', sz16)}</w:p>`, '2440')}
    ${cell(`<w:p><w:pPr><w:jc w:val="center"/><w:spacing w:before="60" w:after="60"/></w:pPr>${r('{%f3}', sz16)}</w:p>`, '2440')}
    ${cell(`<w:p><w:pPr><w:jc w:val="center"/><w:spacing w:before="60" w:after="60"/></w:pPr>${r('{%f4}', sz16)}</w:p><w:p><w:pPr><w:jc w:val="center"/></w:pPr>${r('{/foto_baris}', sz16)}</w:p>`, '2440')}
  </w:tr>
</w:tbl>
${tagPara('{/ada}')}
${tagPara('{/bukan_boq}')}
${tagPara('{/kategori_list}')}

${spacePara}
${p(r('EVIDEN REDAMAN ODP/OTB (PER PORT)', sz20 + B), pStyle('Heading1'))}
${spacePara}
${tagPara('{#redaman_ada}')}
${tagPara('{#redaman_list}')}
${p(r('EVIDEN REDAMAN — {nama} (Kapasitas {kapasitas} port)', sz18 + B), pStyle('Heading2'))}
<w:tbl>${tblPr}${tblGridPort}
  <w:tr>
    ${cell(`<w:p><w:pPr><w:spacing w:before="40" w:after="40"/></w:pPr>${r('PORT', B + sz16)}</w:p>`, '1300')}
    ${cell(`<w:p><w:pPr><w:spacing w:before="40" w:after="40"/></w:pPr>${r('STATUS', B + sz16)}</w:p>`, '1500')}
    ${cell(`<w:p><w:pPr><w:spacing w:before="40" w:after="40"/></w:pPr>${r('FOTO EVIDEN', B + sz16)}</w:p>`, '7000')}
  </w:tr>
  <w:tr>
    ${cell(`<w:p><w:pPr><w:spacing w:before="40" w:after="40"/><w:jc w:val="center"/></w:pPr>${r('{#port_list}PORT {nomor}', sz16)}</w:p>`, '1300')}
    ${cell(`<w:p><w:pPr><w:spacing w:before="40" w:after="40"/><w:jc w:val="center"/></w:pPr>${r('{#adaFoto}UKUR{/adaFoto}{#adaIsi}ADA USER (ISI){/adaIsi}{#adaKosong}KOSONG{/adaKosong}', sz16)}</w:p>`, '1500')}
    ${cell(`<w:p><w:pPr><w:spacing w:before="40" w:after="40"/><w:jc w:val="center"/></w:pPr>${r('{#adaFoto}', sz16)}</w:p>` +
           `<w:p><w:pPr><w:spacing w:before="40" w:after="40"/><w:jc w:val="center"/></w:pPr>${r('{%foto_path}', sz16)}</w:p>` +
           `<w:p><w:pPr><w:spacing w:before="40" w:after="40"/><w:jc w:val="center"/></w:pPr>${r('{/adaFoto}', sz16)}</w:p>` +
           `<w:p><w:pPr><w:spacing w:before="40" w:after="40"/><w:jc w:val="center"/></w:pPr>${r('{#bukan_foto}—{/bukan_foto}{/port_list}', sz16)}</w:p>`, '7000')}
  </w:tr>
</w:tbl>
${tagPara('{/redaman_list}')}
${tagPara('{/redaman_ada}')}
${tagPara('{#redaman_kosong}')}
${p(r('— Belum ada data eviden redaman —', I + sz16 + '<w:color w:val="808080"/>'), pStyle('BodyText'))}
${tagPara('{/redaman_kosong}')}

${spacePara}
${p(r('PETA / KML (HALAMAN PENUH)', sz20 + B), pStyle('Heading1'))}
${spacePara}
${tagPara('{#kml_ada}')}
${tagPara('{#kml_list}')}
${p(r('{nama}', sz18 + B), pStyle('Heading2'))}
${p(r('{%foto}', sz16), pStyle('BodyText', '<w:jc w:val="center"/>'))}
${tagPara('{/kml_list}')}
${tagPara('{/kml_ada}')}
${tagPara('{#kml_kosong}')}
${p(r('— Belum ada data peta / KML —', I + sz16 + '<w:color w:val="808080"/>'), pStyle('BodyText'))}
${tagPara('{/kml_kosong}')}

${spacePara}
${p(r('INPUT KHUSUS (BOQ AKTUAL / INPUT RFC / BON MATERIAL)', sz20 + B), pStyle('Heading1'))}
${spacePara}

${tagPara('{#ada_input}')}
<w:tbl>${tblPr}${tblGrid5}
  <w:tr>
    ${cell(`<w:p><w:pPr><w:spacing w:before="40" w:after="40"/></w:pPr>${r('KATEGORI', B + sz16)}</w:p>`, '1400')}
    ${cell(`<w:p><w:pPr><w:spacing w:before="40" w:after="40"/></w:pPr>${r('JUDUL', B + sz16)}</w:p>`, '1800')}
    ${cell(`<w:p><w:pPr><w:spacing w:before="40" w:after="40"/></w:pPr>${r('VOLUME', B + sz16)}</w:p>`, '1200')}
    ${cell(`<w:p><w:pPr><w:spacing w:before="40" w:after="40"/></w:pPr>${r('ISI', B + sz16)}</w:p>`, '4000')}
    ${cell(`<w:p><w:pPr><w:spacing w:before="40" w:after="40"/></w:pPr>${r('WAKTU', B + sz16)}</w:p>`, '1390')}
  </w:tr>
  <w:tr>
    ${cell(`<w:p><w:pPr><w:spacing w:before="40" w:after="40"/></w:pPr>${r('{#input_list}{kategori}', sz16)}</w:p>`, '1400')}
    ${cell(`<w:p><w:pPr><w:spacing w:before="40" w:after="40"/></w:pPr>${r('{judul}', sz16)}</w:p>`, '1800')}
    ${cell(`<w:p><w:pPr><w:spacing w:before="40" w:after="40"/></w:pPr>${r('{volume}', sz16)}</w:p>`, '1200')}
    ${cell(`<w:p><w:pPr><w:spacing w:before="40" w:after="40"/></w:pPr>${r('{isi}', sz16)}</w:p>`, '4000')}
    ${cell(`<w:p><w:pPr><w:spacing w:before="40" w:after="40"/></w:pPr>${r('{created_at}{/input_list}', sz16)}</w:p>`, '1390')}
  </w:tr>
</w:tbl>
${tagPara('{/ada_input}')}
${tagPara('{#kosong_input}')}
${p(r('— Belum ada input khusus (BOQ / RFC / Bon Material) —', I + sz16 + '<w:color w:val="808080"/>'), pStyle('BodyText'))}
${tagPara('{/kosong_input}')}

${sectPr[0]}
</w:body>`;

const newDocXml = docXml.slice(0, bodyStart) + body + docXml.slice(bodyEnd + '</w:body>'.length);

for (const out of OUT) {
  const z = new PizZip(fs.readFileSync(SRC, 'binary'));
  z.file('word/document.xml', newDocXml);
  fs.writeFileSync(out, z.generate({ type: 'nodebuffer', compression: 'DEFLATE' }));
  console.log('OK ->', out);
}

const check = new PizZip(fs.readFileSync(OUT[0], 'binary'));
const xml = check.files['word/document.xml'].asText();
for (const tag of ['{#kategori_list}', '{#foto_baris}', '{#ada}', '{#kosong}', '{#terlampir}', '{#boq}', '{#boq_kosong}', '{#bukan_boq}', '{#redaman_ada}', '{#redaman_list}', '{#port_list}', '{#adaFoto}', '{#bukan_foto}', '{%foto_path}', '{#input_list}', '{#ada_input}', '{#kosong_input}', '{%f1}', '{tematik}', '{nama_file}', '{boq_total_material}', '{boq_total_jasa}', '{#kml_ada}', '{#kml_list}', '{%foto}', '{#kml_kosong}']) {
  if (!xml.includes(tag)) throw new Error('Tag hilang: ' + tag);
}
console.log('VERIFIKASI TAG: semua loop ada di template_HEM.docx');