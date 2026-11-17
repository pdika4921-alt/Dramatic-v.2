/* =====================================================================
   util.js — Helper bersama (non-breaking)
   ---------------------------------------------------------------------
   • Wrapper window.fetch yang menyisipkan header X-CSRF-Token secara
     otomatis BILA <meta name="csrf-token"> tersedia. Aman: jika meta
     tidak ada, fetch berjalan persis seperti biasa (tidak merusak
     endpoint yang sudah ada).
   • Helper escapeHtml() untuk mencegah XSS pada konten dinamis.
   ===================================================================== */
(function () {
  'use strict';

  /* ── CSRF: baca token dari <meta name="csrf-token"> ─────── */
  function getCsrfToken() {
    var m = document.querySelector('meta[name="csrf-token"]');
    return m ? m.getAttribute('content') : '';
  }

  /* ── Bungkus fetch asli agar selalu kirim CSRF (jika ada) ── */
  var _nativeFetch = window.fetch ? window.fetch.bind(window) : null;

  if (_nativeFetch) {
    window.fetch = function (input, opts) {
      opts = opts || {};
      var method = (opts.method || (input && input.method) || 'GET').toUpperCase();
      var token = getCsrfToken();

      if (token && (method === 'POST' || method === 'PUT' || method === 'PATCH' || method === 'DELETE')) {
        opts.headers = opts.headers || {};
        if (!(opts.headers instanceof Headers)) {
          // pastikan bentuk plain object
          var h = {};
          if (opts.headers.forEach) {
            opts.headers.forEach(function (v, k) { h[k] = v; });
          } else if (typeof opts.headers === 'object') {
            h = Object.assign({}, opts.headers);
          }
          opts.headers = h;
        }
        if (!(opts.headers instanceof Headers) && !opts.headers['X-CSRF-Token']) {
          opts.headers['X-CSRF-Token'] = token;
        }
      }
      return _nativeFetch(input, opts);
    };
  }

  /* ── Escape HTML (cegah XSS saat menyuntikkan teks user) ── */
  window.escapeHtml = function (str) {
    if (str === null || str === undefined) return '';
    return String(str)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  };

  /* ── raw(): tandai sebuah nilai sebagai HTML yang SUDAH aman,
        sehingga esc() tidak akan meng-escape-nya lagi. Dipakai
        untuk interpolasi yang memang menghasilkan markup HTML
        (ikon, badge, helper yang sudah sanitasi sendiri). ── */
  window.raw = function (value) {
    return { __raw: String(value === null || value === undefined ? '' : value) };
  };

  /* ── esc: tagged template yang otomatis escape semua ${...}
        interpolasi KECUALI yang ditandai raw(). Markup literal
        tetap utuh. Dipakai untuk innerHTML guna mencegah XSS
        pada data dinamis (nama lokasi, nilai RAB, input user). ── */
  window.esc = function (strings) {
    var values = Array.prototype.slice.call(arguments, 1);
    var out = strings[0] || '';
    for (var i = 0; i < values.length; i++) {
      var v = values[i];
      if (v && typeof v === 'object' && v.__raw !== undefined) {
        out += v.__raw;
      } else {
        out += window.escapeHtml(v);
      }
      out += (strings[i + 1] || '');
    }
    return out;
  };
})();
