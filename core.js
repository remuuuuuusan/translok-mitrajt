/* Bagian bersama untuk semua dokumen. Jangan ubah kecuali perlu.
   Dokumen baru cukup menambah satu file di js/docs/ (lihat PANDUAN-TAMBAH-DOKUMEN.md). */
(function () {
  "use strict";
  var Core = (window.Core = { docs: [], BULAN: ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"],
    HARI: ["Minggu", "Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu"] });
  var A = Array.prototype;

  /* ---------- util ---------- */
  var $ = (Core.$ = function (id) { return document.getElementById(id); });
  Core.el = function (tag, props, kids) {
    var e = document.createElement(tag); props = props || {};
    Object.keys(props).forEach(function (k) {
      var v = props[k];
      if (k === "class") e.className = v;
      else if (k === "text") e.textContent = v;
      else if (k.slice(0, 2) === "on") e.addEventListener(k.slice(2), v);
      else if (v !== false && v != null) e.setAttribute(k, v === true ? "" : v);
    });
    (kids || []).forEach(function (c) { e.appendChild(typeof c === "string" ? document.createTextNode(c) : c); });
    return e;
  };
  var el = Core.el;
  Core.store = {
    get: function (k, d) { try { var v = localStorage.getItem("gd-" + k); return v == null ? d : JSON.parse(v); } catch (e) { return d; } },
    set: function (k, v) { try { localStorage.setItem("gd-" + k, JSON.stringify(v)); } catch (e) { /* abaikan */ } },
    del: function (k) { try { localStorage.removeItem("gd-" + k); } catch (e) { /* abaikan */ } }
  };
  function pad(n) { return String(n).padStart(2, "0"); }
  Core.todayISO = function () { var d = new Date(); return d.getFullYear() + "-" + pad(d.getMonth() + 1) + "-" + pad(d.getDate()); };
  Core.fmtDate = function (s) { if (!s) return ""; var p = s.split("-"); return p[2] + " " + Core.BULAN[+p[1] - 1] + " " + p[0]; };
  Core.hari = function (iso) { if (!iso) return ""; var p = iso.split("-"); return Core.HARI[new Date(+p[0], +p[1] - 1, +p[2]).getDay()]; };
  Core.addDays = function (iso, n) { var p = iso.split("-"), d = new Date(+p[0], +p[1] - 1, +p[2] + n); return d.getFullYear() + "-" + pad(d.getMonth() + 1) + "-" + pad(d.getDate()); };
  Core.diffDays = function (a, b) { var x = a.split("-"), y = b.split("-"); return Math.round((new Date(+y[0], +y[1] - 1, +y[2]) - new Date(+x[0], +x[1] - 1, +x[2])) / 864e5); };
  // terima "2026-09-17", "17 September 2026", "17/09/2026", atau nomor seri Excel
  Core.parseDate = function (v) {
    v = String(v == null ? "" : v).trim(); if (!v) return "";
    if (/^\d{4}-\d{2}-\d{2}/.test(v)) return v.slice(0, 10);
    if (/^\d{5}(\.\d+)?$/.test(v)) return new Date(Date.UTC(1899, 11, 30) + Math.floor(+v) * 864e5).toISOString().slice(0, 10);
    var m = /^(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})$/.exec(v);
    if (m) { var i = Core.BULAN.map(function (b) { return b.toLowerCase(); }).indexOf(m[2].toLowerCase()); if (i >= 0) return m[3] + "-" + pad(i + 1) + "-" + pad(m[1]); }
    m = /^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/.exec(v);
    return m ? m[3] + "-" + pad(m[2]) + "-" + pad(m[1]) : "";
  };
  Core.uniq = function (list) { var seen = {}; return list.filter(function (v) { if (!v || seen[v]) return false; seen[v] = 1; return true; }); };
  Core.firstWord = function (s) { return String(s || "").split(/\s+/)[0].replace(/[^A-Za-z]/g, "").toUpperCase(); };
  var firstWord = Core.firstWord;
  Core.rupiah = function (n) { return "RP. " + String(Math.round(n)).replace(/\B(?=(\d{3})+(?!\d))/g, ".") + ", -"; };
  function tb(n) {
    var s = ["", "satu", "dua", "tiga", "empat", "lima", "enam", "tujuh", "delapan", "sembilan", "sepuluh", "sebelas"];
    if (n < 12) return s[n];
    if (n < 20) return tb(n - 10) + " belas";
    if (n < 100) return tb(Math.floor(n / 10)) + " puluh" + (n % 10 ? " " + tb(n % 10) : "");
    if (n < 200) return "seratus" + (n > 100 ? " " + tb(n - 100) : "");
    if (n < 1000) return tb(Math.floor(n / 100)) + " ratus" + (n % 100 ? " " + tb(n % 100) : "");
    if (n < 2000) return "seribu" + (n > 1000 ? " " + tb(n - 1000) : "");
    if (n < 1e6) return tb(Math.floor(n / 1000)) + " ribu" + (n % 1000 ? " " + tb(n % 1000) : "");
    if (n < 1e9) return tb(Math.floor(n / 1e6)) + " juta" + (n % 1e6 ? " " + tb(n % 1e6) : "");
    return tb(Math.floor(n / 1e9)) + " miliar" + (n % 1e9 ? " " + tb(n % 1e9) : "");
  }
  Core.terbilang = function (n) { n = Math.round(n); if (n <= 0) return "Nol rupiah"; var t = tb(n); return t.charAt(0).toUpperCase() + t.slice(1) + " rupiah"; };

  /* ---------- baca Excel (.xlsx) ---------- */
  function colIdx(ref) { var m = /^([A-Z]+)/.exec(ref || ""), n = 0; if (!m) return 0; for (var i = 0; i < m[1].length; i++) n = n * 26 + m[1].charCodeAt(i) - 64; return n - 1; }
  Core.excel = {
    // hasil: [{name, rows:[[sel,sel,...],...]}]
    read: async function (buf) {
      var z = await JSZip.loadAsync(buf), dp = new DOMParser();
      async function xml(p) { var f = z.file(p); return f ? dp.parseFromString(await f.async("string"), "application/xml") : null; }
      var ss = [], sd = await xml("xl/sharedStrings.xml");
      if (sd) A.forEach.call(sd.getElementsByTagName("si"), function (si) { ss.push(A.map.call(si.getElementsByTagName("t"), function (t) { return t.textContent; }).join("")); });
      var wb = await xml("xl/workbook.xml"), rl = await xml("xl/_rels/workbook.xml.rels");
      if (!wb || !rl) throw new Error("Bukan file .xlsx yang valid.");
      var map = {}; A.forEach.call(rl.getElementsByTagName("Relationship"), function (r) { map[r.getAttribute("Id")] = r.getAttribute("Target"); });
      var out = [], sheets = A.slice.call(wb.getElementsByTagName("sheet"));
      for (var s = 0; s < sheets.length; s++) {
        var sh = sheets[s], rid = sh.getAttribute("r:id") || sh.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id");
        var tg = map[rid]; if (!tg) continue;
        var doc = await xml(tg.charAt(0) === "/" ? tg.slice(1) : "xl/" + tg), rows = [];
        if (!doc) continue;
        A.forEach.call(doc.getElementsByTagName("row"), function (row) {
          var arr = [];
          A.forEach.call(row.getElementsByTagName("c"), function (c) {
            var t = c.getAttribute("t"), v = "";
            if (t === "inlineStr") v = A.map.call(c.getElementsByTagName("t"), function (x) { return x.textContent; }).join("");
            else { var ve = c.getElementsByTagName("v")[0]; v = ve ? ve.textContent : ""; if (t === "s") v = ss[+v] || ""; }
            arr[colIdx(c.getAttribute("r"))] = v;
          });
          rows.push(arr);
        });
        out.push({ name: sh.getAttribute("name") || "", rows: rows });
      }
      return out;
    },
    sheet: function (sheets, re) { return sheets.filter(function (s) { return re.test(s.name); })[0]; },
    // peta judul kolom -> nomor kolom pada sebuah sheet (judul di baris pertama, huruf kecil)
    header: function (sheet) {
      var h = {}; (sheet.rows[0] || []).forEach(function (t, i) { var k = String(t || "").trim().toLowerCase().replace(/\s+/g, " "); if (k && !(k in h)) h[k] = i; });
      return h;
    },
    // cari kolom: kembalikan nomor kolom pertama yang judulnya memuat salah satu kata
    col: function (h, words) { for (var k in h) for (var i = 0; i < words.length; i++) if (k.indexOf(words[i]) >= 0) return h[k]; return -1; }
  };

  /* ---------- pegawai (dipakai semua dokumen) ---------- */
  var staffListeners = [];
  Core.staff = {
    list: [], info: "", custom: false, defaults: [],
    setDefaults: function (l) { this.defaults = l; if (!this.custom && !this.list.length) this.list = l; },
    find: function (key) { return this.list.filter(function (s) { return s[0] === key; })[0]; },
    onChange: function (fn) { staffListeners.push(fn); return function () { staffListeners = staffListeners.filter(function (f) { return f !== fn; }); }; },
    set: function (list, info, custom) { this.list = list; this.info = info; this.custom = custom; staffListeners.slice().forEach(function (f) { f(); }); },
    fromSheets: function (sheets) {
      var ref = Core.excel.sheet(sheets, /referensi|pegawai/i);
      if (!ref) throw new Error('Tidak ada sheet "Referensi Pegawai".');
      var staff = [], seen = {};
      ref.rows.forEach(function (r) { var k = String(r[0] || "").trim(); if (k && !seen[k]) { seen[k] = 1; staff.push([k, String(r[1] || k).trim(), String(r[2] || "").trim()]); } });
      if (!staff.length) throw new Error("Sheet Referensi Pegawai kosong.");
      return staff;
    },
    init: function () {
      var st = Core.store.get("pegawai", null);
      if (st && st.list && st.list.length) { this.list = st.list; this.info = st.info; this.custom = true; }
      else { this.list = this.defaults; this.info = "Memakai data pegawai bawaan (" + this.defaults.length + " orang)."; this.custom = false; }
    },
    useDefaults: function () { Core.store.del("pegawai"); this.set(this.defaults, "Memakai data pegawai bawaan (" + this.defaults.length + " orang).", false); },
    loadFile: async function (file) {
      var list = this.fromSheets(await Core.excel.read(await file.arrayBuffer()));
      var info = file.name + ": " + list.length + " pegawai.";
      Core.store.set("pegawai", { list: list, info: info });
      this.set(list, info, true);
    }
  };

  /* ---------- daftar data dari Excel (mitra, baris surat tugas) ---------- */
  // o: {key, kosong (teks bila kosong), parse(sheets) -> [item dengan .key]}
  Core.registry = function (o) {
    var ls = [], R = {
      list: [], info: o.kosong || "", custom: false,
      parse: o.parse,
      find: function (k) { return R.list.filter(function (i) { return i.key === k; })[0]; },
      onChange: function (fn) { ls.push(fn); return function () { ls = ls.filter(function (f) { return f !== fn; }); }; },
      set: function (list, info, custom) { R.list = list; R.info = info; R.custom = custom; ls.slice().forEach(function (f) { f(); }); },
      init: function () { var st = Core.store.get(o.key, null); if (st && st.list && st.list.length) { R.list = st.list; R.info = st.info; R.custom = true; } else if (o.dasar) { R.list = o.dasar(); R.info = o.dasarInfo(R.list); R.custom = false; } },
      reset: function () { Core.store.del(o.key); var d = o.dasar ? o.dasar() : []; R.set(d, d.length && o.dasarInfo ? o.dasarInfo(d) : (o.kosong || ""), false); },
      loadFile: async function (file) {
        var list = o.parse(await Core.excel.read(await file.arrayBuffer())), info = file.name + ": " + list.length + " " + (o.satuan || "baris") + "." + (o.catatan ? o.catatan() : "");
        Core.store.set(o.key, { list: list, info: info }); R.set(list, info, true);
      }
    };
    return R;
  };
  function cell(r, i) { return i >= 0 && r[i] != null ? String(r[i]).trim() : ""; }
  function mainSheet(sheets, needs) {
    var best = null;
    sheets.forEach(function (s) { var h = Core.excel.header(s); if (!best && needs.every(function (n) { return Core.excel.col(h, [n]) >= 0; })) best = { s: s, h: h }; });
    return best;
  }
  // Database mitra bawaan (Core.MITRA) + Excel opsional: Nama | Tanggal Berangkat | Lama (hari) | [NIK] [No.SPD] [Email]
  function normNama(s) { return String(s || "").trim().toLowerCase().replace(/\s+/g, " "); }
  function dasarMitra() {
    return (Core.MITRA || []).map(function (m, i) {
      var lama = m[6] || 0, b = m[5] || "";
      return { key: m[0] + "|" + i, nama: m[0], email: m[1] || "", kec: m[2] || "", nik: m[3] || "", nomor: m[4] || "", berangkat: b, lama: lama, kembali: b && lama ? Core.addDays(b, lama - 1) : "", biaya: 0 };
    });
  }
  Core.mitra = Core.registry({
    key: "mitra", satuan: "mitra", kosong: "Belum ada data mitra.",
    dasar: dasarMitra,
    dasarInfo: function (l) { return "Memakai database mitra bawaan (" + l.length + " orang). Unggah Excel (Nama, Email) untuk memilih mitra kegiatan ini saja."; },
    catatan: function () { return Core.mitraCatatan || ""; },
    parse: function (sheets) {
      var m = mainSheet(sheets, ["nama"]);
      if (!m) throw new Error("Tidak ada sheet dengan kolom Nama di baris pertama.");
      function K(words) { for (var i = 0; i < words.length; i++) { var x = Core.excel.col(m.h, [words[i]]); if (x >= 0) return x; } return -1; }
      var c = { n: K(["nama"]), k: K(["nik"]), no: K(["no.spd", "nomor spd", "no spd", "spd", "nomor"]), b: K(["berangkat", "mulai jalan"]), l: K(["lama", "jumlah ok"]), t: K(["kembali", "tiba"]), r: K(["biaya", "ok terbayar"]), e: K(["email", "e-mail"]) };
      var db = {}; dasarMitra().forEach(function (d) { db[normNama(d.nama)] = d; });
      var out = [], baru = 0, tglKosong = 0;
      m.s.rows.slice(1).forEach(function (r, idx) {
        var nama = cell(r, c.n).replace(/\s+/g, " "); if (!nama) return;
        var d = db[normNama(nama)] || null; if (!d) baru++;
        var nik = cell(r, c.k); if (/e\+/i.test(nik)) throw new Error("NIK " + nama + " terbaca sebagai angka ilmiah. Ubah kolom NIK menjadi format Teks di Excel.");
        var b = Core.parseDate(cell(r, c.b)) || (d ? d.berangkat : ""), t = Core.parseDate(cell(r, c.t)), lama = parseInt(cell(r, c.l), 10);
        if (!(lama > 0) && cell(r, c.b) && t) lama = Core.diffDays(Core.parseDate(cell(r, c.b)), t) + 1;
        if (!(lama > 0)) lama = d ? d.lama : 0;
        if (!b) tglKosong++;
        var biaya = parseInt(String(cell(r, c.r)).replace(/\D/g, ""), 10);
        out.push({ key: nama + "|" + idx, nama: nama, nik: nik || (d ? d.nik : ""), nomor: cell(r, c.no) || (d ? d.nomor : ""), berangkat: b, lama: lama > 0 ? lama : 0,
          kembali: b && lama > 0 ? Core.addDays(b, lama - 1) : "", biaya: biaya > 0 ? biaya : 0, email: cell(r, c.e) || (d ? d.email : "") });
      });
      Core.mitraCatatan = (baru ? " " + baru + " nama belum ada di database (NIK dan nomor SPD harus diisi di Excel)." : "") + (tglKosong ? " " + tglKosong + " belum punya tanggal berangkat." : "");
      if (!out.length) throw new Error("Tidak ada baris mitra.");
      return out;
    }
  });
  // Baris dari Excel Surat Tugas: dipakai Visum
  Core.baris = Core.registry({
    key: "baris-st", satuan: "surat tugas", kosong: "Belum ada data. Unggah Excel dengan kolom Tanggal Jalan, Tujuan, Asal.",
    parse: function (sheets) {
      var ref = Core.excel.sheet(sheets, /referensi|pegawai/i), m = mainSheet(sheets.filter(function (s) { return s !== ref; }), ["tanggal jalan", "tujuan"]);
      if (!m) throw new Error("Tidak ada sheet dengan kolom Tanggal Jalan dan Tujuan.");
      var c = { no: Core.excel.col(m.h, ["nomor"]), tj: Core.excel.col(m.h, ["tanggal jalan"]), tg: Core.excel.col(m.h, ["tugas"]), tu: Core.excel.col(m.h, ["tujuan"]), o: Core.excel.col(m.h, ["orang"]), a: Core.excel.col(m.h, ["asal"]) };
      var out = [];
      m.s.rows.slice(1).forEach(function (r, idx) {
        var tj = Core.parseDate(cell(r, c.tj)), tu = cell(r, c.tu); if (!tj || !tu) return;
        out.push({ key: (cell(r, c.no) || "baris") + "|" + idx, nomor: cell(r, c.no), tglJalan: tj, tugas: cell(r, c.tg), tujuan: tu, orang: cell(r, c.o), asal: cell(r, c.a) });
      });
      if (!out.length) throw new Error("Tidak ada baris dengan tanggal jalan dan tujuan.");
      return out;
    }
  });
  Core.emailMitra = function (nama) { var k = normNama(nama), l = Core.MITRA || []; for (var i = 0; i < l.length; i++) if (normNama(l[i][0]) === k) return l[i][1] || ""; return ""; };
  Core.pokOf = function (label) { var l = Core.POK || []; return l.filter(function (p) { return p.label === label; })[0] || l[0] || null; };
  Core.spdSufiks = function () { return Core.store.get("f-spdSufiks", null) || Core.FS.spdSufiks.default; };
  // Data visum mitra dari Excel (tanpa database mitra): Tanggal Berangkat | Lama (hari) | [Asal] [Tujuan] [Nama]
  Core.visumMitra = Core.registry({
    key: "visum-mitra", satuan: "perjalanan", kosong: "Belum ada data. Unggah Excel (Tanggal Berangkat, Lama) atau isi manual di bawah.",
    parse: function (sheets) {
      var m = mainSheet(sheets, ["berangkat", "lama"]);
      if (!m) throw new Error("Tidak ada sheet dengan kolom Tanggal Berangkat dan Lama (hari).");
      function K(w) { for (var i = 0; i < w.length; i++) { var x = Core.excel.col(m.h, [w[i]]); if (x >= 0) return x; } return -1; }
      var c = { b: K(["berangkat", "mulai jalan"]), l: K(["lama", "jumlah ok"]), a: K(["asal", "dari"]), t: K(["tujuan"]), n: K(["nama"]) }, out = [];
      m.s.rows.slice(1).forEach(function (r, idx) {
        var b = Core.parseDate(cell(r, c.b)), lama = parseInt(cell(r, c.l), 10); if (!b || !(lama > 0)) return;
        out.push({ key: "vm" + idx, berangkat: b, lama: lama, asal: cell(r, c.a), tujuan: cell(r, c.t), nama: cell(r, c.n) });
      });
      if (!out.length) throw new Error("Tidak ada baris dengan tanggal berangkat dan lama.");
      return out;
    }
  });
  Core.mitraV = function (m, tarif, v) {
    v = v || {}; var t = parseInt(String(tarif).replace(/\D/g, ""), 10) || 0;
    var b = m.berangkat || v.tglBerangkat || Core.todayISO(), lama = m.lama > 0 ? m.lama : (parseInt(v.lamaMitra, 10) || 1);
    return Object.assign({}, m, { berangkat: b, lama: lama, kembali: Core.addDays(b, lama - 1), biaya: m.biaya || lama * t,
      nomor: /^\d+$/.test(m.nomor) ? m.nomor + "/" + Core.spdSufiks() : (m.nomor || "(nomor SPD belum ada)"), email: m.email || Core.emailMitra(m.nama) });
  };

  /* ---------- PDF ---------- */
  var FONTKEYS = { R: "TimesRoman", B: "TimesRomanBold", BI: "TimesRomanBoldItalic", H: "Helvetica", HB: "HelveticaBold" };
  async function embedAll(doc) { var f = {}; for (var k in FONTKEYS) f[k] = await doc.embedFont(PDFLib.StandardFonts[FONTKEYS[k]]); return f; }
  Core.pdf = {
    f: null, _ready: null,
    init: function () {
      var self = this;
      if (!this._ready) this._ready = (async function () { self.f = await embedAll(await PDFLib.PDFDocument.create()); })();
      return this._ready;
    },
    ok: function () { return !!this.f; },
    clean: function (s) {
      return String(s == null ? "" : s).replace(/[‘’]/g, "'").replace(/[“”]/g, '"').replace(/[–—]/g, "-")
        .replace(/[^ -~ -ÿ]/g, "").replace(/\s+/g, " ").trim();
    },
    w: function (font, size, text) { return this.f ? this.f[font].widthOfTextAtSize(this.clean(text), size) : 0; },
    wrap: function (text, font, size, width) {
      var self = this, lines = [], cur = "";
      this.clean(text).split(" ").forEach(function (word) {
        if (!word) return;
        var t = cur ? cur + " " + word : word;
        if (!cur || self.w(font, size, t) <= width) cur = t; else { lines.push(cur); cur = word; }
      });
      if (cur) lines.push(cur);
      return lines;
    },
    render: async function (ops, title) {
      var L = PDFLib, doc = await L.PDFDocument.create(), f = await embedAll(doc), sizes = ops.pages || [{ w: 595.5, h: 842 }];
      var pgs = sizes.map(function (s) { return doc.addPage([s.w, s.h]); }), logo = null, imgs = new Map();
      doc.setTitle(this.clean(title || "Dokumen")); doc.setCreator("Generator Dokumen");
      for (var i = 0; i < ops.length; i++) {
        var o = ops[i], n = o.pg || 0, page = pgs[n], H = sizes[n].h, blk = L.rgb(0, 0, 0);
        if (o.k === "t") {
          var tw = f[o.font].widthOfTextAtSize(o.text, o.size), x = o.align === "c" ? o.x - tw / 2 : o.align === "r" ? o.x - tw : o.x;
          page.drawText(o.text, { x: x, y: H - o.base, size: o.size, font: f[o.font], color: o.blue ? L.rgb(0, 111 / 255, 192 / 255) : blk });
        } else if (o.k === "l") {
          page.drawLine({ start: { x: o.x1, y: H - o.y1 }, end: { x: o.x2, y: H - o.y2 }, thickness: o.th, color: blk });
        } else if (o.k === "r") {
          var ro = { x: o.x, y: H - o.y - o.h, width: o.w, height: o.h };
          if (o.fill != null) ro.color = L.rgb(o.fill, o.fill, o.fill);
          if (o.stroke) { ro.borderColor = blk; ro.borderWidth = o.stroke; }
          page.drawRectangle(ro);
        } else if (o.k === "img") {
          if (!logo) { var bin = atob(Core.LOGO.split(",")[1]), u = new Uint8Array(bin.length); for (var j = 0; j < bin.length; j++) u[j] = bin.charCodeAt(j); logo = await doc.embedPng(u); }
          page.drawImage(logo, { x: o.x, y: H - o.y - o.h, width: o.w, height: o.h });
        } else if (o.k === "photo") {
          var im = imgs.get(o.photo); if (!im) { im = o.photo.mime === "image/png" ? await doc.embedPng(o.photo.bytes) : await doc.embedJpg(o.photo.bytes); imgs.set(o.photo, im); }
          page.drawImage(im, { x: o.x, y: H - o.y - o.h, width: o.w, height: o.h });
        }
      }
      return await doc.save();
    },
    // gambar halaman ke-n dari ops ke dalam elemen pageEl
    paint: function (pageEl, ops, n) {
      var FAM = { R: ["400", "normal", "var(--font-doc)"], B: ["700", "normal", "var(--font-doc)"], BI: ["700", "italic", "var(--font-doc)"], H: ["400", "normal", "var(--font-sans)"], HB: ["700", "normal", "var(--font-sans)"] };
      pageEl.innerHTML = ""; n = n || 0;
      ops.forEach(function (o) {
        if ((o.pg || 0) !== n) return;
        var e;
        if (o.k === "t") {
          var F = FAM[o.font];
          e = el("span", { text: o.text, class: o.blue ? "blue" : "" });
          e.style.cssText = "font-family:" + F[2] + ";font-weight:" + F[0] + ";font-style:" + F[1] + ";font-size:" + o.size + "px;line-height:" + (1.107 * o.size) + "px;top:" + (o.base - 0.891 * o.size) + "px;left:" + o.x + "px;";
          if (o.align === "c") e.style.transform = "translateX(-50%)"; else if (o.align === "r") e.style.transform = "translateX(-100%)";
        } else if (o.k === "l") {
          e = el("div", { class: "rule" });
          var x = Math.min(o.x1, o.x2), y = Math.min(o.y1, o.y2), w = Math.max(Math.abs(o.x2 - o.x1), o.th), h = Math.max(Math.abs(o.y2 - o.y1), o.th);
          e.style.cssText = "left:" + (o.x1 === o.x2 ? x - o.th / 2 : x) + "px;top:" + (o.y1 === o.y2 ? y - o.th / 2 : y) + "px;width:" + w + "px;height:" + h + "px;";
        } else if (o.k === "r") {
          e = el("div", { class: "rule" });
          e.style.cssText = "left:" + o.x + "px;top:" + o.y + "px;width:" + o.w + "px;height:" + o.h + "px;box-sizing:border-box;" +
            (o.fill != null ? "background:rgb(" + Math.round(o.fill * 255) + "," + Math.round(o.fill * 255) + "," + Math.round(o.fill * 255) + ");" : "background:transparent;") + (o.stroke ? "border:" + o.stroke + "px solid #000;" : "");
        } else if (o.k === "img") {
          e = el("img", { src: Core.LOGO, alt: "" });
          e.style.cssText = "left:" + o.x + "px;top:" + o.y + "px;width:" + o.w + "px;height:" + o.h + "px;";
        } else if (o.k === "slot") {
          e = el("div", { class: "sigslot", text: "TTE" });
          e.style.cssText = "left:" + o.x + "px;top:" + o.y + "px;width:" + o.w + "px;height:" + o.h + "px;";
        } else if (o.k === "photo") {
          e = el("img", { src: o.photo.url, alt: "" });
          e.style.cssText = "left:" + o.x + "px;top:" + o.y + "px;width:" + o.w + "px;height:" + o.h + "px;";
        }
        if (e) pageEl.appendChild(e);
      });
    }
  };

  /* ---------- Word (.docx) dari ops yang sama dengan PDF ---------- */
  Core.docx = {
    render: async function (ops, title) {
      var sizes = ops.pages || [{ w: 595.5, h: 842 }], tw = function (pt) { return Math.round(pt * 20); };
      var FN = { R: "Times New Roman", B: "Times New Roman", BI: "Times New Roman", H: "Arial", HB: "Arial" };
      function esc(t) { return String(t).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;"); }
      var media = [], rels = [], rid = 0, imgCache = new Map(), id = 0;
      function addMedia(key, bytes, ext) {
        var m = imgCache.get(key); if (m) return m;
        rid++; m = { rid: "rId" + (rid + 10), name: "image" + rid + "." + ext }; media.push({ name: m.name, bytes: bytes }); imgCache.set(key, m); return m;
      }
      function dataBytes(url) { var b = atob(url.split(",")[1]), u = new Uint8Array(b.length); for (var i = 0; i < b.length; i++) u[i] = b.charCodeAt(i); return u; }
      function anchor(o, m) {
        id++; var cx = Math.round(o.w * 12700), cy = Math.round(o.h * 12700);
        return '<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="' + id + '" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/>' +
          '<wp:positionH relativeFrom="page"><wp:posOffset>' + Math.round(o.x * 12700) + '</wp:posOffset></wp:positionH><wp:positionV relativeFrom="page"><wp:posOffset>' + Math.round(o.y * 12700) + '</wp:posOffset></wp:positionV>' +
          '<wp:extent cx="' + cx + '" cy="' + cy + '"/><wp:wrapNone/><wp:docPr id="' + id + '" name="Gambar ' + id + '"/><wp:cNvGraphicFramePr/>' +
          '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
          '<pic:nvPicPr><pic:cNvPr id="' + id + '" name="' + m.name + '"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="' + m.rid + '"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>' +
          '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="' + cx + '" cy="' + cy + '"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>';
      }
      var body = [];
      for (var pg = 0; pg < sizes.length; pg++) {
        var S = sizes[pg], frames = "", shapes = "";
        for (var i = 0; i < ops.length; i++) {
          var o = ops[i]; if ((o.pg || 0) !== pg) continue;
          if (o.k === "t" && o.text) {
            var wd = Core.pdf.w(o.font, o.size, o.text), box = wd * 1.08 + 8, x = o.align === "c" ? o.x - box / 2 : o.align === "r" ? o.x - box : o.x;
            var lh = 1.107 * o.size, top = o.base - 0.891 * o.size;
            frames += '<w:p><w:pPr><w:framePr w:w="' + tw(box) + '" w:h="' + tw(lh) + '" w:hRule="exact" w:wrap="through" w:vAnchor="page" w:hAnchor="page" w:x="' + tw(x) + '" w:y="' + tw(top) + '"/>' +
              '<w:spacing w:before="0" w:after="0" w:line="' + tw(lh) + '" w:lineRule="exact"/><w:jc w:val="' + (o.align === "c" ? "center" : o.align === "r" ? "right" : "left") + '"/></w:pPr>' +
              '<w:r><w:rPr><w:rFonts w:ascii="' + FN[o.font] + '" w:hAnsi="' + FN[o.font] + '" w:cs="' + FN[o.font] + '"/>' + (o.font === "B" || o.font === "BI" || o.font === "HB" ? "<w:b/>" : "") + (o.font === "BI" ? "<w:i/>" : "") +
              (o.blue ? '<w:color w:val="006FC0"/>' : "") + '<w:sz w:val="' + Math.round(o.size * 2) + '"/><w:szCs w:val="' + Math.round(o.size * 2) + '"/></w:rPr><w:t xml:space="preserve">' + esc(o.text) + '</w:t></w:r></w:p>';
          } else if (o.k === "l") {
            shapes += '<w:r><w:pict><v:line xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" style="position:absolute;mso-position-horizontal-relative:page;mso-position-vertical-relative:page" from="' + o.x1 + 'pt,' + o.y1 + 'pt" to="' + o.x2 + 'pt,' + o.y2 + 'pt" strokecolor="black" strokeweight="' + o.th + 'pt"/></w:pict></w:r>';
          } else if (o.k === "r") {
            var g = o.fill != null ? Math.round(o.fill * 255) : 0, hx = ("0" + g.toString(16)).slice(-2);
            shapes += '<w:r><w:pict><v:rect xmlns:v="urn:schemas-microsoft-com:vml" style="position:absolute;margin-left:' + o.x + 'pt;margin-top:' + o.y + 'pt;width:' + o.w + 'pt;height:' + o.h + 'pt;mso-position-horizontal-relative:page;mso-position-vertical-relative:page" ' +
              (o.fill != null ? 'filled="t" fillcolor="#' + hx + hx + hx + '"' : 'filled="f"') + (o.stroke ? ' stroked="t" strokecolor="black" strokeweight="' + o.stroke + 'pt"' : ' stroked="f"') + '/></w:pict></w:r>';
          } else if (o.k === "img") {
            shapes += anchor(o, addMedia("logo", dataBytes(Core.LOGO), "png"));
          } else if (o.k === "photo") {
            shapes += anchor(o, addMedia(o.photo, o.photo.bytes, o.photo.mime === "image/png" ? "png" : "jpeg"));
          }
        }
        var sect = '<w:sectPr><w:pgSz w:w="' + tw(S.w) + '" w:h="' + tw(S.h) + '"/><w:pgMar w:top="0" w:right="0" w:bottom="0" w:left="0" w:header="0" w:footer="0" w:gutter="0"/></w:sectPr>';
        var last = pg === sizes.length - 1;
        body.push(frames + '<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="20" w:lineRule="exact"/>' + (last ? "" : sect) + '</w:pPr>' + shapes + '</w:p>');
        if (last) body.push(sect);
      }
      var NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"';
      var z = new JSZip();
      z.file("[Content_Types].xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Default Extension="jpeg" ContentType="image/jpeg"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/></Types>');
      z.file("_rels/.rels", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/></Relationships>');
      z.file("docProps/core.xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>' + esc(Core.pdf.clean(title || "Dokumen")) + '</dc:title><dc:creator>Generator Dokumen</dc:creator></cp:coreProperties>');
      z.file("word/styles.xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/><w:sz w:val="24"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>');
      var rl;
      rl = '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>';
      imgCache.forEach(function (m) { rl += '<Relationship Id="' + m.rid + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/' + m.name + '"/>'; });
      z.file("word/_rels/document.xml.rels", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + rl + '</Relationships>');
      media.forEach(function (m) { z.file("word/media/" + m.name, m.bytes); });
      z.file("word/document.xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ' + NS + '><w:body>' + body.join("") + '</w:body></w:document>');
      return await z.generateAsync({ type: "uint8array", compression: "DEFLATE", mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
    }
  };
  // pilihan format unduhan: pdf | docx | both
  Core.getFormat = function () { var f = Core.store.get("format", "pdf"); return f === "docx" || f === "both" ? f : "pdf"; };
  Core.fmtNama = function () { var f = Core.getFormat(); return f === "docx" ? "Word" : f === "both" ? "PDF + Word" : "PDF"; };
  Core.ui = Core.ui || {};
  // kotak pilihan format; onChange dipanggil saat berubah
  Core.ui.formatPicker = function (onChange) {
    var sel = el("select", { id: "fmt-" + Math.random().toString(36).slice(2, 6), "aria-label": "Format unduhan" }, [["pdf", "PDF"], ["docx", "Word (.docx)"], ["both", "PDF dan Word"]].map(function (o) { return el("option", { value: o[0], text: o[1] }); }));
    sel.value = Core.getFormat();
    sel.addEventListener("change", function () { Core.store.set("format", sel.value); if (onChange) onChange(); });
    return el("div", { class: "field" }, [el("label", { for: sel.id, text: "Format unduhan" }), sel]);
  };

  /* ---------- pembuat tata letak (dipakai di fungsi layout tiap dokumen) ----------
     Satuan pt, A4 = 595.5 x 842. y = garis dasar teks, diukur dari tepi ATAS kertas. */
  Core.painter = function () {
    var ops = [], P = { ops: ops, CX: 297.9, cur: 0 }; ops.pages = [{ w: 595.5, h: 842 }];
    function add(o) { o.pg = P.cur; ops.push(o); return P; }
    P.size = function (w, h) { ops.pages[P.cur] = { w: w, h: h }; P.W = w; P.H = h; return P; };
    P.newPage = function (w, h) { var s = ops.pages[P.cur]; ops.pages.push({ w: w || s.w, h: h || s.h }); P.cur++; P.W = ops.pages[P.cur].w; P.H = ops.pages[P.cur].h; return P; };
    P.W = 595.5; P.H = 842;
    // font: R B BI (Times), H HB (Helvetica/Arial)
    P.text = function (text, x, y, o) { o = o || {}; return add({ k: "t", text: Core.pdf.clean(text), x: x, base: y, font: o.font || "R", size: o.size || 12, align: o.align || "l", blue: !!o.blue }); };
    // sama seperti text, tetapi y = tepi BAWAH kotak teks (angka yang diukur dari PDF contoh)
    P.b = function (text, x, bottom, o) { o = o || {}; return P.text(text, x, bottom - 0.216 * (o.size || 12), o); };
    P.center = function (text, y, o) { return P.text(text, P.CX, y, Object.assign({}, o, { align: "c" })); };
    P.right = function (text, x, y, o) { return P.text(text, x, y, Object.assign({}, o, { align: "r" })); };
    P.hline = function (x1, x2, y, th) { return add({ k: "l", x1: x1, x2: x2, y1: y, y2: y, th: th || 0.6 }); };
    P.vline = function (x, y1, y2, th) { return add({ k: "l", x1: x, x2: x, y1: y1, y2: y2, th: th || 0.6 }); };
    P.line = function (x1, x2, y, th) { return P.hline(x1, x2, y, th); };
    P.rect = function (x, y, w, h, o) { o = o || {}; return add({ k: "r", x: x, y: y, w: w, h: h, fill: o.fill == null ? null : o.fill, stroke: o.stroke || 0 }); };
    P.slot = function (x, y, w, h) { return add({ k: "slot", x: x, y: y, w: w, h: h }); };
    P.logo = function (x, y, w, h) { return add({ k: "img", x: x, y: y, w: w, h: h }); };
    // foto {bytes,mime,w,h,url} dimuat pas ke dalam kotak, tetap proporsional dan di tengah
    P.photo = function (img, x, y, w, h) {
      if (!img) return P; var s = Math.min(w / img.w, h / img.h), iw = img.w * s, ih = img.h * s;
      return add({ k: "photo", photo: img, x: x + (w - iw) / 2, y: y + (h - ih) / 2, w: iw, h: ih });
    };
    P.width = function (text, font, size) { return Core.pdf.w(font || "R", size || 12, text); };
    P.wrap = function (text, width, o) { o = o || {}; return Core.pdf.wrap(text, o.font || "R", o.size || 12, width); };
    P.lines = function (arr, x, y, o) { o = o || {}; var lh = o.lh || 17.2; arr.forEach(function (t, i) { P.text(t, x, y + i * lh, o); }); return y + arr.length * lh; };
    P.para = function (text, x, y, width, o) { return P.lines(P.wrap(text, width, o), x, y, o); };
    // paragraf rata kiri-kanan; baris terakhir rata kiri. Mengembalikan y berikutnya
    P.justify = function (text, x, y, width, o) {
      o = o || {}; var lh = o.lh || 17.2, font = o.font || "R", size = o.size || 12, lines = P.wrap(text, width, o);
      lines.forEach(function (ln, i) {
        var yy = y + i * lh, words = ln.split(" ");
        if (i === lines.length - 1 || words.length < 2) { P.text(ln, x, yy, o); return; }
        var total = words.reduce(function (s, w) { return s + Core.pdf.w(font, size, w); }, 0), gap = (width - total) / (words.length - 1), cx = x;
        words.forEach(function (w) { P.text(w, cx, yy, o); cx += Core.pdf.w(font, size, w) + gap; });
      });
      return y + lines.length * lh;
    };
    P.kop = function () {
      P.logo(57.4, 56.6, 75, 60);
      P.b("BADAN PUSAT STATISTIK", 133.9, 86.8, { font: "BI", size: 18, blue: true });
      P.b("KOTA JAKARTA TIMUR", 133.9, 108.2, { font: "BI", size: 18, blue: true });
      return P;
    };
    return P;
  };

  /* ---------- komponen UI ---------- */
  Core.ui = Core.ui || {};
  // pratinjau semua halaman, bertumpuk
  Core.ui.preview = function () {
    var box = el("div", { class: "sheets" }), shown = [];
    function fit() { shown.forEach(function (s) { var k = s.sheet.clientWidth / s.w; s.page.style.transform = "scale(" + k + ")"; s.sheet.style.height = (s.h * k) + "px"; }); }
    if (window.ResizeObserver) new ResizeObserver(fit).observe(box); else window.addEventListener("resize", fit);
    return {
      el: box,
      show: function (ops) {
        box.innerHTML = ""; shown = [];
        (ops.pages || [{ w: 595.5, h: 842 }]).forEach(function (s, i) {
          var page = el("div", { class: "page" }); page.style.width = s.w + "px"; page.style.height = s.h + "px";
          var sheet = el("div", { class: "sheet" }, [page]); box.appendChild(sheet);
          Core.pdf.paint(page, ops, i); shown.push({ sheet: sheet, page: page, w: s.w, h: s.h });
        });
        fit();
      }
    };
  };

  // Pemilih daftar (pegawai, mitra, baris). opts: {reg, keyOf, label(item)->[judul,sub], selected, onChange(sel, prev), judul, hint, single, upload}
  Core.ui.picker = function (opts) {
    var reg = opts.reg, selected = (opts.selected || []).filter(function (k) { return !!findItem(k); }), preview = selected[0] || null;
    function findItem(k) { return reg.list.filter(function (i) { return opts.keyOf(i) === k; })[0]; }
    var q = el("input", { type: "search", placeholder: "Cari", "aria-label": "Cari " + opts.judul }), list = el("div", { class: "list" }), chips = el("div", { class: "chips" });
    var tools = el("div", { class: "ptools" }), box = el("div", { class: "pick", role: "group" }, [q, list]);
    var root = el("div", { class: "field" }, [el("span", { class: "lbl", text: opts.judul })]);
    var info = null;
    if (opts.upload) {
      var file = el("input", { type: "file", accept: ".xlsx", "aria-label": "Unggah Excel " + opts.judul }), reset = el("button", { type: "button", class: "ghost", text: "Kosongkan", hidden: !reg.custom, onclick: function () { reg.reset(); } });
      info = el("span", { class: "hint", text: reg.info });
      file.addEventListener("change", async function () {
        var f = file.files && file.files[0]; if (!f) return;
        var ok = false;
        try { await reg.loadFile(f); ok = true; } catch (e) { info.textContent = "File tidak bisa dipakai. " + (e && e.message ? e.message : ""); info.className = "hint err"; }
        file.value = "";
        if (ok && opts.onUpload && !opts.single && reg.list.length) {
          selected = reg.list.map(opts.keyOf); if (opts.single) selected = selected.slice(0, 1);
          fix(); drawList(); drawChips(); emit();
          setTimeout(function () { opts.onUpload(); }, 50);
        }
      });
      var tb = opts.templat ? Core.ui.templateBtn(opts.templat) : null; root.appendChild(el("div", { class: "tpl" }, [file, reset].concat(tb ? [tb] : []))); root.appendChild(info);
      reg.onChange(function () { if (!document.body.contains(root)) return; info.textContent = reg.info; info.className = "hint"; reset.hidden = !reg.custom; });
    }
    if (!opts.single) {
      tools.appendChild(el("button", { type: "button", class: "ghost", text: "Pilih semua yang tampil", onclick: function () { shown().forEach(function (i) { var k = opts.keyOf(i); if (selected.indexOf(k) < 0) selected.push(k); }); fix(); drawList(); drawChips(); emit(); } }));
      tools.appendChild(el("button", { type: "button", class: "ghost", text: "Kosongkan pilihan", onclick: function () { selected = []; fix(); drawList(); drawChips(); emit(); } }));
      root.appendChild(tools);
    }
    root.appendChild(box); root.appendChild(chips);
    if (opts.hint) root.appendChild(el("span", { class: "hint", text: opts.hint }));
    var count = el("span", { class: "hint" }); root.appendChild(count);
    function shown() { var term = q.value.trim().toLowerCase(); return reg.list.filter(function (i) { var l = opts.label(i); return !term || (l[0] + " " + l[1]).toLowerCase().indexOf(term) >= 0; }); }
    function emit() { if (opts.onChange) opts.onChange(selected.slice(), preview); }
    function fix() { if (!preview || selected.indexOf(preview) < 0) preview = selected[0] || null; count.textContent = selected.length ? selected.length + " dipilih." : ""; }
    function drawChips() {
      chips.innerHTML = "";
      if (selected.length > 12) { chips.appendChild(el("span", { class: "hint", text: "Klik satu nama di daftar untuk melihat pratinjau." })); return; }
      selected.forEach(function (k) {
        var it = findItem(k); if (!it) return;
        var nb = el("button", { type: "button", class: "nmb", text: opts.label(it)[0], "aria-label": "Pratinjau " + opts.label(it)[0], onclick: function () { preview = k; drawChips(); emit(); } });
        var x = el("button", { type: "button", text: "×", "aria-label": "Hapus " + opts.label(it)[0], onclick: function () { selected.splice(selected.indexOf(k), 1); fix(); drawList(); drawChips(); emit(); } });
        chips.appendChild(el("span", { class: "chip" + (k === preview ? " on" : "") }, [nb, x]));
      });
    }
    function drawList() {
      list.innerHTML = ""; var items = shown();
      if (!items.length) { list.appendChild(el("div", { class: "none", text: reg.list.length ? "Tidak ada yang cocok." : (reg.info || "Daftar kosong.") })); return; }
      items.forEach(function (it) {
        var k = opts.keyOf(it), l = opts.label(it), c = el("input", { type: opts.single ? "radio" : "checkbox", name: "pick-" + opts.judul });
        c.checked = selected.indexOf(k) >= 0;
        c.addEventListener("change", function () {
          if (opts.single) selected = c.checked ? [k] : [];
          else { var i = selected.indexOf(k); if (c.checked && i < 0) selected.push(k); if (!c.checked && i >= 0) selected.splice(i, 1); }
          preview = selected.indexOf(k) >= 0 ? k : preview;
          fix(); drawChips(); emit();
        });
        var mail = opts.email ? opts.email(it) : "", kids = [el("span", { class: "nm", text: l[0] }), el("span", { class: "jb", text: l[1] })];
        if (mail) kids.push(el("a", { class: "mail", href: "mailto:" + mail + "?subject=" + encodeURIComponent("Dokumen perjalanan dinas - " + l[0]), text: "\u2709 " + mail, onclick: function (e) { e.stopPropagation(); } }));
        list.appendChild(el("label", {}, [c, el("span", {}, kids)]));
      });
    }
    q.addEventListener("input", drawList);
    reg.onChange(function () {
      if (!document.body.contains(root)) return;
      selected = selected.filter(function (k) { return !!findItem(k); }); fix(); drawList(); drawChips(); emit();
    });
    fix(); drawList(); drawChips();
    return { el: root, selected: function () { return selected.slice(); }, preview: function () { return preview; } };
  };
  Core.ui.staffPicker = function (opts) {
    return Core.ui.picker({ reg: Core.staff, keyOf: function (s) { return s[0]; }, label: function (s) { return [s[1], s[2]]; }, judul: "Nama pegawai",
      hint: opts.single ? "" : "Pilih lebih dari satu nama untuk membuat beberapa dokumen sekaligus.", selected: opts.selected, onChange: opts.onChange, single: opts.single });
  };

  /* ---------- unduh ---------- */
  async function saveFile(dl, name, data) {
    if (dl) return dl.save({ filename: name, data: data });
    var blob = data instanceof Blob ? data : new Blob([data]), a = document.createElement("a");
    a.href = URL.createObjectURL(blob); a.download = name; document.body.appendChild(a); a.click(); a.remove();
    setTimeout(function () { URL.revokeObjectURL(a.href); }, 4000);
    return { status: "saved" };
  }
  // tombol "Unduh template Excel"
  Core.ui.templateBtn = function (key) {
    var t = (Core.TEMPLATES || {})[key]; if (!t) return null;
    return el("button", { type: "button", class: "ghost", text: "Unduh template Excel", onclick: async function () {
      var dl = null; try { dl = window.claude ? await window.claude.use("downloads") : null; } catch (e) { dl = null; }
      var bin = atob(t.b64), u8 = new Uint8Array(bin.length); for (var i = 0; i < bin.length; i++) u8[i] = bin.charCodeAt(i);
      try { await saveFile(dl, t.nama, u8); } catch (e) {}
    } });
  };
  // items: [{name, ops, title}]. Mengembalikan teks hasil. Lempar error bila gagal.
  Core.generate = async function (items, zipName) {
    var dl = null;
    try { dl = window.claude ? await window.claude.use("downloads") : null; } catch (e) { dl = null; }
    if (window.claude && !dl) throw new Error("Unduhan tidak tersedia di tampilan ini. Buka halaman ini langsung di claude.ai.");
    var files = [], fm = Core.getFormat();
    for (var i = 0; i < items.length; i++) {
      if (fm !== "docx") files.push({ name: items[i].name, bytes: await Core.pdf.render(items[i].ops, items[i].title) });
      if (fm !== "pdf") files.push({ name: items[i].name.replace(/\.pdf$/i, "") + ".docx", bytes: await Core.docx.render(items[i].ops, items[i].title) });
    }
    try {
      if (files.length === 1) await saveFile(dl, files[0].name, files[0].bytes);
      else {
        var z = new JSZip(); files.forEach(function (f) { z.file(f.name, f.bytes); });
        await saveFile(dl, zipName || "Dokumen_" + files.length + ".zip", await z.generateAsync({ type: "blob" }));
      }
    } catch (e) { if (e && e.code === "declined") return "Unduhan dibatalkan."; throw e; }
    return files.length === 1 ? (fm === "docx" ? "Word disimpan." : "PDF disimpan.") : files.length + " file disimpan dalam ZIP.";
  };

  // baca foto dari <input type=file>, kecilkan, simpan sebagai JPEG
  Core.readPhoto = function (file) {
    return new Promise(function (resolve, reject) {
      var url = URL.createObjectURL(file), img = new Image();
      img.onload = function () {
        var s = Math.min(1, 1100 / Math.max(img.width, img.height)), w = Math.round(img.width * s), h = Math.round(img.height * s), c = document.createElement("canvas");
        c.width = w; c.height = h; var g = c.getContext("2d"); g.fillStyle = "#fff"; g.fillRect(0, 0, w, h); g.drawImage(img, 0, 0, w, h); URL.revokeObjectURL(url);
        c.toBlob(function (b) { if (!b) return reject(new Error("Foto tidak bisa dibaca.")); b.arrayBuffer().then(function (buf) { resolve({ bytes: new Uint8Array(buf), mime: "image/jpeg", w: w, h: h, url: c.toDataURL("image/jpeg", 0.85) }); }); }, "image/jpeg", 0.85);
      };
      img.onerror = function () { URL.revokeObjectURL(url); reject(new Error("Foto tidak bisa dibaca.")); };
      img.src = url;
    });
  };

  /* ---------- dokumen sederhana: isi daftar field dan fungsi layout ---------- */
  var SUMBER = {
    pegawai: { vkey: "pegawai", judul: "Nama pegawai", reg: function () { return Core.staff; }, keyOf: function (s) { return s[0]; }, label: function (s) { return [s[1], s[2]]; },
      toV: function (s) { return s ? { key: s[0], nama: s[1], jabatan: s[2] } : { key: "", nama: "(pilih nama pegawai)", jabatan: "" }; }, nama: function (v) { return v.pegawai.nama; } },
    mitra: { vkey: "mitra", judul: "Mitra", upload: true, templat: "mitra", email: function (m) { return m.email || Core.emailMitra(m.nama); }, reg: function () { return Core.mitra; }, keyOf: function (m) { return m.key; },
      label: function (m) { return [m.nama, (m.berangkat ? Core.fmtDate(m.berangkat) : "tanggal belum ada") + (m.lama ? " · " + m.lama + " hari" : "") + (m.nik ? "" : " · NIK belum ada")]; },
      toV: function (m, v) {
        var today = Core.todayISO();
        return Core.mitraV(m || { nama: "Nama Mitra (contoh)", nik: "3175000000000000", nomor: "0000/539184-92800/TRANSLOK-2902/09/2026", berangkat: today, kembali: today, lama: 1, biaya: 0 }, v.tarif || 0, v);
      }, nama: function (v) { return v.mitra.nama; } },
    visumMitra: { vkey: "vm", judul: "Data perjalanan (Excel)", upload: true, templat: "visumMitra", reg: function () { return Core.visumMitra; }, keyOf: function (b) { return b.key; },
      label: function (b) { return [(b.nama || "Perjalanan") + " · " + Core.fmtDate(b.berangkat), b.lama + " hari" + (b.tujuan ? " · " + b.tujuan : "")]; },
      toV: function (b) { return b || null; }, nama: function (v) { return v.vm ? v.vm.nama : ""; } },
    baris: { vkey: "baris", judul: "Data surat tugas (Excel)", upload: true, templat: "visum", reg: function () { return Core.baris; }, keyOf: function (b) { return b.key; },
      label: function (b) { return [b.tujuan, Core.fmtDate(b.tglJalan) + (b.asal ? " · dari " + b.asal : "")]; },
      toV: function (b) { return b || null; }, nama: function (v) { return v.baris ? v.baris.orang : ""; } }
  };

  Core.simpleDoc = function (spec) {
    Core.specs = Core.specs || {}; Core.specs[spec.id] = spec;
    var sumberKey = spec.orang || (spec.pegawai ? "pegawai" : null), S = sumberKey ? SUMBER[sumberKey] : null;
    Core.register({
      id: spec.id, nama: spec.nama, deskripsi: spec.deskripsi,
      mount: function (root) {
        var fields = (spec.fields || []).slice(), ctrl = {}, status = el("div", { class: "status", role: "status", "aria-live": "polite" });
        if (spec.orang === "mitra" && !spec.tanpaTambahan && !fields.some(function (f) { return f.id === "spdSufiks"; })) fields = fields.concat([{ id: "tglBerangkat", label: "Tanggal berangkat (untuk mitra yang belum punya tanggal)", type: "date" }, { id: "lamaMitra", label: "Lama perjalanan, hari (untuk mitra yang belum punya)", default: "1" }], Core.F(["spdSufiks"]));
        var form = el("form", { class: "card", autocomplete: "off" }), picker = null, preview = Core.ui.preview(), groups = {};
        var go = el("button", { type: "submit", class: "primary", text: "Unduh PDF" }), gabung = null;
        function say(m, err) { status.textContent = m; status.className = "status" + (err ? " err" : ""); }

        function tableControl(f, onChange) {
          var rows = [], box = el("div", { class: "trows" }), add = el("button", { type: "button", class: "ghost", text: f.tambah || "Tambah baris", onclick: function () { addRow({}); onChange(); } });
          function addRow(init) {
            var r = { vals: {}, ctl: {} }, card = el("div", { class: "trow" });
            f.kolom.forEach(function (k) {
              var c, id = "t-" + spec.id + "-" + f.id + "-" + rows.length + "-" + k.id;
              if (k.type === "foto") {
                var th = el("img", { class: "thumb", alt: "", hidden: true }), inp = el("input", { type: "file", accept: "image/*", id: id }), del = el("button", { type: "button", class: "ghost", text: "Hapus foto", hidden: true });
                inp.addEventListener("change", async function () { var fl = inp.files && inp.files[0]; if (!fl) return; try { r.vals[k.id] = await Core.readPhoto(fl); th.src = r.vals[k.id].url; th.hidden = false; del.hidden = false; onChange(); } catch (e) { say(e.message, true); } });
                del.addEventListener("click", function () { r.vals[k.id] = null; th.hidden = true; del.hidden = true; inp.value = ""; onChange(); });
                card.appendChild(el("div", { class: "field" }, [el("label", { for: id, text: k.label }), inp, th, del])); return;
              }
              c = k.type === "textarea" ? el("textarea", { id: id }) : el("input", { type: k.type === "date" ? "date" : "text", id: id, placeholder: k.placeholder || "" });
              var dv = init[k.id] != null ? init[k.id] : (k.default === "hariini" ? Core.todayISO() : k.default); if (dv != null) c.value = dv;
              r.ctl[k.id] = c; c.addEventListener("input", onChange);
              card.appendChild(el("div", { class: "field" }, [el("label", { for: id, text: k.label }), c]));
            });
            card.appendChild(el("button", { type: "button", class: "ghost", text: "Hapus baris", onclick: function () { rows.splice(rows.indexOf(r), 1); card.remove(); onChange(); } }));
            r.card = card; rows.push(r); box.appendChild(card);
          }
          (f.default && f.default.length ? f.default : [{}]).forEach(addRow);
          return { el: el("div", { class: "field" }, [el("span", { class: "lbl", text: f.label }), box, add]), get: function () {
            return rows.map(function (r) { var o = {}; f.kolom.forEach(function (k) { if (k.type === "foto") o[k.id] = r.vals[k.id] || null; else { var v = r.ctl[k.id].value.trim(); o[k.id] = v; if (k.type === "date") o[k.id + "Teks"] = Core.fmtDate(v); } }); return o; });
          } };
        }

        if (S) {
          picker = Core.ui.picker({ reg: S.reg(), keyOf: S.keyOf, label: S.label, judul: S.judul, upload: S.upload, selected: spec.pegawaiAwal ? [spec.pegawaiAwal] : [], single: spec.satu, email: S.email, templat: S.templat, onChange: refresh,
            onUpload: function () { if (form.requestSubmit) form.requestSubmit(); else form.dispatchEvent(new Event("submit", { cancelable: true })); },
            hint: spec.satu ? "" : (S.upload ? "" : "Pilih lebih dari satu nama untuk membuat beberapa dokumen sekaligus.") });
          form.appendChild(picker.el);
        }
        fields.forEach(function (f) {
          var c, other = null, id = "f-" + spec.id + "-" + f.id, row, tc = null;
          if (f.type === "tabel") { tc = tableControl(f, refresh); ctrl[f.id] = { tc: tc, f: f }; row = tc.el; }
          else {
            if (f.type === "textarea") c = el("textarea", { id: id });
            else if (f.type === "select") {
              c = el("select", { id: id }, (f.options || []).map(function (o) { return el("option", { value: o, text: o }); }));
              if (f.other) { c.appendChild(el("option", { value: "__other", text: "Lainnya..." })); other = el("input", { type: "text", placeholder: "Isi manual", hidden: true }); }
            } else if (f.type === "checkbox") c = el("input", { type: "checkbox", id: id });
            else c = el("input", { type: f.type === "date" ? "date" : "text", id: id, placeholder: f.placeholder || "" });
            var saved = f.simpan ? Core.store.get("f-" + f.id, null) : null;
            if (f.type === "date") c.value = saved || (f.default === "hariini" || !f.default ? Core.todayISO() : f.default);
            else if (f.type === "checkbox") c.checked = saved != null ? !!saved : !!f.default;
            else if (saved != null) { if (f.type === "select" && (f.options || []).indexOf(saved) < 0) { c.value = "__other"; other.value = saved; } else c.value = saved; }
            else if (f.default != null) c.value = f.default;
            ctrl[f.id] = { c: c, other: other, f: f };
            if (f.simpan) c.addEventListener("change", function () { Core.store.set("f-" + f.id, f.type === "checkbox" ? c.checked : (f.type === "select" && c.value === "__other" ? other.value.trim() : c.value)); });
            row = f.type === "checkbox" ? el("label", { class: "check", for: id }, [c, el("span", { text: f.label })])
              : el("div", { class: "field" }, [el("label", { for: id, text: f.label }), c].concat(other ? [other] : []));
            if (f.simpan && other) other.addEventListener("change", function () { Core.store.set("f-" + f.id, other.value.trim()); });
          }
          if (f.grup) {
            if (!groups[f.grup]) { var inner = el("div", { class: "inner" }); groups[f.grup] = inner; form.appendChild(el("details", { class: "set" }, [el("summary", { text: f.grup }), inner])); }
            groups[f.grup].appendChild(row);
          } else form.appendChild(row);
        });
        var nomorIn = null;
        if (spec.nomor) {
          nomorIn = el("input", { type: "text", inputmode: "numeric", maxlength: "6", "aria-label": "Nomor urut", id: "f-" + spec.id + "-nomor" });
          nomorIn.value = Core.store.get("nomor-" + spec.id, String(spec.nomor.awal || 1));
          nomorIn.addEventListener("input", function () { this.value = this.value.replace(/\D/g, ""); });
          form.appendChild(el("div", { class: "field" }, [el("label", { for: nomorIn.id, text: "Nomor" }),
            el("div", { class: "nomor" }, [el("span", { text: spec.nomor.awalan || "" }), nomorIn, el("span", { text: spec.nomor.akhiran || "" })])]));
        }
        if (!spec.satu && (S || spec.salinan)) {
          gabung = el("input", { type: "checkbox", id: "f-" + spec.id + "-gabung" }); gabung.checked = !!spec.gabungAwal;
          form.appendChild(el("label", { class: "check", for: gabung.id }, [gabung, el("span", { text: "Gabungkan semua halaman menjadi satu file (mudah dicetak sekaligus)." })]));
        }
        form.appendChild(Core.ui.formatPicker(refresh));
        form.appendChild(el("div", { class: "actions" }, [go, status]));
        root.appendChild(el("div", { class: "grid" }, [form, el("aside", {}, [el("div", { class: "pv-head" }, [el("span", { class: "lbl", text: "Pratinjau" }), el("span", { class: "sub", id: "pv-" + spec.id })]), preview.el])]));

        function value(f) {
          var o = ctrl[f.id];
          if (f.type === "tabel") return o.tc.get();
          if (f.type === "checkbox") return o.c.checked;
          if (f.type === "select" && o.c.value === "__other") return o.other.value.trim();
          return o.c.value.trim();
        }
        function values(key, offset) {
          var v = {};
          fields.forEach(function (f) { v[f.id] = value(f); if (f.type === "date") v[f.id + "Teks"] = Core.fmtDate(v[f.id]); });
          if (spec.nomor) { var n = parseInt(nomorIn.value, 10); v.nomor = (spec.nomor.awalan || "") + (isNaN(n) ? "..." : n + offset) + (spec.nomor.akhiran || ""); }
          if (S) { var it = key ? S.reg().list.filter(function (i) { return S.keyOf(i) === key; })[0] : null; v[S.vkey] = S.toV(it, v); }
          return v;
        }
        function keys() {
          if (!S) return [null];
          var sel = picker.selected();
          if (sel.length || !spec.salinan) return sel;
          var n = Math.max(1, Math.min(200, parseInt(value(fields.filter(function (f) { return f.id === spec.salinan; })[0]), 10) || 1)), a = []; for (var i = 0; i < n; i++) a.push(null); return a;
        }
        function one(k, off) { var p = Core.painter(); spec.layout(values(k, off), p); return p.ops; }
        function build(ks, combined) {
          var p = Core.painter(); ks.forEach(function (k, i) { if (combined && i) p.newPage(); spec.layout(values(k, i), p); }); return p.ops;
        }
        function refresh() {
          Object.keys(ctrl).forEach(function (id) { var o = ctrl[id]; if (o.other) o.other.hidden = o.c.value !== "__other"; });
          if (!Core.pdf.ok()) return;
          var ks = keys(), pk = S ? (picker.preview() || null) : null, off = Math.max(0, ks.indexOf(pk));
          preview.show(one(pk, off));
          var many = ks.length > 1, comb = many && gabung && gabung.checked;
          var fn = Core.fmtNama(); go.textContent = many ? (comb ? "Unduh " + fn + " gabungan (" + ks.length + " halaman)" : "Unduh ZIP (" + ks.length + " dokumen)") : "Unduh " + fn;
          var pv = $("pv-" + spec.id); if (pv) pv.textContent = many ? "dokumen " + (off + 1) + " dari " + ks.length : "";
        }
        form.addEventListener("input", refresh); form.addEventListener("change", refresh);
        Core.pdf.init().then(refresh, function () { say("Pustaka PDF gagal dimuat. Muat ulang halaman.", true); });
        refresh();

        form.addEventListener("submit", async function (ev) {
          ev.preventDefault();
          var ks = keys();
          if (S && !ks.length) return say("Pilih minimal satu " + (S.judul.toLowerCase().indexOf("pegawai") >= 0 ? "nama pegawai" : "data") + " di daftar.", true);
          for (var i = 0; i < fields.length; i++) {
            var f = fields[i]; if (f.type === "checkbox" || f.opsional || f.type === "tabel") continue;
            if (!value(f)) return say("Isi \"" + f.label + "\".", true);
          }
          if (spec.nomor && !/^\d+$/.test(nomorIn.value)) return say("Nomor harus berupa angka.", true);
          if (spec.cek) { var msg0 = spec.cek(values(ks[0], 0)); if (msg0) return say(msg0, true); }
          if (!Core.pdf.ok()) return say("Pustaka PDF belum siap.", true);
          go.disabled = true; say("Membuat dokumen...");
          try {
            var items, comb = ks.length > 1 && gabung && gabung.checked;
            if (comb || (spec.satuFile && ks.length >= 1)) items = [{ name: spec.id + (ks.length > 1 ? "_" + ks.length + "halaman" : "") + ".pdf", ops: build(ks, true), title: spec.nama }];
            else {
              var seen = {};
              items = ks.map(function (k, i) {
                var v = values(k, i), base = spec.namaFile ? spec.namaFile(v, i) : spec.id + (S && S.nama(v) ? "_" + firstWord(S.nama(v)) : (ks.length > 1 ? "_" + (i + 1) : ""));
                base = String(base).replace(/\.pdf$/i, "").replace(/[^\w.\-]+/g, "_"); seen[base] = (seen[base] || 0) + 1;
                return { name: base + (seen[base] > 1 ? "_" + seen[base] : "") + ".pdf", ops: one(k, i), title: spec.nama + (S && S.nama(v) ? " " + S.nama(v) : "") };
              });
            }
            var m = await Core.generate(items, spec.id + "_" + items.length + ".zip");
            if (spec.nomor) Core.store.set("nomor-" + spec.id, nomorIn.value);
            say(m);
          } catch (e) { say("Gagal membuat atau menyimpan file" + (e && e.message ? ": " + e.message : "."), true); }
          finally { go.disabled = false; }
        });
      }
    });
  };

  /* ---------- daftar dokumen & halaman depan ---------- */
  Core.register = function (doc) { Core.docs.push(doc); };

  // Database dari Google Sheets (lewat Apps Script). Diatur di js/data/sumber.js. Gagal = pakai data bawaan.
  function sheetKeMitra(rows) {
    var h = {}; (rows[0] || []).forEach(function (t, i) { var k = String(t || "").trim().toLowerCase().replace(/\s+/g, " "); if (k && !(k in h)) h[k] = i; });
    function K(w) { for (var i = 0; i < w.length; i++) { var x = Core.excel.col(h, [w[i]]); if (x >= 0) return x; } return -1; }
    var c = { n: K(["nama"]), e: K(["email", "e-mail"]), kc: K(["kecamatan"]), k: K(["nik"]), no: K(["no.spd", "nomor spd", "no spd", "spd"]), b: K(["mulai jalan", "berangkat"]), l: K(["jumlah ok", "lama"]) };
    if (c.n < 0) throw new Error("Kolom Nama tidak ditemukan di Google Sheets.");
    var seen = {}, out = [];
    rows.slice(1).forEach(function (r) {
      var nama = cell(r, c.n).replace(/\s+/g, " "); if (!nama) return;
      var key = normNama(nama), lama = parseInt(cell(r, c.l), 10) || 0, ent = [nama, cell(r, c.e).toLowerCase(), cell(r, c.kc), cell(r, c.k), cell(r, c.no), Core.parseDate(cell(r, c.b)), lama];
      if (seen[key]) { var e0 = seen[key]; for (var i = 1; i < 7; i++) if (!e0[i] && ent[i]) e0[i] = ent[i]; } else { seen[key] = ent; out.push(ent); }
    });
    return out;
  }
  Core.muatSumber = async function () {
    var cfg = Core.SUMBER_DATA || {}, m = cfg.mitra; if (!m || !m.url) return "";
    try {
      var ctl = new AbortController(), t = setTimeout(function () { ctl.abort(); }, 15000);
      var r = await fetch(m.url + (m.url.indexOf("?") >= 0 ? "&" : "?") + "token=" + encodeURIComponent(cfg.token || "") + "&sheet=" + encodeURIComponent(m.sheet || ""), { signal: ctl.signal });
      clearTimeout(t); var j = await r.json();
      if (!j || j.error || !j.rows) throw new Error(j && j.error ? j.error : "respons tidak dikenal");
      var list = sheetKeMitra(j.rows); if (!list.length) throw new Error("tidak ada baris mitra");
      Core.MITRA = list; Core.dataOnline = true; return "";
    } catch (e) { return "Database Google Sheets tidak bisa dimuat (" + (e && e.message ? e.message : e) + "). Memakai data bawaan."; }
  };
  Core.start = function () {
    var app = $("app"), SITE = Core.SITE || {}, sumberMsg = "";
    if (Core.SUMBER_DATA && Core.SUMBER_DATA.mitra && Core.SUMBER_DATA.mitra.url) {
      app.innerHTML = '<div class="wrap"><p class="sub">Memuat database dari Google Sheets...</p></div>';
      Core.muatSumber().then(function (msg) { sumberMsg = msg; init2(); });
    } else init2();
    function init2() { boot(); }
    function boot() {
    Core.staff.init(); Core.mitra.init(); Core.baris.init(); Core.visumMitra.init();
    var logo = $("brandLogo"); if (logo && Core.LOGO) logo.src = Core.LOGO;
    $("brandTitle").textContent = SITE.judul || "Generator Dokumen"; $("brandSub").textContent = SITE.sub || "BPS Kota Jakarta Timur";
    document.title = SITE.judul || "Generator Dokumen";
    function staffCard() {
      var info = el("span", { class: "hint", text: Core.staff.info }), reset = el("button", { type: "button", class: "ghost", text: "Pakai data bawaan", hidden: !Core.staff.custom, onclick: function () { Core.staff.useDefaults(); } });
      var file = el("input", { type: "file", accept: ".xlsx", id: "staffFile", "aria-label": "Unggah Excel data pegawai" });
      file.addEventListener("change", async function () {
        var f = file.files && file.files[0]; if (!f) return;
        try { await Core.staff.loadFile(f); } catch (e) { info.textContent = "File tidak bisa dipakai. " + (e && e.message ? e.message : ""); info.className = "hint err"; }
        file.value = "";
      });
      var off = Core.staff.onChange(function () { if (!document.body.contains(info)) { off(); return; } info.textContent = Core.staff.info; info.className = "hint"; reset.hidden = !Core.staff.custom; });
      return el("section", { class: "card stack" }, [el("h2", { text: "Data pegawai" }),
        el("p", { class: "sub", text: 'Dipakai semua dokumen. Unggah Excel dengan sheet "Referensi Pegawai": kolom A nama, B nama bergelar, C jabatan.' }),
        el("div", { class: "tpl" }, [file, reset]), info]);
    }
    function stat(v, l, o) { return el("div", { class: "stat" }, [el("b", { class: o ? "o" : "", text: String(v) }), el("span", { text: l })]); }
    function home() {
      app.innerHTML = ""; $("topBtn").hidden = false; document.title = SITE.judul || "Generator Dokumen";
      var cards = Core.docs.map(function (d) {
        return el("a", { class: "doc card", href: "#" + d.id, onclick: function (e) { e.preventDefault(); go(d.id); } }, [
          el("strong", { text: d.nama }), el("span", { class: "sub", text: d.deskripsi || "" }), el("span", { class: "open", text: "Buka" })]);
      });
      var stats = [stat(Core.docs.length, "Jenis dokumen", true)];
      if (SITE.stat === "mitra") stats.push(stat((Core.MITRA || []).length, "Mitra terdaftar")); else if (SITE.stat === "pegawai") stats.push(stat(Core.staff.list.length, "Pegawai terdaftar"));
      stats.push(stat((Core.POK || []).length, "Pembebanan POK"), stat("PDF", "Hasil unduhan"));
      var steps = [["1", "Pilih dokumen", "Buka salah satu dokumen di daftar"], ["2", "Pilih data dan isi", "Pilih nama, atur tanggal, atau unggah Excel"], ["3", "Unduh PDF atau Word", "Hasil siap cetak atau diedit, banyak dokumen jadi ZIP"]];
      var kids = [
        el("section", { class: "wrap hero fade-in" }, [el("span", { class: "pill", text: SITE.judul || "Generator Dokumen" }),
          el("h1", { text: (SITE.hero && SITE.hero[0]) || "Buat dokumen kantor dengan cepat dan rapi" }), el("p", { text: (SITE.hero && SITE.hero[1]) || "Pilih dokumen, isi data, lalu unduh PDF." }),
          ].concat(SITE.tanpaStat ? [] : [el("div", { class: "stats" }, stats)])),
        el("section", { class: "band" }, [el("div", { class: "wrap", style: "padding-block:0" }, [el("h2", { text: "Cara Kerja" }), el("p", { class: "lead", text: "Tiga langkah dari data sampai dokumen" }),
          el("div", { class: "steps" }, steps.map(function (t) { return el("div", { class: "step" }, [el("i", { text: t[0] }), el("strong", { text: t[1] }), el("span", { text: t[2] })]); }))])]),
        el("section", { class: "wrap sect", id: "docs" }, [el("h2", { text: "Pilih Dokumen" }), el("p", { class: "lead", text: "Dokumen yang tersedia" }), el("div", { class: "cards" }, cards)].concat(SITE.staff === false ? [] : [el("div", { class: "stack", style: "margin-top:24px" }, [staffCard()])]))
      ];
      if (sumberMsg) kids.unshift(el("div", { class: "wrap", style: "padding-block:12px 0" }, [el("p", { class: "hint err", text: sumberMsg })]));
      kids.forEach(function (k) { app.appendChild(k); });
    }
    function open(id) {
      var d = Core.docs.filter(function (x) { return x.id === id; })[0];
      if (!d) return home();
      app.innerHTML = ""; $("topBtn").hidden = true; document.title = d.nama + " - " + (SITE.judul || "Generator Dokumen");
      var box = el("div", {});
      var back = el("a", { class: "back", href: "#", text: "\u2190 Semua dokumen", onclick: function (e) { e.preventDefault(); go(""); } });
      app.appendChild(el("div", { class: "wrap fade-in" }, [back, el("div", { class: "pagehead" }, [el("h1", { text: d.nama }), el("p", { class: "sub", text: d.deskripsi || "" })]), box]));
      d.mount(box); window.scrollTo(0, 0);
    }
    function go(id) { try { location.hash = id || ""; } catch (e) { /* abaikan */ } id ? open(id) : home(); }
    $("brand").addEventListener("click", function (e) { e.preventDefault(); go(""); });
    $("topBtn").addEventListener("click", function (e) { e.preventDefault(); var t = $("docs"); if (t) t.scrollIntoView({ behavior: "smooth" }); });
    window.addEventListener("hashchange", function () { var id = location.hash.replace("#", ""); id && id !== "docs" ? open(id) : home(); });
    var id = location.hash.replace("#", ""); id && id !== "docs" ? open(id) : home();
    Core.pdf.init().catch(function () { /* ditangani tiap dokumen */ });
    }
  };
})();
