/* Laporan Perjalanan Dinas + Lampiran 1 (tabel kegiatan dengan foto dokumentasi). Versi mitra: satu laporan per mitra, NIK dari database mitra. */
Core.simpleDoc({
  id: "laporan-mitra", nama: "Laporan Perjalanan Dinas Mitra", orang: "mitra", satu: true, tanpaTambahan: true,
  deskripsi: "Laporan perjalanan dinas mitra beserta lampiran kegiatan dan foto dokumentasi.",
  fields: [
    { id: "tampilNik", label: "Cantumkan NIK pelaksana di laporan", type: "checkbox", default: true },
    { id: "lokasi", label: "Lokasi tujuan perjalanan dinas", default: "Kecamatan Cakung dan Kecamatan Jatinegara" },
    { id: "tanggal", label: "Tanggal laporan", type: "date" },
    { id: "kegiatanRows", label: "Kegiatan (Lampiran 1)", type: "tabel", tambah: "Tambah kegiatan", default: [{ tanggal: "hariini", jam: "13.00-16.00" }],
      kolom: [{ id: "tanggal", label: "Tanggal", type: "date", default: "hariini" }, { id: "jam", label: "Jam", type: "text", placeholder: "13.00-16.00" },
        { id: "uraian", label: "Uraian kegiatan", type: "textarea" }, { id: "masalah", label: "Permasalahan", type: "textarea" }, { id: "solusi", label: "Solusi", type: "textarea" },
        { id: "foto", label: "Foto dokumentasi", type: "foto" }] }
  ].concat(Core.F(["pok"])).concat([
    { id: "judulKop", label: "Judul kop baris 2", default: "REVISIT SENSUS EKONOMI TAHUN 2026" },
    { id: "peran", label: "Judul kop baris 3", default: "PETUGAS PENDATA LAPANGAN" },
    { id: "kegiatan", label: "Nama kegiatan", default: "Revisit Sensus Ekonomi 2026" },
    { id: "unitKerja", label: "Unit kerja", default: "BPS Kota Jakarta Timur" },
    { id: "ketuaJabatan", label: "Jabatan yang mengetahui", default: "Ketua Tim Pelaksana Sensus Ekonomi 2026" },
    { id: "ketuaNama", label: "Nama yang mengetahui", default: "Nuraliyah, S.Si., SE., M.Si." },
    { id: "ketuaNip", label: "NIP yang mengetahui", default: "197601071996122001" }
  ].map(function (f) { return Object.assign({ simpan: true, grup: "Pengaturan" }, f); })),
  namaFile: function (v) { return "Laporan_Mitra_" + Core.firstWord(v.mitra.nama); },
  layout: function (v, p) {
    var H = { font: "H" }, C = 297.7;
    function T(t, x, bottom, size, o) { p.b(t, x, bottom, Object.assign({ font: "H", size: size }, o)); }
    ["LAPORAN PERJALANAN DINAS", v.judulKop, v.peran, "BPS KOTA JAKARTA TIMUR"].forEach(function (t, i) { T(t, C, 85.8 + i * 13.8, 12, { align: "c", font: "HB" }); });
    p.hline(71.0, 575.1, 129.6, 2);
    T("Perihal", 72.1, 165.1, 11); T(":", 108.5, 165.1, 11); T("Laporan Perjalanan Dinas", 114.5, 165.1, 11); T(v.kegiatan, 114.5, 177.8, 11);
    T("Jakarta Timur, " + v.tanggalTeks, 521.5, 165.1, 11, { align: "r" });
    T("Kepada Yang Terhormat:", 78.0, 215.7, 11); T("Kepala BPS Kota Jakarta Timur", 78.0, 228.4, 11); T("di", 78.0, 241.0, 11); T("Tempat", 89.2, 256.7, 11);
    var y = p.para("Bersama ini disampaikan laporan perjalanan dinas " + v.kegiatan + " sebagai berikut.", 78.0, 282.0 - 2.4, 445, { font: "H", size: 11, lh: 12.6 });
    function item(n, label, b, val) { T(n, 85.6, b, 11); T(label, 103.6, b, 11); if (val != null) { T(":", 252.1, b, 11); T(val, 259.0, b, 11); } }
    var nik = !!v.tampilNik, dy = nik ? 0 : -19.0, n = 0;
    function no() { n++; return n + "."; }
    item(no(), "Nama", 313.3, v.mitra.nama);
    if (nik) item(no(), "NIK", 332.3, v.mitra.nik || "");
    item(no(), "Unit Kerja", 351.4 + dy, v.unitKerja); item(no(), "Lokasi Tujuan Perjadin", 370.0 + dy, v.lokasi);
    item(no(), "Pembiayaan Kegiatan", 388.7 + dy);
    (function (k) { return [["Program", k.program], ["Kegiatan", k.kegiatan], ["KRO", k.kro], ["Komponen", k.komponen], ["Akun", k.akun]]; })(Core.pokOf(v.pok) || {}).forEach(function (r, i) {
      var b = 407.3 + dy + i * 18.66; T(r[0], 108.1, b, 11); T(":", 187.0, b, 11); T(r[1], 193.5, b, 11);
    });
    item(no(), "Uraian kegiatan perjalanan dinas dan dokumentasi dapat dilihat pada Lampiran 1.", 519.2 + dy - 18.66);
    T("Demikian laporan yang dapat disampaikan untuk dijadikan sebagai bahan evaluasi.", 78.0, 550.5, 11);
    var L = 162.1, R = 380.3;
    T("Mengetahui,", L, 587.4, 10, { align: "c" }); T("Pelaksana Perjalanan", R, 587.4, 10, { align: "c" });
    T(v.ketuaJabatan, L, 599.8, 10, { align: "c" });
    T(v.ketuaNama, L, 661.8, 10, { align: "c" }); T("NIP. " + v.ketuaNip, L, 674.2, 10, { align: "c" });
    T(v.mitra.nama, R, 661.8, 10, { align: "c" }); if (nik) T("NIK. " + (v.mitra.nik || ""), R, 674.2, 10, { align: "c" });

    // ---- Lampiran 1 ----
    var X = [58.4, 119.6, 163.8, 245.6, 342.8, 426.1, 570.2], CEN = X.slice(0, 6).map(function (x, i) { return (x + X[i + 1]) / 2; }), LIM = 790, rows = v.kegiatanRows;
    var first = true;
    function header() {
      p.newPage(); T("Lampiran 1", 523.4, 84.6, 11, { align: "r" });
      p.rect(X[0], 141.8, X[6] - X[0], 35.5, { fill: 0.851 }); p.rect(X[0], 177.3, X[6] - X[0], 14.5, { fill: 0.851 });
      p.hline(X[0], X[6], 142.1, 0.5); p.hline(X[0], X[6], 177.6, 0.5); p.hline(X[0], X[6], 192.1, 0.5);
      T("Hari/", CEN[0], 159.8, 10, { align: "c" }); T("Tanggal", CEN[0], 171.3, 10, { align: "c" }); T("Jam", CEN[1], 165.5, 10, { align: "c" });
      T("Uraian", CEN[2], 159.8, 10, { align: "c" }); T("Kegiatan", CEN[2], 171.3, 10, { align: "c" });
      T("Permasalahan", CEN[3], 165.5, 10, { align: "c" }); T("Solusi", CEN[4], 165.5, 10, { align: "c" }); T("Dokumentasi", CEN[5], 165.5, 10, { align: "c" });
      for (var i = 0; i < 6; i++) T("(" + (i + 1) + ")", CEN[i], 189.9, 9, { align: "c" });
      return 192.1;
    }
    function close(bottom) { X.forEach(function (x) { p.vline(x, 141.8, bottom, 0.5); }); }
    var top = header();
    rows.forEach(function (r) {
      var cols = [Core.hari(r.tanggal) ? Core.hari(r.tanggal) + ", " + r.tanggalTeks : "", r.jam, r.uraian, r.masalah, r.solusi], lines = cols.map(function (c, i) { return p.wrap(c, X[i + 1] - X[i] - 8, { font: "H", size: 10 }); });
      var n = Math.max.apply(null, lines.map(function (l) { return l.length; })), h = Math.max(70.8, n * 11.5 + 12, 74.7 + 14);
      if (top + h > LIM) { close(top); top = header(); }
      lines.forEach(function (ls, i) { var y0 = top + (h - ls.length * 11.5) / 2 + 11.5 - 2; ls.forEach(function (l, k) { T(l, CEN[i], y0 + k * 11.5 + 2, 10, { align: "c" }); }); });
      if (r.foto) p.photo(r.foto, X[5] + 5.8, top + 7, X[6] - X[5] - 11.6, h - 14);
      top += h; p.hline(X[0], X[6], top, 0.5);
    });
    close(top);
  }
});
