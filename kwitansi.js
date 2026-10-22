/* Kwitansi transport lokal, satu per mitra. */
Core.simpleDoc({
  id: "kwitansi", nama: "Kwitansi Perjalanan Dinas", orang: "mitra",
  deskripsi: "Kwitansi pembayaran transport lokal mitra dengan terbilang otomatis.",
  fields: Core.F(["ppk", "ppkNip", "bendahara", "bendaharaNip", "tarif", "tglBayar", "lokasi"]),
  namaFile: function (v) { return "Kwitansi_" + Core.firstWord(v.mitra.nama); },
  layout: function (v, p) {
    var m = v.mitra, rp = Core.rupiah(m.biaya);
    function T(t, x, bottom, o) { p.b(t, x, bottom, o); }
    T("BADAN PUSAT STATISTIK", 137.0, 74.0, { align: "c" }); T("BPS KOTA JAKARTA TIMUR", 137.0, 91.2, { align: "c" });
    T("KWITANSI", 297.3, 130.6, { align: "c", font: "B", size: 18 }); p.hline(250.6, 343.8, 127.4, 1.68);
    function row(label, b, val) { T(label, 57.4, b); T(":", 202.0, b); if (val != null) T(val, 216.5, b); }
    row("Sudah terima dari", 203.4, "Kuasa Pengguna Anggaran BPS KOTA JAKARTA TIMUR");
    row("Uang sebesar", 219.1, rp);
    row("Untuk pembayaran", 234.8, "Perjalanan Dinas Dalam Kota");
    row("Berdasarkan SPD", 250.6, "Nomor"); T(":", 264.6, 250.6); T(m.nomor, 279.1, 250.6);
    T("Tanggal", 216.5, 266.3); T(":", 264.6, 266.3); T(Core.fmtDate(m.berangkat), 279.1, 266.3);
    row("Untuk perjalanan dinas", 282.1, v.lokasi.replace(/^Kota\s+/i, "")); T("dalam kota pada", 57.4, 296.3);
    row("Terbilang", 312.1, Core.terbilang(m.biaya));
    T("Bendahara Pengeluaran", 64.2, 369.8); T("BPS", 64.2, 384.0); T("Lunas pada", 64.2, 426.4); T("tanggal,", 64.2, 440.6);
    T("Setuju dibayar", 243.25, 369.8, { align: "c" }); T("Pejabat Pembuat Komitmen", 296.85, 426.4, { align: "c" });
    T("Yang menerima,", 441.1, 369.8, { align: "c" });
    p.wrap("Jakarta Timur, " + v.tglBayarTeks, 90).forEach(function (l, i) { T(l, 410.3, 398.3 + i * 14.1); });
    function nm(t, c) { T(t, c, 497.3, { align: "c", font: "B" }); var w = p.width(t, "B"); p.hline(c - w / 2, c + w / 2, 495.2, 1.2); }
    nm(v.bendahara, 128.2); nm(v.ppk, 297.0); nm(m.nama, 444.0);
    T("Nip." + v.bendaharaNip, 128.2, 511.9, { align: "c" }); T("Nip." + v.ppkNip, 297.0, 511.9, { align: "c" }); T("Nip." + m.nik, 444.0, 511.9, { align: "c" });
  }
});
