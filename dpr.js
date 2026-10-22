/* Daftar Pengeluaran Riil, satu per mitra. Jumlah = lama perjalanan x tarif per hari (atau kolom Biaya di Excel). */
Core.simpleDoc({
  id: "dpr", nama: "Daftar Pengeluaran Riil (DPR)", orang: "mitra",
  deskripsi: "Pernyataan biaya transport mitra yang tidak ada bukti pengeluarannya. Jumlah dihitung dari lama perjalanan.",
  fields: Core.F(["ppk", "ppkNip", "tarif", "tglBayar"]),
  namaFile: function (v) { return "DPR_" + Core.firstWord(v.mitra.nama); },
  layout: function (v, p) {
    var m = v.mitra, rp = Core.rupiah(m.biaya), pdf = Core.pdf;
    function T(t, x, bottom, o) { p.b(t, x, bottom, o); }
    p.b("DAFTAR PENGELUARAN RIIL", 296.95, 89.2, { align: "c", font: "B", size: 14 });
    T("Yang bertanda tangan di bawah ini:", 57.5, 119.2);
    [["Nama", m.nama, 135.0], ["NIK", m.nik, 150.7]].forEach(function (r) { T(r[0], 57.2, r[2]); T(":", 154.6, r[2]); T(r[1], 169.2, r[2]); });
    T("Berdasarkan Surat Perjalanan Dinas (SPD)", 57.6, 168.7);
    [["Tanggal", Core.fmtDate(m.berangkat), 183.6], ["Nomor", m.nomor, 198.0]].forEach(function (r) { T(r[0], 57.2, r[2]); T(":", 154.3, r[2]); T(r[1], 169.7, r[2]); });
    T("dengan ini kami menyatakan dengan sesungguhnya bahwa :", 57.7, 213.7);
    T("1.  Biaya transport pegawai di bawah ini yang tidak dapat diperoleh bukti-bukti pengeluaran,", 72.8, 242.6); T("meliputi:", 86.6, 256.0);
    [254.35, 273.5, 290.3, 308.7, 326.0, 342.75].forEach(function (y) { p.hline(87.0, 499.5, y, 0.7); });
    p.vline(107.2, 255.9, 308.4, 0.7); p.vline(313.4, 255.9, 325.6, 0.7); p.vline(411.2, 255.9, 325.6, 0.7);
    [["NO", 97.1], ["URAIAN", 210.3], ["JUMLAH", 362.3], ["KETERANGAN", 455.2]].forEach(function (h) { T(h[0], h[1], 273.3, { align: "c", font: "B" }); });
    [["(1)", 97.1], ["(2)", 210.3], ["(3)", 362.3], ["(4)", 455.2]].forEach(function (h) { T(h[0], h[1], 289.4, { align: "c" }); });
    T("1", 88.8, 308.3); T("Transportasi Lokal", 108.6, 308.3); T(rp, 410.3, 308.3, { align: "r" });
    T("JUMLAH", 210.3, 325.9, { align: "c", font: "B" }); T(rp, 410.3, 325.9, { align: "r", font: "B" });
    T(Core.terbilang(m.biaya), 87.1, 342.1, { font: "BI" });
    T("2.  Jumlah uang tersebut pada angka 1 di atas benar-benar dikeluarkan untuk pelaksanaan", 71.6, 372.7);
    T("perjalanan dinas dimaksud dan apabila di kemudian hari terdapat kelebihan atas pembayaran,", 86.6, 388.9);
    T("kami bersedia untuk menyetorkan kelebihan tersebut ke Kas Negara", 86.6, 402.2);
    T("Demikian pernyataan ini kami buat dengan sebenarnya, untuk dipergunakan sebagaimana mestinya.", 56.9, 441.7);
    var L = 165, R = 430;
    T("Jakarta Timur, " + v.tglBayarTeks, R, 482.3, { align: "c" }); T("Mengetahui/Menyetujui", L, 484.2, { align: "c" });
    T("Pejabat Negara/Pegawai Negeri yang", R, 496.4, { align: "c" }); T("Pejabat Pembuat Komitmen", L, 497.5, { align: "c" });
    T("melakukan perjalanan dinas", R, 510.7, { align: "c" });
    function nm(t, c) { T(t, c, 567.5, { align: "c", font: "B" }); var w = p.width(t, "B"); p.hline(c - w / 2, c + w / 2, 564.2, 1.0); }
    nm(v.ppk, L); nm(m.nama, R);
    T("NIP " + v.ppkNip, L, 581.4, { align: "c" }); T("NIK " + m.nik, R, 581.4, { align: "c" });
  }
});
