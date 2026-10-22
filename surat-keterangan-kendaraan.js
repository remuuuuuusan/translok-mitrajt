/* Surat Keterangan Tidak Menggunakan Kendaraan Dinas, satu per mitra. */
Core.simpleDoc({
  id: "sk-kendaraan", nama: "Surat Keterangan Tidak Menggunakan Kendaraan Dinas", orang: "mitra",
  deskripsi: "Pernyataan tidak memakai kendaraan dinas selama perjalanan. Tanggal surat mengikuti tanggal kembali mitra.",
  fields: Core.F(["ppk", "ppkNip", "kegiatan", "suratTugasNo", "suratTugasTgl", "jabatanMitra", "unitKerja"]),
  namaFile: function (v) { return "SK_Kendaraan_" + Core.firstWord(v.mitra.nama); },
  layout: function (v, p) {
    var m = v.mitra, X = 85.2, W = 417.4, C = 293.9, LH = 15.85;
    p.b("SURAT KETERANGAN", C, 85.8, { align: "c", font: "B" });
    p.b("Yang bertanda tangan dibawah ini :", X, 133.9);
    [["Nama", m.nama, 165.6], ["Jabatan", v.jabatanMitra, 191.5], ["Pangkat/Gol", "-", 217.3], ["Unit Kerja", v.unitKerja, 243.2]].forEach(function (r) {
      p.b(r[0], X, r[2]); p.b(":", 157.2, r[2]); p.b(r[1], 165.2, r[2]);
    });
    var y = p.justify("Menerangkan bahwa dalam rangka melaksanakan perjalanan dinas dalam rangka " + v.kegiatan + " pada tanggal " + Core.fmtDate(m.berangkat) + " s.d. " + Core.fmtDate(m.kembali) +
      " sesuai Surat Tugas Nomor " + v.suratTugasNo + " tanggal " + v.suratTugasTglTeks + " benar-benar tidak menggunakan kendaraan dinas.", X, 269.0 - 2.6, W, { lh: LH });
    y = p.justify("Demikian pernyataan ini dibuat dengan sebenar-benarnya untuk dipergunakan sebagaimana mestinya. Apabila terdapat kekeliruan dalam pertanggungjawaban SPD dan mengakibatkan kerugian negara, daftar nama terlampir bersedia dituntut sesuai peraturan yang berlaku dan pengembalian biaya transport lokal yang sudah diterima ke kas negara.",
      X, y + 9.95, W, { lh: LH });
    var sh = Math.max(0, (y - LH + 2.6) - 405.8), L = 188.1, R = 395.2;
    p.b("Mengetahui,", L, 463.4 + sh, { align: "c" }); p.b("Jakarta, " + Core.fmtDate(m.kembali), R, 463.4 + sh, { align: "c" });
    p.b("Pejabat Pembuat Komitmen", L, 479.2 + sh, { align: "c" }); p.b("Pelaksana Perjalanan Dinas", R, 479.2 + sh, { align: "c" });
    function nm(t, c, b) { p.b(t, c, b, { align: "c", font: "B" }); var w = p.width(t, "B"); p.hline(c - w / 2, c + w / 2, b - 1.2, 0.7); }
    nm(v.ppk.replace(/\.?$/, "."), L, 558.5 + sh); nm(m.nama, R, 558.5 + sh);
    p.b("NIP. " + v.ppkNip, L, 572.3 + sh, { align: "c" }); p.b("NIK." + m.nik, R, 572.3 + sh, { align: "c" });
  }
});
