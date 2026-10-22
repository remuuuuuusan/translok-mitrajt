/* Visum SPD mitra: satu baris per hari perjalanan, ditutup blok pemeriksaan PPK. Tanggal dari data mitra. */
Core.simpleDoc({
  id: "visum-mitra", nama: "Visum Mitra (SPD halaman belakang)", orang: "visumMitra", salinan: "lembar", gabungAwal: true,
  deskripsi: "Halaman belakang SPD mitra: berangkat dan tiba tiap hari selama lama perjalanan, lalu pemeriksaan PPK. Tanpa database mitra: unggah Excel atau isi manual.",
  fields: [
    { id: "tglBerangkat", label: "Tanggal berangkat (bila tidak memilih data di atas)", type: "date", opsional: true },
    { id: "lamaMitra", label: "Lama perjalanan, hari (bila tidak memilih data di atas)", default: "1", opsional: true },
    { id: "lembar", label: "Jumlah lembar (bila tidak memilih data di atas)", default: "1", opsional: true }
  ].concat(Core.F(["lokasi", "kepala", "kepalaNip", "ppk", "ppkNip"])),
  cek: function (v) { return !v.vm && (!v.tglBerangkat || !(parseInt(v.lamaMitra, 10) > 0)) ? "Pilih data perjalanan, atau isi tanggal berangkat dan lama." : ""; },
  namaFile: function (v, i) { return "Visum_Mitra" + (v.vm && v.vm.nama ? "_" + Core.firstWord(v.vm.nama) : ""); },
  layout: function (v, p) {
    var vm = v.vm, m = vm || { berangkat: v.tglBerangkat, lama: parseInt(v.lamaMitra, 10) || 1 }, asal = (vm && vm.asal) || v.lokasi, tujuan = (vm && vm.tujuan) || v.lokasi, S = { size: 10 };
    function T(t, x, bottom, o) { p.b(t, x, bottom, Object.assign({ size: 10 }, o)); }
    function ul(t, c, bottom, th) { T(t, c, bottom, { align: "c" }); var w = p.width(t, "R", 10); p.hline(c - w / 2, c + w / 2, bottom - 1.5, 0.5); }
    var days = []; for (var i = 0; i < Math.min(m.lama, 90); i++) days.push(Core.fmtDate(Core.addDays(m.berangkat, i)));
    var first = days[0], last = days[days.length - 1], X0 = 39.8, X1 = 562.8, DV = 320.0, ROW = 145.0;
    // kepala + berangkat
    var y = 83.7;
    T("Berangkat dari", 322.1, y + 13.5); T(":", 414.4, y + 13.5); T(asal, 426.6, y + 13.5); T("(Tempat Kedudukan)", 426.6, y + 28.9);
    T("Pada Tanggal", 322.1, y + 44.3); T(":", 414.4, y + 44.3); T(first, 426.6, y + 44.3);
    T("Ke", 322.1, y + 59.7); T(":", 414.4, y + 59.7); T(tujuan, 426.6, y + 59.7);
    T("Kepala Badan Pusat Statistik", 479.4, 170.3, { align: "c" }); T("Kota Jakarta Timur", 479.4, 181.8, { align: "c" });
    ul(v.kepala, 494.4, 239.3); T("NIP. " + v.kepalaNip, 494.4, 250.8, { align: "c" });
    p.vline(DV, 83.7, 253.1, 0.75); y = 253.1; p.hline(X0, X1, y, 0.75);
    function kolon(l, x, b, val) { T(l, x, b); T(":", x === 41.8 ? 138.3 : 414.4, b); T(val, x === 41.8 ? 161.3 : 426.6, b); }
    function top() { p.newPage(); y = 27.3; p.hline(X0, X1, y, 0.5); }
    days.forEach(function (d) {
      if (y + ROW > 800) top();
      kolon("Tiba", 41.8, y + 13.8, tujuan); kolon("Pada tanggal", 41.8, y + 29.2, d);
      kolon("Berangkat dari", 322.1, y + 13.8, asal); kolon("Ke", 322.1, y + 29.2, tujuan); kolon("Pada tanggal", 322.1, y + 44.6, d);
      p.vline(DV, y, y + ROW, 0.75); y += ROW; p.hline(X0, X1, y, 0.5);
    });
    if (y + 318 > 830) top();
    var f = y;
    kolon("Tiba", 41.8, f + 13.8, tujuan); T("(Tempat Kedudukan)", 161.3, f + 25.3); kolon("Pada tanggal", 41.8, f + 40.9, last);
    ["Telah diperiksa dengan keterangan bahwa perjalanan", "tersebut atas perintahnya dan semata-mata untuk"].forEach(function (l, i) { T(l, 322.1, f + 13.8 + i * 27.2); });
    ["kepentingan jabatan dalam waktu yang sesingkat", "singkatnya"].forEach(function (l, i) { T(l, 322.1, f + 68.0 + i * 11.5); });
    T("Pejabat Pembuat Komitmen", 178.5, f + 122.4, { align: "c" }); T("Pejabat Pembuat Komitmen", 449.7, f + 122.4, { align: "c" });
    [179.5, 452.6].forEach(function (c) { var n = v.ppk.replace(/\.?$/, "."), w = p.width(n, "B", 10); T(n, c, f + 187.9, { align: "c", font: "B" }); p.hline(c - w / 2, c + w / 2, f + 186.4, 0.5); T("NIP. " + v.ppkNip, c - 1.5, f + 214.8, { align: "c" }); });
    p.vline(DV, f, f + 174.3, 0.75); p.hline(X0, X1, f + 236.5, 0.75);
    var e = f + 236.5;
    T("CATATAN LAIN - LAIN :", 41.8, e + 13.8); T("PERHATIAN", 41.8, e + 29.2); T(":", 138.3, e + 29.2);
    ["Pejabat yang berwenang menerbitkan SPD pegawai yang melakukan perjalanan dinas para pejabat", "yang mengesahkan tanggal berangkat/tiba serta bendaharawan bertanggung jawab berdasarkan",
      "peraturan-peraturan keuangan Negara apabila Negara menderita rugi akibat kesalahan, kelalaian", "dan kealpaannya."].forEach(function (l, i) { T(l, i ? 161.3 : 161.3, e + 29.2 + i * 17.2); });
  }
});
