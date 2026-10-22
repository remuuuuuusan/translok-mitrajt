/* SPD (Surat Perjalanan Dinas) halaman depan, satu per mitra. Data mitra dari Excel. */
Core.simpleDoc({
  id: "spd", nama: "SPD (Surat Perjalanan Dinas)", orang: "mitra", gabungAwal: false,
  deskripsi: "Halaman depan SPD untuk mitra. Nomor, tanggal, dan lama perjalanan diambil dari Excel mitra.",
  fields: Core.F(["ppk", "ppkNip", "maksud", "lokasi", "pok"]).concat([
    { id: "angkutan", label: "Alat angkutan", default: "Kendaraan Umum", simpan: true, grup: "Pengaturan" },
    { id: "jabatan", label: "Jabatan/instansi pelaksana", default: "Mitra BPS Kota Jakarta Timur", simpan: true, grup: "Pengaturan" },
    { id: "dikeluarkan", label: "Dikeluarkan di", default: "Jakarta", simpan: true, grup: "Pengaturan" }
  ]),
  namaFile: function (v) { return "SPD_" + Core.firstWord(v.mitra.nama); },
  layout: function (v, p) {
    var m = v.mitra, S = { size: 10 }, X0 = 74.9, X1 = 559.1;
    function T(t, x, bottom, o) { p.b(t, x, bottom, Object.assign({ size: 10 }, o)); }
    T("Nomor", 371.3, 40.5); T(":", 398.0, 40.5);
    (function(t){var out=[],cur="";t.split(/(?<=[-\/])/).forEach(function(seg){if(cur&&p.width(cur+seg,"R",10)>144){out.push(cur);cur=seg}else cur+=seg});if(cur)out.push(cur);return out})(String(m.nomor)).forEach(function (l, i) { T(l, 414.1, 40.5 + i * 11.5); });
    T("Lembar", 371.3, 67.4); T(":", 398.0, 67.4);
    T("BADAN PUSAT STATISTIK", 76.9, 82.8); T("KOTA JAKARTA TIMUR", 76.9, 94.3); T("Jl. Cipinang Baru Raya No.14", 76.9, 109.7); T("Jakarta Timur 13240", 76.9, 125.1);
    [[142.7, 0.5], [158.7, 0.75], [174.9, 0.75], [221.8, 0.75], [252.1, 0.75], [268.2, 0.75], [299.8, 0.75], [346.7, 0.75], [377.0, 0.75], [489.2, 0.75], [505.3, 0.5]].forEach(function (l) { p.hline(X0, X1, l[0], l[1]); });
    p.vline(315.2, 142.4, 505.5, 0.75); p.vline(369.3, 346.3, 377.4, 0.75);
    function row(n, b, label, val, sub, letter) {
      if (n) T(n, 76.9, b);
      if (letter) { T(letter, 96.0, b); T(label, 117.3, b); } else T(label, 96.0, b);
      if (val != null) T(val, 317.2, b);
    }
    row("1.", 156.4, "Pejabat Pembuat Komitmen", v.ppk);
    row("2.", 172.5, "Nama pegawai yang melaksanakan perjalanan dinas", m.nama);
    row("3", 188.7, "Pangkat dan golongan", "-", 0, "a."); row("", 204.1, "Jabatan/ instansi", v.jabatan, 0, "b."); row("", 219.5, "Tingkat Biaya Perjalanan Dinas", null, 0, "c.");
    row("4.", 235.6, "Maksud perjalanan dinas", v.maksud);
    row("5.", 265.9, "Alat Angkutan yang dipergunakan", v.angkutan);
    row("6.", 282.0, "Tempat keberangkatan", v.lokasi, 0, "a."); row("", 297.4, "Tempat tujuan", v.lokasi, 0, "b.");
    row("7.", 313.6, "Lamanya perjalanan Dinas", m.lama + "  hari", 0, "a."); row("", 329.0, "Tanggal Berangkat", Core.fmtDate(m.berangkat), 0, "b.");
    row("", 344.4, "Tanggal harus kembali/ tiba ditempat baru *)", Core.fmtDate(m.kembali), 0, "c.");
    T("8", 76.9, 360.5); T("Pengikut :", 96.0, 360.5); T("Nama", 198.0, 360.5); T("Umur", 317.2, 360.5); T("Hubungan keluarga/keterangan", 371.3, 360.5);
    T("9", 76.9, 390.8); T("Pembebanan anggaran", 96.0, 390.8);
    var pk = Core.pokOf(v.pok) || { kode: "", progKode: "", kegKode: "", komKode: "", progNama: "", kegNama: "", komNama: "" };
    function dua(t) { var l = p.wrap(t, 150, S); if (l.length > 2) l = [l[0], l.slice(1).join(" ")]; return [l[0] || "", l[1] || ""]; }
    [["Program", pk.progKode, dua(pk.progNama), 390.8], ["Kegiatan", pk.kegKode, dua(pk.kegNama), 417.7], ["Komponen", pk.komKode, dua(pk.komNama), 444.6]].forEach(function (r) {
      T(r[0], 245.1, r[3]); T(r[1], 317.2, r[3]); T(r[2][0], 371.3, r[3]); T(r[2][1], 371.3, r[3] + 11.5);
    });
    row("", 471.5, "Instansi", "Badan Pusat Statistik Kota Jakarta Timur", 0, "a."); row("", 486.9, "Mata anggaran", pk.kode, 0, "b.");
    T("10", 76.9, 503.0); T("Keterangan lain-lain", 96.0, 503.0);
    T("Dikeluarkan di", 399, 534.3, { align: "r" }); T(":", 403, 534.3); T(v.dikeluarkan, 411, 534.3);
    T("Pada Tanggal", 399, 549.7, { align: "r" }); T(":", 403, 549.7); T(Core.fmtDate(m.berangkat), 411, 549.7);
    T("PEJABAT PEMBUAT KOMITMEN", 437.25, 580.5, { align: "c" });
    var nm = v.ppk.replace(/\.?$/, "."), w = p.width(nm, "B", 10);
    T(nm, 437.25, 634.3, { align: "c", font: "B" }); p.hline(437.25 - w / 2, 437.25 + w / 2, 632.8, 0.5);
    T("NIP. " + v.ppkNip, 437.25, 649.7, { align: "c" });
    T("Tembusan disampaikan kepada :", 96.0, 665.1); T("1.", 96.0, 680.5); T("2.", 96.0, 695.9);
  }
});
