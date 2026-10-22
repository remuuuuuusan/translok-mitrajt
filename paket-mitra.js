/* Paket dokumen mitra: SPD + Surat Keterangan + DPR + Kwitansi sekaligus, satu PDF per mitra.
   Tata letak dan isian diambil dari modul masing-masing (spd.js, surat-keterangan-kendaraan.js, dpr.js, kwitansi.js). */
(function () {
  var PARTS = [["spd", "ikutSpd", "SPD (halaman depan)"], ["sk-kendaraan", "ikutSk", "Surat Keterangan Tidak Menggunakan Kendaraan Dinas"], ["dpr", "ikutDpr", "Daftar Pengeluaran Riil (DPR)"], ["kwitansi", "ikutKw", "Kwitansi"]];
  var fields = PARTS.map(function (p) { return { id: p[1], label: "Sertakan " + p[2], type: "checkbox", default: true }; }), seen = {};
  PARTS.forEach(function (p) { ((Core.specs[p[0]] || {}).fields || []).forEach(function (f) { if (!seen[f.id]) { seen[f.id] = 1; fields.push(f); } }); });
  Core.simpleDoc({
    id: "paket-mitra", nama: "Paket Dokumen Mitra (SPD, Surat Keterangan, DPR, Kwitansi)", orang: "mitra", gabungAwal: false,
    deskripsi: "Membuat SPD, Surat Keterangan, DPR, dan Kwitansi sekaligus. Satu PDF per mitra, urutan: SPD, Surat Keterangan, DPR, Kwitansi.",
    fields: fields,
    cek: function (v) { return PARTS.some(function (p) { return v[p[1]]; }) ? "" : "Pilih minimal satu dokumen untuk disertakan."; },
    namaFile: function (v) { return "Paket_" + Core.firstWord(v.mitra.nama); },
    layout: function (v, p) {
      var first = true;
      PARTS.forEach(function (x) {
        if (!v[x[1]]) return;
        if (!first) p.newPage(); first = false;
        Core.specs[x[0]].layout(v, p);
      });
    }
  });
})();
