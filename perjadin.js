/* Isian yang dipakai bersama oleh dokumen perjalanan dinas (SPD, SK, DPR, Kwitansi, Laporan, Visum).
   Nilai disimpan di browser: ubah sekali di satu dokumen, dokumen lain ikut berubah.
   Pemakaian di file dokumen:  fields: Core.F(["ppk", "ppkNip"]).concat([ ...isian lain... ]) */
Core.FS = {
  ppk: { label: "Nama Pejabat Pembuat Komitmen", default: "Ragil Hermanto, SST" },
  ppkNip: { label: "NIP Pejabat Pembuat Komitmen", default: "199406212017011001" },
  bendahara: { label: "Nama Bendahara Pengeluaran", default: "Ade Yotifali, A.Md." },
  bendaharaNip: { label: "NIP Bendahara Pengeluaran", default: "198107062011011011" },
  kepala: { label: "Nama Kepala BPS", default: "Widiastuti,SST, MM" },
  kepalaNip: { label: "NIP Kepala BPS", default: "197702061997122001" },
  tarif: { label: "Tarif transport per hari (Rp)", default: "170000" },
  kegiatan: { label: "Nama kegiatan", default: "Revisit Pendataan Lengkap Sensus Ekonomi 2026" },
  maksud: { label: "Maksud perjalanan dinas", default: "Transport Pendataan Lengkap" },
  lokasi: { label: "Tempat keberangkatan dan tujuan", default: "Kota Jakarta Timur" },
  mataAnggaran: { label: "Mata anggaran", default: "GG.2902.BMA.006.005.A.524113" },
  pok: { label: "Pembebanan anggaran (POK, akun 524113)", type: "select", options: (Core.POK || []).map(function (p) { return p.label; }),
    default: ((Core.POK || []).filter(function (p) { return p.kode === "GG.2902.BMA.006.005.A.524113"; })[0] || { label: "" }).label },
  spdSufiks: { label: "Akhiran Nomor SPD (setelah nomor urut)", default: "539184-92800/TRANSLOK-2902/09/2026" },
  suratTugasNo: { label: "Nomor Surat Tugas", default: "B-3578.105/3172/SS.330/2026" },
  suratTugasTgl: { label: "Tanggal Surat Tugas", type: "date", default: "2026-09-17" },
  tglBayar: { label: "Tanggal dokumen pembayaran", type: "date", default: "2026-09-30" },
  jabatanMitra: { label: "Jabatan mitra", default: "Mitra BPS Kota Jakarta Timur" },
  unitKerja: { label: "Unit kerja", default: "BPS Kota Jakarta Timur" }
};
Core.F = function (ids) { return ids.map(function (i) { return Object.assign({ id: i, simpan: true, grup: "Pengaturan" }, Core.FS[i]); }); };
