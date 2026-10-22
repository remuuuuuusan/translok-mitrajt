# Administrasi Translok Mitra

SPD, Visum, Laporan Perjalanan Dinas, Surat Keterangan, DPR, dan Kwitansi untuk mitra, memakai database mitra.

Situs statis tanpa build dan tanpa server. PDF dibuat di browser pengguna.

- `index.html` halaman utama
- `js/core.js` bagian bersama
- `js/docs/` satu file per dokumen
- `js/data/` data bawaan (logo, POK, database)
- `TEMPLATE_*.xlsx` template Excel (juga bisa diunduh dari situs)
- `PANDUAN-TAMBAH-DOKUMEN.md` cara menambah dokumen

Deploy: hubungkan repo ini ke Vercel (Framework Preset: Other, tanpa build command). Jadikan repo Private.
