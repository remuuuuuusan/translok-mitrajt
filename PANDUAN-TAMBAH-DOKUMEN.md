# Menambah dokumen baru

Setiap dokumen adalah satu file di `js/docs/`. Dokumen lain tidak ikut berubah.

## Langkah

1. Salin `js/docs/_contoh-dokumen.js` menjadi `js/docs/nama-dokumen-anda.js`.
2. Ubah isinya (lihat bagian "Isi file" di bawah).
3. Buka `index.html`, tambahkan satu baris di bawah baris `surat-tugas.js`:
   `<script src="js/docs/nama-dokumen-anda.js"></script>`
4. Unggah file yang berubah ke GitHub (Add file > Upload files). Vercel memperbarui situs otomatis.

## Isi file

- `id`: huruf kecil tanpa spasi, jadi alamat halaman (`#id`).
- `nama`, `deskripsi`: tampil di kartu halaman depan.
- `pegawai: true`: pengguna memilih satu atau banyak pegawai, satu PDF per pegawai.
- `nomor`: awalan, akhiran, dan nomor awal. Hapus baris ini kalau dokumen tanpa nomor.
- `fields`: isian formulir. Tipe: `text`, `textarea`, `date`, `select` (tambah `other: true` untuk pilihan "Lainnya"), `checkbox`. Tambah `opsional: true` bila boleh kosong.
- `layout(v, p)`: posisi teks di kertas. `v` berisi nilai isian, `p` adalah alat gambar.

## Alat gambar `p`

Satuan pt. Kertas A4 = 595 x 842. `y` = garis dasar teks, diukur dari tepi atas kertas.

- `p.kop()` kop surat BPS (logo dan dua baris judul)
- `p.text("teks", x, y, {font:"R"|"B"|"BI", size:12})`
- `p.center("teks", y)`, `p.right("teks", x, y)`
- `p.para("teks panjang", x, y, lebar)` paragraf otomatis turun baris, mengembalikan y berikutnya
- `p.lines([...], x, y)` beberapa baris berurutan
- `p.line(x1, x2, y, tebal)` garis
- `p.slot(x, y, lebar, tinggi)` kotak kosong untuk tanda tangan elektronik (hanya di pratinjau)

## Cara paling mudah

Kirim ke Claude: contoh PDF atau Word dokumen baru, contoh Excel isiannya (beberapa baris),
dan kalimat ini:

> Tambahkan dokumen baru ke proyek Generator Dokumen sebagai satu file di js/docs/,
> mengikuti PANDUAN-TAMBAH-DOKUMEN.md dan contoh js/docs/surat-tugas.js.
> Cocokkan posisi teks dengan contoh PDF dan beri tahu baris yang harus saya tambahkan di index.html.

Lampirkan juga `PANDUAN-TAMBAH-DOKUMEN.md`, `js/core.js`, dan `js/docs/surat-tugas.js`.

## Batasan

Satu halaman A4 per dokumen. Kop memakai logo BPS. Tanda tangan elektronik tidak dibuat
oleh situs ini dan ditambahkan setelah PDF diunduh.

---

## Tambahan: fitur untuk dokumen yang lebih rumit

**Sumber data (`orang`)** — menentukan daftar apa yang dipilih pengguna:
- `"pegawai"` : daftar pegawai bawaan (`js/data/pegawai.js`).
- `"mitra"`   : Excel mitra (kolom: Nama, NIK, Nomor SPD, Tanggal Berangkat, Lama (hari)). Contoh: `DATA_MITRA_REVISIT.xlsx`. Kolom NIK harus berformat **Teks** di Excel.
- `"baris"`   : Excel Surat Tugas (Nomor, Tanggal, Tugas, Tujuan, Asal). Dipakai Surat Tugas dan Visum.

**Pengaturan bersama** — `Core.F(["ppk","bendahara","tarif",...])` memakai ulang isian yang sama di banyak dokumen; nilainya diingat di peramban (`simpan:true`, grup "Pengaturan"). Daftar isian ada di `js/data/perjadin.js`.

**Opsi modul**: `satu:true` (satu dokumen saja, tanpa daftar pilihan), `gabungAwal:true` (kotak "gabungkan jadi satu PDF" tercentang), `salinan:"lembar"` (jumlah salinan per orang).

**Perintah tata letak tambahan** pada `p` di dalam `layout`: `p.size(w,h)` ukuran halaman (mis. F4 595.35×935.55), `p.newPage()`, `p.rect(...)`, `p.photo(...)`, `p.justify(...)` rata kiri-kanan, `p.b(teks,x,bottom,{size})` menaruh teks berdasarkan nilai "bottom" dari pdfplumber.

Tips menyalin dokumen baru: ekspor PDF contoh, ambil koordinat teks dengan `pdfplumber` (`extract_words`), lalu tempel ke `p.b(...)`.

## Data bawaan yang ditanam di situs
- `js/data/pok.js` : daftar POK (akun 524113) dari Rincian Kertas Kerja Satker. Dipilih lewat isian `pok` (`Core.F(["pok"])`); di layout ambil dengan `Core.pokOf(v.pok)` → `{kode, program, kegiatan, kro, komponen, subKomponen, akun, progKode, kegKode, komKode, progNama, kegNama, komNama}`. Untuk memperbarui POK, ganti isi file ini.
- `js/data/mitra.js` : daftar mitra `[nama, email, kecamatan]` untuk mengisi email otomatis (tanpa NIK). Email muncul sebagai tautan "kirim email" di daftar mitra. Edit file ini untuk menambah mitra/email.
- Excel mitra boleh memakai format PPL (No.SPD, Tanggal Mulai Jalan, Jumlah OK, Email, NIK).

## Format `js/data/mitra.js` (database mitra)
Tiap baris: `[nama, email, kecamatan, NIK, no.SPD, tanggal berangkat YYYY-MM-DD, lama hari]`. Excel yang diunggah hanya menimpa/menentukan tanggal dan lama untuk mitra kegiatan itu; NIK dan nomor SPD diambil dari database. File ini berisi data pribadi, jadi situs jangan dibuka untuk umum.
