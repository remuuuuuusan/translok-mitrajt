# Menghubungkan database mitra ke Google Sheets (tetap privat)

Hasilnya: data mitra dibaca dari Google Sheets setiap situs dibuka. Mengubah data cukup di tabel, tanpa menyentuh kode. Bila Sheets tidak bisa dibuka, situs memakai data bawaan di `js/data/mitra.js`.

## A. Siapkan Sheet
1. Buka file di Google Drive. Bila tampil tulisan ".XLSX", pilih **File > Simpan sebagai Google Sheets**, lalu pakai Sheet yang baru itu.
2. Klik **Bagikan** dan ubah akses umum menjadi **Dibatasi** (Restricted). Cukup akun Anda dan orang yang memang perlu.
3. Kolom NIK harus berformat **Teks** (Format > Angka > Teks biasa) agar 16 digit tidak berubah.
4. Judul kolom di baris pertama sheet "PPL REVISIT" dibaca: Nama, Email, Kecamatan, NIK, No.SPD, Tanggal Mulai Jalan, Jumlah OK. Boleh ada kolom lain.

## B. Pasang Apps Script
1. Di Sheet: **Ekstensi > Apps Script**.
2. Hapus isi editor, tempel seluruh isi file `apps-script-database.gs`.
3. Ganti `GANTI-DENGAN-KUNCI-RAHASIA-ANDA` dengan kata sandi acak panjang. Simpan kata sandi itu.
4. Klik **Terapkan (Deploy) > Penerapan baru > Aplikasi web**. Isi: *Jalankan sebagai* = **Saya**, *Yang memiliki akses* = **Siapa saja**. Klik Terapkan, setujui izin yang diminta.
5. Salin **URL aplikasi web** (berakhiran `/exec`).

## C. Hubungkan ke situs
1. Di GitHub, buka `js/data/sumber.js`, klik ikon pensil.
2. Isi `token` dengan kata sandi tadi, dan `url` dengan URL aplikasi web.
3. Commit changes. Situs diperbarui otomatis oleh Vercel.

## Catatan keamanan
- Sheet tetap privat. Data hanya keluar bila permintaan membawa kunci yang benar.
- Kunci ikut tertanam di kode situs, jadi orang yang bisa membuka situs bisa menemukannya. Karena itu situs dan repo GitHub sebaiknya tetap privat/diberi password.
- Bila kunci bocor: ganti TOKEN di Apps Script, terapkan ulang (Deploy > Kelola penerapan > Edit > versi baru), lalu ganti di `sumber.js`.
- Setiap perubahan kode Apps Script perlu versi penerapan baru agar berlaku.
