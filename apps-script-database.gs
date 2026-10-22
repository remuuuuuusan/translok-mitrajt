// Apps Script untuk membagikan isi Google Sheets ke situs Generator Dokumen Mitra.
// Tempel di: Google Sheets > Ekstensi > Apps Script. Ganti TOKEN dengan kata sandi acak panjang (minimal 20 karakter).
const TOKEN = 'GANTI-DENGAN-KUNCI-RAHASIA-ANDA';

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (!p.token || p.token !== TOKEN) return keluar({ error: 'tidak diizinkan' });
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(p.sheet || 'PPL REVISIT');
  if (!sheet) return keluar({ error: 'sheet tidak ditemukan' });
  // getDisplayValues = nilai persis seperti tampil di tabel (NIK dan tanggal tidak berubah format)
  return keluar({ rows: sheet.getDataRange().getDisplayValues() });
}

function keluar(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
