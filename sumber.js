/* Sumber database online (Google Sheets lewat Apps Script). Kosongkan url untuk memakai data bawaan di js/data/mitra.js.
   Langkah pengaturan ada di PANDUAN-GOOGLE-SHEETS.md. */
Core.SUMBER_DATA = {
  token: "",   // kunci rahasia yang sama dengan TOKEN di Apps Script
  mitra: { url: "", sheet: "PPL REVISIT" }   // url = alamat Web app Apps Script (berakhiran /exec)
};
