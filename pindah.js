function processSelectedCell() {
  try {
    // Dapatkan spreadsheet aktif
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getActiveSheet();
    const range = sheet.getActiveRange(); // Dapatkan sel yang sedang dipilih
    const sheetName = "Search"; // Nama sheet target
    const searchColumn = 1; // Kolom A (index 1)

    if (
      sheet.getName() === sheetName && // Pastikan di sheet yang benar
      range.getColumn() === searchColumn && // Pastikan di kolom A
      range.getRow() >= 8 && range.getRow() <= 30 // Pastikan dalam batasan baris
    ) {
      // Ambil nilai dari sel yang dipilih
      const rowOffset = range.getRow() - 8; // Hitung offset baris

      // Pindahkan data berdasarkan offset
      sheet.getRange("H2").setValue(sheet.getRange(8 + rowOffset, 1).getValue()); // EMP_NIP
      sheet.getRange("H3").setValue(sheet.getRange(8 + rowOffset, 2).getValue()); // EMP_Name
      sheet.getRange("H4").setValue(sheet.getRange(8 + rowOffset, 3).getValue()); // EMP_Fak
      sheet.getRange("H6").setValue(sheet.getRange(8 + rowOffset, 5).getValue()); // EMP_TMT

      // Format nilai di H4 menjadi kapital pada huruf pertama setiap kata
      formatH4(sheet);
    }
  } catch (error) {
    SpreadsheetApp.getUi().alert("Terjadi kesalahan: " + error.message);
  }
}
