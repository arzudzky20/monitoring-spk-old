function formatH4() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Search"); // Akses sheet "Search"
  var cell = sheet.getRange("H4"); // Akses sel H4
  var value = cell.getValue(); // Ambil nilai dari sel H4

  if (value) { // Pastikan sel tidak kosong
    var formattedValue = value
      .toString() // Konversi nilai ke string
      .toLowerCase() // Ubah seluruh teks ke huruf kecil
      .replace(/\b\w/g, function(char) { // Huruf pertama tiap kata menjadi kapital
        return char.toUpperCase();
      });
    cell.setValue(formattedValue); // Simpan kembali nilai yang sudah diformat
  }
}
