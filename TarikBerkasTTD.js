function pindahkanData() {
  const SHEET_SUMBER = "Database";
  const SHEET_TUJUAN = "Notifikasi Direktur";
  const KOLOM_START = 3; // Kolom C
  const BARIS_START = 8; // Baris 8 di sheet tujuan

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sumber = ss.getSheetByName(SHEET_SUMBER);
  const tujuan = ss.getSheetByName(SHEET_TUJUAN);

  const data = sumber.getRange(4, KOLOM_START, sumber.getLastRow() - 3, 8).getValues();
  const hasil = data.filter(row => row[7] === "Hold")
                    .map(row => row.slice(0, 4));

  if (hasil.length > 0) {
    tujuan.getRange(BARIS_START, KOLOM_START, hasil.length, hasil[0].length).setValues(hasil);
    SpreadsheetApp.getUi().alert("Data berhasil dipindahkan ke sheet 'Notifikasi Direktur'!");
  } else {
    SpreadsheetApp.getUi().alert("Tidak ada data yang memenuhi kriteria.");
  }
}