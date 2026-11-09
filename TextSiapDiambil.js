function updateToTelahDiambil() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const lastRow = sheet.getLastRow(); // Ambil baris terakhir yang berisi data
  const targetValue = sheet.getRange('F1').getValue(); // Ambil nilai di F1
  const rangeJ = sheet.getRange(`J4:J${lastRow}`); // Range untuk kolom J
  const rangeE = sheet.getRange(`E4:E${lastRow}`); // Range untuk kolom E
  const valuesJ = rangeJ.getValues(); // Nilai di kolom J
  const valuesE = rangeE.getValues(); // Nilai di kolom E

  // Format tanggal ke dd/MM/yyyy
  const today = new Date();
  const formattedDate = Utilities.formatDate(today, Session.getScriptTimeZone(), 'dd/MM/yyyy');

  // Loop melalui setiap nilai dalam range
  for (let i = 0; i < valuesJ.length; i++) {
    if (
      valuesJ[i][0].trim() === 'Siap Diambil' && // Periksa jika nilai di J adalah "Siap Diambil"
      valuesE[i][0].trim() === targetValue // Periksa jika nilai di E sama dengan F1
    ) {
      sheet.getRange(`J${i + 4}`).setValue('Telah Diambil'); // Ubah nilai di J
      sheet.getRange(`K${i + 4}`).setValue(formattedDate); // Tambahkan tanggal di K
    }
  }
}