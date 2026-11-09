function updateToKirimEmail() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const lastRow = sheet.getLastRow(); // Ambil baris terakhir yang berisi data
  const range = sheet.getRange(`J4:J${lastRow}`); // Batasi range hingga baris terakhir
  const values = range.getValues();

  // Loop melalui setiap nilai dalam range
  for (let i = 0; i < values.length; i++) {
    if (values[i][0].trim() === 'Penomoran') { // Periksa jika nilai adalah "Penomoran"
      sheet.getRange(`J${i + 4}`).setValue('Kirim Email'); // Ubah nilai menjadi "Kirim Email"
    }
  }
}
