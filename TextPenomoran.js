function changeText() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const lastRow = sheet.getLastRow(); // Get the last row with data
  const range = sheet.getRange(`J4:J${lastRow}`); // Limit range to non-empty rows
  const values = range.getValues();

  for (let i = 0; i < values.length; i++) {
    if (values[i][0].trim() === 'TTD Direktur') { // Trim to remove extra spaces
      sheet.getRange(`J${i + 4}`).setValue('Penomoran'); // Update value
    }
  }
}