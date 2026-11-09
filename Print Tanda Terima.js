function printDataRangeAutomatically() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // Tentukan kolom dan baris awal yang diperiksa
  const startRow = 3; // Baris awal
  const startCol = 2; // Kolom B
  const numCols = 6; // Kolom B hingga G
  const lastRow = sheet.getLastRow(); // Baris terakhir di sheet

  // Ambil data dari kolom B:G
  const data = sheet.getRange(startRow, startCol, lastRow - startRow + 1, numCols).getValues();

  // Cari baris terakhir yang memiliki data di kolom B:G
  let lastRowWithData = startRow;
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    if (row.some(cell => cell !== "")) { // Cek jika ada data
      lastRowWithData = startRow + i;
    }
  }

  // Jika tidak ada data di rentang, tampilkan pesan dan hentikan
  if (lastRowWithData === startRow) {
    SpreadsheetApp.getUi().alert("Tidak ada data yang tersedia untuk dicetak.");
    return;
  }

  // Ambil ID sheet
  const sheetId = sheet.getSheetId();

  // Ambil URL Spreadsheet
  const url = SpreadsheetApp.getActiveSpreadsheet().getUrl().replace(/edit$/, '');

  // Generate URL untuk print range
  const printUrl = `${url}export?format=pdf` + 
                   `&size=A4` +
                   `&portrait=true` + 
                   `&fitw=true` + 
                   `&sheetnames=false` + 
                   `&printtitle=false` + 
                   `&pagenumbers=false` + 
                   `&gridlines=true` + 
                   `&fzr=false` + 
                   `&gid=${sheetId}` + 
                   `&r1=${startRow - 1}` + // Baris awal (zero-indexed)
                   `&c1=${startCol - 1}` + // Kolom awal (zero-indexed)
                   `&r2=${lastRowWithData}` + // Baris akhir (zero-indexed)
                   `&c2=${startCol + numCols - 1}`; // Kolom akhir (zero-indexed)

  // Buka tab baru dengan URL print
  const html = HtmlService.createHtmlOutput(`<script>window.open("${printUrl}", "_blank");google.script.host.close();</script>`)
    .setWidth(100)
    .setHeight(10);
  SpreadsheetApp.getUi().showModalDialog(html, "Mencetak...");
}
