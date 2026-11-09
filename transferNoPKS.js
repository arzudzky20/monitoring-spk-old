function transferNIPandNoPKS() {
  var ssA = SpreadsheetApp.getActiveSpreadsheet();
  var sheetA = ssA.getSheetByName("Database");
  var rangeNoPKS = sheetA.getRange("I4:I").getValues(); // Kolom No. PKS
  var rangeNIP = sheetA.getRange("B4:B").getValues(); // Kolom NIP
  var rangeDate = sheetA.getRange("F4:F").getValues(); // Kolom Tanggal (ambil tahunnya saja)
  var rangeStatus = sheetA.getRange("J4:J").getValues(); // Kolom Status

  var ssB = SpreadsheetApp.openById("1aXIU39ossBdDRtxMoq0535pU06N-dKADmPuNiqTTs_Q");
  var sheetB = ssB.getSheetByName("Penomoran Kontrak");
  
  var startRow = 640; // Baris awal tujuan di Spreadsheet B
  var lastRow = sheetB.getLastRow(); // Baris terakhir yang memiliki data

  // Mencari baris kosong pertama di kolom C dari baris 640 ke bawah
  var targetRow = startRow;
  var columnCValues = sheetB.getRange(startRow, 3, lastRow - startRow + 1, 1).getValues();
  
  for (var i = 0; i < columnCValues.length; i++) {
    if (!columnCValues[i][0]) { // Jika sel kosong ditemukan
      targetRow = startRow + i;
      break;
    }
  }

  var newDataNIP = [];
  var newDataNoPKS = [];
  var newDataYear = [];
  var updateStatusRows = []; // Baris yang statusnya akan diubah

  for (var i = 0; i < rangeNoPKS.length; i++) {
    if (rangeNoPKS[i][0] !== "" && rangeStatus[i][0] === "Penomoran") { 
      newDataNIP.push([rangeNIP[i][0]]);
      newDataNoPKS.push([rangeNoPKS[i][0]]);
      var date = new Date(rangeDate[i][0]);
      newDataYear.push([date.getFullYear()]);
      updateStatusRows.push(i + 4); // Simpan baris yang statusnya akan diupdate
    }
  }

  if (newDataNIP.length > 0) {
    var rangeB_NIP = sheetB.getRange(targetRow, 2, newDataNIP.length, 1);
    rangeB_NIP.setValues(newDataNIP);

    var rangeB_NoPKS = sheetB.getRange(targetRow, 3, newDataNoPKS.length, 1);
    rangeB_NoPKS.setValues(newDataNoPKS);

    var rangeB_Year = sheetB.getRange(targetRow, 4, newDataYear.length, 1);
    rangeB_Year.setValues(newDataYear);

    // Update status hanya untuk baris yang berhasil ditransfer
    for (var i = 0; i < updateStatusRows.length; i++) {
      var statusCell = sheetA.getRange(updateStatusRows[i], 10); // Kolom J
      if (statusCell.getValue() === "Penomoran") { 
        statusCell.setValue("Transfer Penomoran");
      }
    }
  }
}
