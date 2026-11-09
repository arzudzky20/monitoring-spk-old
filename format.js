/**************************************
 * WEB APP ENTRY POINT (WAJIB ADA)
 **************************************/
function doGet() {
  return HtmlService.createHtmlOutputFromFile("Index")
    .setTitle("SPK Management System")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**************************************
 * SIMPAN DATA DARI SHEET "Search"
 **************************************/
function simpanData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetInput = ss.getSheetByName("Search");
  const sheetDatabase = ss.getSheetByName("Database");

  const inputValues = sheetInput.getRange("B2:H6").getValues();

  const tanggalMasukBerkas = inputValues[0][0]; // B2
  const nipNup             = inputValues[0][6]; // H2
  const namaPegawai        = inputValues[1][6]; // H3
  const namaPembawa        = inputValues[2][0]; // B4
  const jenis              = inputValues[2][6]; // H4
  const unitKerjaFakultas  = inputValues[3][6]; // H5
  const tmtw               = inputValues[4][6]; // H6

  if (
    !tanggalMasukBerkas ||
    !nipNup ||
    !namaPegawai ||
    !namaPembawa ||
    !jenis ||
    !unitKerjaFakultas ||
    !tmtw
  ) {
    SpreadsheetApp.getUi().alert("Mohon isi semua data dengan benar.");
    return;
  }

  const nextRow = sheetDatabase.getLastRow() + 1;

  sheetDatabase.getRange(nextRow, 1, 1, 10).setValues([[
    tanggalMasukBerkas,      // A
    String(nipNup),          // B (AMAN UNTUK SEARCH)
    namaPegawai,             // C
    jenis,                   // D
    unitKerjaFakultas,       // E
    tmtw,                    // F
    "",                      // G Nomor PKS
    namaPembawa,             // H
    "",                      // I File
    "Hold"                   // J Status
  ]]);

  sheetInput.getRange("B2").setValue(new Date());
  ss.toast("Data SPK berhasil disimpan ke Database.", "Sukses");
}

/**************************************
 * AMBIL DATA SPK (UNTUK WEB APP)
 **************************************/
function getSPKData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Database");
  const lastRow = sheet.getLastRow();

  if (lastRow < 4) return [];

  const data = sheet.getRange("A4:K" + lastRow).getValues();

  return data
    .map((row, index) => ({
      id: index + 4,
      tanggalMasuk: row[0],
      nip: String(row[1] || ""),
      nama: String(row[2] || ""),
      jenis: row[3],
      unitKerja: row[4],
      tmt: row[5],
      nomorSPK: String(row[8] || "-"),
      status: String(row[9] || ""),
      tanggalDiambil: row[10]
    }))
    .filter(item => item.nama);
}

/**************************************
 * SEARCH SPK (ANTI ERROR)
 **************************************/
function searchSPK(keyword) {
  const data = getSPKData();
  const search = String(keyword).toLowerCase();

  return data.filter(item =>
    String(item.nomorSPK).toLowerCase().includes(search) ||
    String(item.nama).toLowerCase().includes(search) ||
    String(item.nip).toLowerCase().includes(search)
  );
}
