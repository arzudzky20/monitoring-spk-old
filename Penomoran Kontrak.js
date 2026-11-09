const SHEET_NAME = 'Penomoran Kontrak'; // <--- Ganti jika nama sheet Anda berbeda

/**
 * Adds a custom menu to the spreadsheet when it's opened.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Utility Kontrak') // Menu name in Indonesian
      .addItem('Rekam Kontrak PTT', 'openContractModal') // Menu item in Indonesian
      .addToUi();
}

/**
 * Opens the HTML modal dialog for processing contracts.
 */
function openContractModal() {
  // Adjusted size for more fields
  const htmlOutput = HtmlService.createHtmlOutputFromFile('Modal')
      .setWidth(450)
      .setHeight(500); // Increased height again for new field
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Form Rekam Kontrak PTT'); // Dialog title in Indonesian
}

/**
 * Calculates the next sequential PKS number for a specific category
 * within the current year from the Google Sheet.
 * @param {string} kategori - The employee category.
 * @returns {number} The next sequence number for that category in the current year.
 */
function getNextPksSequence(kategori) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
      // If sheet doesn't exist, assume this is the very first entry ever
      // although onOpen should ensure sheet exists with headers.
      // For robustness, returning 1 here if sheet not found.
      Logger.log(`Sheet "${SHEET_NAME}" not found in getNextPksSequence. Assuming first entry.`);
      return 1;
  }

  // Get values from 'Nomor Urut PKS' (Col A), 'Kategori Pegawai' (Col G), and 'Tanggal Rekam' (Col J)
  // Assuming header is in row 1, start reading from row 2
  // Pastikan range sesuai dengan kolom di sheet Anda: A=0, G=6, J=9
  const range = sheet.getRange('A2:J' + sheet.getLastRow());
  const values = range.getValues();

  const currentYear = new Date().getFullYear();
  let maxSequenceForCategoryInCurrentYear = 0;

  if (values.length > 0) {
     // Filter rows by category AND by the current year
     const categoryRowsInCurrentYear = values.filter(row => {
         const rowKategori = row[6]; // Kolom G (Kategori Pegawai)
         const rowDate = row[9];     // Kolom J (Tanggal Rekam)

         // Ensure rowKategori matches input kategori
         const categoryMatch = rowKategori === kategori;

         // Ensure rowDate is a valid Date object and its year matches current year
         let yearMatch = false;
         if (rowDate instanceof Date && !isNaN(rowDate)) {
             yearMatch = rowDate.getFullYear() === currentYear;
         }

         return categoryMatch && yearMatch;
     });

     if (categoryRowsInCurrentYear.length > 0) {
         const sequenceNumbers = categoryRowsInCurrentYear
             .map(row => row[0]) // Kolom A (Nomor Urut PKS)
             .filter(seq => typeof seq === 'number' && seq > 0); // Filter out non-numbers or <= 0

         if (sequenceNumbers.length > 0) {
             maxSequenceForCategoryInCurrentYear = Math.max(...sequenceNumbers);
         }
     }
  }

  return maxSequenceForCategoryInCurrentYear + 1;
}

/**
 * Calculates the *predicted* next PKS number and formats the contract number
 * string for preview in the modal. Does NOT interact with the sheet for incrementing.
 * @param {string} kategori - The employee category.
 * @param {string} nama - The employee name.
 * @returns {string} The formatted predicted contract number string.
 */
function getPredictedNomorKontrak(kategori, nama) {
  if (!kategori || !nama) {
    return "Isi Nama & Kategori untuk Preview"; // Placeholder if data is missing
  }

  try {
      // Get the *next predicted* sequence number for this category in the current year
      const predictedSequence = getNextPksSequence(kategori); // Reuse the logic, but this call is just for prediction

      // Format the sequence number with leading zeros
      const formattedSequence = Utilities.formatString('%03d', predictedSequence);

      // Clean up name for use in the contract number string
      const cleanedNama = nama.trim().replace(/[^a-zA-Z0-9\s]/g, '').replace(/\s+/g, '_').toUpperCase();

      // Generate the predicted contract number string
      return `PKS${formattedSequence}/${cleanedNama} (${kategori})`;

  } catch (e) {
      Logger.log("Error generating predicted contract number: " + e);
      return "Error Preview"; // Indicate error in preview
  }
}


/**
 * Processes the contract PDF: receives Base64 data, decodes it,
 * creates the file, logs data to the sheet, and generates the contract number.
 * @param {Object} data - An object containing nama, nipNup, unitKerja, jenis, kategori, fileData, fileName.
 * @returns {Object} An object with status (success/error) and message.
 */
function processContract(data) {
  try {
    const scriptProperties = PropertiesService.getScriptProperties();
    const destinationFolderId = scriptProperties.getProperty('DESTINATION_FOLDER_ID');

    if (!destinationFolderId) {
      return { status: 'error', message: 'ID folder tujuan Google Drive belum diatur. Mohon konfigurasi properti skrip dengan menjalankan fungsi setFolderProperties.' };
    }

    // --- Input Validation ---
    if (!data.nama || !data.nipNup || !data.unitKerja || !data.jenis || !data.kategori || !data.fileData) {
       return { status: 'error', message: 'Mohon lengkapi semua field data personel, pilih kategori, dan unggah file PDF.' };
    }

    const nama = data.nama.trim();
    const nipNup = data.nipNup.trim();
    const unitKerja = data.unitKerja;
    const jenis = data.jenis;
    const kategori = data.kategori;
    const fileData = data.fileData;
    const originalFileName = data.fileName;

     if (nama === '' || nipNup === '' || unitKerja === '' || jenis === '' || kategori === '' || fileData === '') {
       return { status: 'error', message: 'Semua field dan file PDF harus diisi.' };
    }
     // Basic validation for file type
     if (!originalFileName.toLowerCase().endsWith('.pdf') && !fileData.startsWith('data:application/pdf;base64,')) {
         return { status: 'error', message: 'Mohon pilih file dengan format PDF.' };
     }
    // --- End Input Validation ---


    // --- Generate Contract Number and Sequence (Official) ---
    // This call gets the definitive next sequence number and logs the row to the sheet
    const nextSequence = getNextPksSequence(kategori); // Get sequence specific to category for the current year

    // Format the sequence number
    const formattedSequence = Utilities.formatString('%03d', nextSequence);
    // Clean up name for use in file name
    const cleanedNama = nama.replace(/[^a-zA-Z0-9\s]/g, '').replace(/\s+/g, '_').toUpperCase();

    // Generated contract number using the category-specific sequence and current year (implicitly via getNextPksSequence)
    const nomorKontrak = `PKS${formattedSequence}/${cleanedNama}`; // Format PKS###/Nama

    // --- File Processing ---
    // Extract Base64 string from data URL
    const base64String = fileData.split(',')[1];
    const fileBytes = Utilities.base64Decode(base64String);
    // Use generated nomorKontrak for file name
    const blob = Utilities.newBlob(fileBytes, MimeType.PDF, nomorKontrak + ".pdf");

    const destinationFolder = DriveApp.getFolderById(destinationFolderId);

    // Check for potential filename conflicts in the destination folder (optional but good practice)
    // This does NOT prevent createFil e from adding (1), (2), etc. but you could add custom logic here
    // if you needed stricter uniqueness before creation. Standard Drive behavior is usually fine.
    // const existingFiles = destinationFolder.getFilesByName(nomorKontrak + ".pdf");
    // if (existingFiles.hasNext()) {
    //    // Handle conflict, maybe return error or suggest a different name/action
    // }


    // Create the file directly in the destination folder
    const createdFile = destinationFolder.createFile(blob);
    const createdFileName = createdFile.getName();
    const createdFileId = createdFile.getId();

    // --- Log to Google Sheet ---
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      // This should ideally not happen if onOpen works, but good to handle
      throw new Error('Sheet "' + SHEET_NAME + '" not found for logging.');
    }

    const newRow = [
      nextSequence,          // A: Nomor Urut PKS (This sequence is specific to the category in Col G for the current year)
      nomorKontrak,          // B: Nomor Kontrak (Full generated string)
      nama,                  // C: Nama
      nipNup,                // D: NIP/NUP
      unitKerja,             // E: Unit Kerja
      jenis,                 // F: Jenis
      kategori,              // G: Kategori Pegawai
      createdFileName,       // H: Nama File Drive (could be different from nomorKontrak if Drive added (1))
      createdFileId,         // I: ID File Drive
      new Date()             // J: Tanggal Rekam
    ];

    sheet.appendRow(newRow); // Add the new row to the sheet

    return { status: 'success', message: `Rekam kontrak berhasil. File "${createdFileName}" (${createdFileId}) telah disimpan. Nomor Kontrak yang tercatat: ${nomorKontrak}.` }; // Added File ID to message

  } catch (e) {
    Logger.log(e); // Log the error for debugging
    return { status: 'error', message: 'Terjadi Kesalahan saat memproses: ' + e.message };
  }
}

/**
 * Sets the Google Drive destination folder ID in Script Properties.
 * This is a helper function you can run once to set the property.
 * Replace the placeholder ID with your actual destination folder ID.
 */
function setFolderProperties() {
  const scriptProperties = PropertiesService.getScriptProperties();
  // !!! GANTI DENGAN ID FOLDER TUJUAN ANDA YANG SEBENARNYA di baris bawah ini!!!
  const YOUR_ACTUAL_DESTINATION_FOLDER_ID = '1bQwpuVaPVAllgLsUDz2MeI1yvt_fY2v5'; // Anda sudah meletakkannya di sini, bagus!

  // --- KOREKSI: Gunakan nama kunci (key name) yang benar 'DESTINATION_FOLDER_ID' ---
  scriptProperties.setProperty('DESTINATION_FOLDER_ID', YOUR_ACTUAL_DESTINATION_FOLDER_ID);
  // --- Akhir KOREKSI ---


  Logger.log('ID Folder Tujuan disimpan di Script Properties.');
   SpreadsheetApp.getUi().alert('Berhasil', 'ID Folder Tujuan Google Drive telah diatur di properti skrip.', SpreadsheetApp.getUi().ButtonSet.OK);
}