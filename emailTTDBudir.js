function sendEmailWithFiveColumns() {
  const SHEET_NAMES = {
    DATABASE: "Database",
    NOTIFIKASI: "notifikasi direktur",
  };

  const COLUMNS = {
    EMAIL: 10, // Kolom J
    NAMA_PEGAWAI: 3, // Kolom C
    STATUS: 10, // Kolom J
  };

  const RANGES = {
    EMAIL_DESTINATION: "J8",
    DATA_TABLE: "B8:F",
    D7_VALUE: "D5",
  };

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const emailDestination = sheet.getRange(RANGES.EMAIL_DESTINATION).getValue();

  if (!emailDestination) {
    SpreadsheetApp.getUi().alert("Pilih tujuan email terlebih dahulu.");
    return;
  }

  const data = sheet.getRange(RANGES.DATA_TABLE).getValues();
  const d7Value = sheet.getRange(RANGES.D7_VALUE).getValue();

  const { emailBody, namaPegawaiTerkirim } = buildEmailBody(data, d7Value);

  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    "Konfirmasi Pengiriman",
    "Apakah Anda yakin ingin mengirim email?",
    ui.ButtonSet.YES_NO,
  );

  if (response === ui.Button.YES) {
    try {
      MailApp.sendEmail({
        to: emailDestination,
        subject: "Konfirmasi Penandatanganan SPK PTT",
        htmlBody: emailBody,
      });

      updateDatabaseStatus(SHEET_NAMES.DATABASE, namaPegawaiTerkirim, COLUMNS);
      clearNotificationData(SHEET_NAMES.NOTIFIKASI);

      ui.alert(
        "Email berhasil dikirim, status diperbarui, dan data di kolom C8:F pada 'notifikasi direktur' telah dihapus.",
      );
    } catch (error) {
      ui.alert(`Terjadi kesalahan: ${error.message}`);
    }
  } else {
    ui.alert("Pengiriman email dibatalkan.");
  }
}

function buildEmailBody(data, d7Value) {
  let tableContent = `
    <table border='1' cellpadding='5' cellspacing='0' style='border-collapse: collapse;'>
      <tr>
        <th>No</th>
        <th>Nama Pegawai</th>
        <th>Jenis Pegawai</th>
        <th>Unit Kerja/Fakultas</th>
        <th>TMT</th>
      </tr>
  `;

  const namaPegawaiTerkirim = [];
  let no = 1;

  data.forEach((row) => {
    if (row[0] || row[1]) {
      tableContent += `
        <tr>
          <td>${no++}</td>
          <td>${row[1]}</td>
          <td>${row[2] || "Tidak Diketahui"}</td>
          <td>${row[3]}</td>
          <td>${formatDate(row[4])}</td>
        </tr>
      `;
      namaPegawaiTerkirim.push(row[1]);
    }
  });

  tableContent += "</table>";

  const emailBody = `
    <p>Yth. Bapak/Ibu Sekretariat SDMPT,</p>
    <p>Dengan ini kami informasikan bahwa SPK PTT untuk Dosen/Tendik Fakultas/Unit Kerja yang berjumlah <strong>${d7Value}</strong> berkas siap untuk ditandatangani Direktur Sumber Daya Manusia dan Pengembangan Talenta. Berikut nama-nama SPK yang dimaksud:</p>
    ${no === 1 ? "<p>Tidak ada data yang tersedia untuk ditampilkan.</p>" : tableContent}
    <p>Terima kasih</p>
    <p>Hormat Kami,</p>
    <p>Tim Rekrutmen DSDMPT UI</p>
  `;

  return { emailBody, namaPegawaiTerkirim };
}

function updateDatabaseStatus(sheetName, namaPegawaiTerkirim, COLUMNS) {
  const databaseSheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!databaseSheet) return;

  const databaseData = databaseSheet.getDataRange().getValues();
  const updates = [];

  databaseData.forEach((row, i) => {
    if (namaPegawaiTerkirim.includes(row[COLUMNS.NAMA_PEGAWAI - 1])) {
      updates.push([i + 1, COLUMNS.STATUS, "TTD Direktur"]);
    }
  });

  if (updates.length > 0) {
    updates.forEach(([row, col, value]) =>
      databaseSheet.getRange(row, col).setValue(value),
    );
  }
}

function clearNotificationData(sheetName) {
  const notificationSheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!notificationSheet) return;

  const lastRow = notificationSheet.getLastRow();
  if (lastRow >= 8) {
    notificationSheet.getRange(8, 3, lastRow - 7, 4).clearContent();
  }
}

function formatDate(date) {
  if (!(date instanceof Date)) return date;
  const options = { year: "numeric", month: "long", day: "numeric" };
  return date.toLocaleDateString("id-ID", options);
}
