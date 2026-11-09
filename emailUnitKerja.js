// Fungsi untuk mengonversi tanggal ke format "DD MMMM YYYY"
function formatDate(date) {
  const monthsOfYear = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
  const day = date.getDate();
  const month = monthsOfYear[date.getMonth()];
  const year = date.getFullYear();
  return `${day} ${month} ${year}`;
}

// Fungsi utama untuk mengirim email dengan format HTML
function sendEmailWithHTML() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // Ambil data dari sel yang berisi dropdown tujuan email (misalnya kolom J11)
  const emailDestination = sheet.getRange("J11").getValue();
  
  // Pastikan alamat email valid
  if (!emailDestination) {
    SpreadsheetApp.getUi().alert("Pilih tujuan email terlebih dahulu.");
    return;
  }

  // Ambil data tabel yang akan dikirim (B12:D1000 untuk Tendik dan TMT Tendik, E12:G1000 untuk Dosen dan TMT Dosen)
  const dataRange1 = sheet.getRange("B12:D1000"); // Range untuk "Tendik" dan "TMT Tendik"
  const dataRange2 = sheet.getRange("E12:G1000"); // Range untuk "Dosen" dan "TMT Dosen"

  const data1 = dataRange1.getValues();  // Data Tendik
  const data2 = dataRange2.getValues();  // Data Dosen

  // Ambil nilai dari kolom D7
const d7Value = sheet.getRange("D7").getValue();

// Menyusun body email dengan HTML (menggunakan <table> untuk menjaga format)
let emailBody = "<p>Yth. Bapak/Ibu,</p>";
emailBody += `<p>Dengan ini kami informasikan bahwa SPK PTT untuk Dosen/Tendik <strong>${d7Value}</strong> sebagaimana tercantum pada tabel di bawah ini telah siap untuk diambil di Gedung Pusat Administrasi Universitas Indonesia, Lantai 6 Direktorat Sumber Daya Manusia, Seksi Penempatan Perencanaan Pengembangan SDM.</p>`;


  // Memulai tabel jika ada data yang akan ditambahkan
  let tableContent = "<table border='1' cellpadding='5' cellspacing='0'>";
  tableContent += "<tr><th>No</th><th>Nama Pegawai</th><th>Jenis Pegawai</th><th>TMT</th></tr>";

  let hasData = false;  // Penanda jika ada data yang valid untuk dimasukkan ke dalam tabel
  let no = 1;  // Menambahkan penghitung nomor urut

  // Array untuk menampung data Dosen dan Tendik terpisah
  let dosenData = [];
  let tendikData = [];

  // Pisahkan data Dosen dan Tendik
  for (let i = 0; i < data1.length; i++) {
    const row1 = data1[i];  // Data untuk "Tendik" dan "TMT Tendik"
    const row2 = data2[i];  // Data untuk "Dosen" dan "TMT Dosen"
    
    // Memasukkan data Tendik ke dalam array tendikData
    if (row1[0] || row1[1]) {
      tendikData.push({name: row1[1], tmt: row1[2], type: 'Tendik'});
    }

    // Memasukkan data Dosen ke dalam array dosenData
    if (row2[0] || row2[1]) {
      dosenData.push({name: row2[1], tmt: row2[2], type: 'Dosen'});
    }
  }

  // Tambahkan semua data Dosen terlebih dahulu
  dosenData.forEach((data) => {
    tableContent += "<tr>";
    tableContent += "<td>" + no++ + "</td>"; // Menambahkan nomor urut
    tableContent += "<td>" + data.name + "</td>";
    tableContent += "<td>" + data.type + "</td>";
    tableContent += "<td>" + (data.tmt instanceof Date ? formatDate(data.tmt) : data.tmt) + "</td>";
    tableContent += "</tr>";
  });

  // Tambahkan semua data Tendik setelahnya
  tendikData.forEach((data) => {
    tableContent += "<tr>";
    tableContent += "<td>" + no++ + "</td>"; // Menambahkan nomor urut
    tableContent += "<td>" + data.name + "</td>";
    tableContent += "<td>" + data.type + "</td>";
    tableContent += "<td>" + (data.tmt instanceof Date ? formatDate(data.tmt) : data.tmt) + "</td>";
    tableContent += "</tr>";
  });

  tableContent += "</table>";

  // Jika ada data, tambahkan tabel ke emailBody, jika tidak tampilkan pesan tanpa tabel
  if (dosenData.length > 0 || tendikData.length > 0) {
    emailBody += tableContent;
  } else {
    emailBody += "<p>Tidak ada data yang tersedia untuk ditampilkan.</p>";
  }

  emailBody += "<p>Kami mengucapkan terima kasih atas perhatian dan kerja sama Bapak/Ibu.</p>";
  emailBody += "<p>Hormat Kami,</p>";
  emailBody += "<p>Tim Rekrutmen DSDM UI</p>";


  // Konfirmasi sebelum mengirim email
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert("Konfirmasi Pengiriman", "Apakah Anda yakin ingin mengirim email?", ui.ButtonSet.YES_NO);
  
  if (response == ui.Button.YES) {
    MailApp.sendEmail({
      to: emailDestination,
      subject: "Konfirmasi Pengambilan SPK PTT di Gedung PAU UI",
      htmlBody: emailBody
    });
    
    // Mengupdate status di sheet "Database" kolom J menjadi "Siap Diambil"
    const databaseSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Database");
    const formatSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Format Pengambilan");
    const formatData = formatSheet.getRange("C12:F1000").getValues(); // Kolom C dan F di "Format Pengambilan"
    const databaseData = databaseSheet.getRange("C4:J" + databaseSheet.getLastRow()).getValues(); // Kolom C ke J di "Database"
    
    // Proses setiap baris di "Format Pengambilan" untuk mencari kesesuaian di "Database"
    formatData.forEach(rowFormat => {
      const idC = rowFormat[0]; // Kolom C (nama pertama)
      const idF = rowFormat[3]; // Kolom F (nama kedua)

      for (let i = 0; i < databaseData.length; i++) {
        if (databaseData[i][0] === idC || databaseData[i][0] === idF) { // Cocokkan kolom C di "Database" dengan C atau F di "Format Pengambilan"
          databaseSheet.getRange(i + 4, 10).setValue("Siap Diambil"); // Kolom J di "Database" adalah kolom ke-10
        }
      }
    });

    // Menghapus data setelah email terkirim
    for (let row = 12; row <= 1000; row++) {
      sheet.getRange(row, 3).clearContent(); // Kolom C
      sheet.getRange(row, 6).clearContent(); // Kolom F
    }
    sheet.getRange("D7").clearContent();

    ui.alert("Email berhasil dikirim dan status telah diperbarui.");
  } else {
    ui.alert("Pengiriman email dibatalkan.");
  }
}
