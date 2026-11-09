function onEdit(e) {
  try {
    var sheet = e.source.getActiveSheet();
    if (sheet.getName() !== "Database") return; // Pastikan hanya berjalan di sheet 'Database'

    var editedCell = e.range;
    var statusColumn = 10; // Kolom J (Status)
    var pksColumn = 9; // Kolom I (No. PKS)

    var row = editedCell.getRow();
    var col = editedCell.getColumn();

    // Jika yang diedit adalah kolom No. PKS (I)
    if (col === pksColumn) {
      var statusCell = sheet.getRange(row, statusColumn);
      
      if (editedCell.getValue().trim() !== "") {
        statusCell.setValue("Penomoran");
      }
    }
  } catch (err) {
    Logger.log("Error: " + err.message);
  }
}
