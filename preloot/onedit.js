function onEdit(e) {
  updateTimestamp(e);
}

// --------------------------------------------------
// Update timestamp when column B is edited
// --------------------------------------------------
function updateTimestamp(e) {
  const watchedSheetName = "Ratio Présence/Loot (Préloot)";
  const watchedColumn = 5; // E (presence)
  const watchedColumnBis = 6; // F (bench)

  const targetSheetName = "Admin";
  const targetCell = "B1";

  const editedSheet = e.range.getSheet();

  // Only react to edits in Sheet1
  if (editedSheet.getName() !== watchedSheetName) return;

  // Only react to edits in watched columns
  if (e.range.getColumn() !== watchedColumn && e.range.getColumn() !== watchedColumnBis) return;

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const targetSheet = spreadsheet.getSheetByName(targetSheetName);

  // Update timestamp
  targetSheet.getRange(targetCell).setValue(new Date());

  resetAdminImportRatioCell(e);
}