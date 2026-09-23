function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📒 Spreadsheet Tools')
    .addItem('Assign Whole Sheet', 'autoAssignGroups')
    .addItem('Clear Boss Assigns', 'clearValidatedUnprotectedCells')
    .addToUi();
}

/** The strategy dropdown in AE2 on Shahraz and Illidan switches which layout is shown. */
function onEdit(e) {
  if (!e) return;
  const range = e.range;
  if (range.getRow() !== 2 || range.getColumn() > 36 || range.getLastColumn() < 31) return;

  const sheet = range.getSheet();
  if (sheet.getName() === 'Shahraz') toggleShahrazLayout_(sheet);
  else if (sheet.getName() === 'Illidan') toggleIllidanLayout_(sheet);
}

function toggleShahrazLayout_(sheet) {
  const strategy = sheet.getRange('AE2').getValue();
  const standard = sheet.getRange('BC1:CJ1');
  const fishFountain = sheet.getRange('CK1:DR1');
  const show = range => sheet.showColumns(range.getColumn(), range.getNumColumns());
  const hide = range => sheet.hideColumns(range.getColumn(), range.getNumColumns());

  if (strategy === 'Standard') {
    show(standard);
    hide(fishFountain);
  } else if (strategy === 'Fish Fountain') {
    hide(standard);
    show(fishFountain);
  }
}

function toggleIllidanLayout_(sheet) {
  if (sheet.getRange('AE2').getValue() === 'Lower DPS') sheet.showRows(30, 30);
  else sheet.hideRows(30, 30);
}
