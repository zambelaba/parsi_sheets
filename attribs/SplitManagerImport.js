const ROSTER_IMPORT_SHEET = "External Roster Data"
const ROSTER_TARGET_SHEET = "Roster"
const ROSTER_TARGET_RANGE = "AC5:AF64"
const ROSTER_CHECKBOX_RANGE = "AB5:AB64"

const ROSTER_SPLIT_RANGES = {
  "Split 1": "O2:R31",
  "Split 2": "W2:Z31",
  "Split 3": "AB2:AE31",
  "Split 4": "AG2:AJ31",
  "Split 5": "AL2:AQ31"
}

function showRosterImportModal() {
  const availableSplits = getAvailableSplits(); 
  const template = HtmlService.createTemplateFromFile('RosterImportPopup');
  template.splitOptions = availableSplits;
  const html = template.evaluate().setWidth(400).setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(html, 'Import Roster Split');
}

function getAvailableSplits() {
  return Object.keys(ROSTER_SPLIT_RANGES).sort(function(a,b){
    return (parseInt(a.split(' ')[1])||0)-(parseInt(b.split(' ')[1])||0)
  })
}

function importRosterSplit(splitName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const sourceSheet = ss.getSheetByName(ROSTER_IMPORT_SHEET)
  const targetSheet = ss.getSheetByName(ROSTER_TARGET_SHEET)

  if (!sourceSheet || !targetSheet) {
    toastMessage("⚔️ Roster Import", "Missing required sheets.", 5)
    return
  }

  const sourceRangeA1 = ROSTER_SPLIT_RANGES[splitName]
  if (!sourceRangeA1) {
    toastMessage("⚔️ Roster Import", "Invalid split selected.", 5)
    return
  }

  const sourceRange = sourceSheet.getRange(sourceRangeA1)
  const sourceValues = sourceRange.getValues()

  const targetRange = targetSheet.getRange(ROSTER_TARGET_RANGE)
  targetRange.clearContent()
  resetCheckboxRange(targetSheet, ROSTER_CHECKBOX_RANGE)

  const numRows = Math.min(sourceValues.length, targetRange.getNumRows())
  const numCols = Math.min(sourceValues[0].length, targetRange.getNumColumns())
  const trimmed = sourceValues.slice(0, numRows).map(r => r.slice(0, numCols))

  const startRow = targetRange.getRow()
  const startCol = targetRange.getColumn()
  targetSheet.getRange(startRow, startCol, numRows, numCols).setValues(trimmed)

  updateCheckboxesInDefinedRange(targetSheet, 5, 64, 28, 29)

  toastMessage("⚔️ Roster Import", `✅ ${splitName} imported successfully!`, 5)
}