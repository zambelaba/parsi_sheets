/***** CONSTANTS *****/
const SHEET_ROSTER = "Roster";
const SHEET_IMPORT = "External Roster Data";
const RAIDHELPER_URL_BASE = "https://raid-helper.xyz/raidplan/";

const CLEAR_RANGES = ["AC5:AF64", "AM15:AM18", "AX15:AX22"];
const CHECKBOX_RANGES = ["AB5:AB64"];
const IMPORT_COLUMN_MAPPING = {
  1: 29,
  5: 30,
  6: 31
};

/***** MAIN BUTTON-ACTIVATED FUNCTIONS *****/

function importUsingRaidHelper() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const importSheet = ss.getSheetByName(SHEET_IMPORT);
  const apiKey = importSheet.getRange("K2").getValue().toString().trim();

  const template = HtmlService.createTemplateFromFile("RaidHelperInput");
  template.apiKeyMasked = apiKey ? "*".repeat(apiKey.length) : "";

  const html = template.evaluate().setWidth(600).setHeight(430);
  SpreadsheetApp.getUi().showModalDialog(html, "RaidHelper Import");
}

function processRaidHelperInput(raidHelperId, apiKeyInput) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const importSheet = ss.getSheetByName(SHEET_IMPORT);

  if (!isValidRaidHelperID(raidHelperId)) {
    ui.alert("Error: Invalid RaidHelper ID. Please enter a valid ID.");
    return;
  }

  const storedKey = importSheet.getRange("K2").getValue().toString().trim();

  if (apiKeyInput && apiKeyInput !== "*".repeat(storedKey.length)) {
    importSheet.getRange("K2").setValue(apiKeyInput);
  }

  importSheet.getRange("K1").setValue(raidHelperId);
  SpreadsheetApp.flush();

  const validation = validateRaidHelperData(importSheet, raidHelperId);
  if (validation !== "VALID") return;

  clearRoster(false);
  copyImportData(raidHelperId);
}

function refreshRaidHelperData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const importSheet = ss.getSheetByName(SHEET_IMPORT);

  if (!importSheet) {
    ui.alert(`Sheet "${SHEET_IMPORT}" not found.`);
    return;
  }

  const raidHelperId = importSheet.getRange("K1").getValue().toString().trim();
  if (!raidHelperId) {
    ui.alert("RaidHelper ID is empty. Cannot refresh data. Try again using the 'Import Using RaidHelper' button.");
    return;
  }

  importSheet.getRange("K1").clearContent();
  SpreadsheetApp.flush();
  importSheet.getRange("K1").setValue(raidHelperId);
  SpreadsheetApp.flush();

  const validation = validateRaidHelperData(importSheet, raidHelperId);
  if (validation !== "VALID") return;

  clearRoster(false);
  copyImportData(raidHelperId);
  SpreadsheetApp.getActiveSpreadsheet().toast(
    "✅ RaidHelper Refresh successful!",
    "⚔️ RaidHelper Tools",
    4
  );
}

function clearRoster(withConfirmation = true) {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_ROSTER);

  if (!sheet) {
    ui.alert(`Sheet "${SHEET_ROSTER}" not found.`);
    return;
  }

  if (withConfirmation) {
    const response = ui.alert(
      'Confirm Clearing Roster Data',
      'Are you sure you wish to clear existing roster data?',
      ui.ButtonSet.OK_CANCEL
    );
    if (response !== ui.Button.OK) return;
  }

  CLEAR_RANGES.forEach(range => sheet.getRange(range).clearContent());
  CHECKBOX_RANGES.forEach(range => resetCheckboxRange(sheet, range));

  if (withConfirmation) {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      "Raid composition has been reset.",
      "⚔️ RaidHelper Tools",
      4
    );
  }
}

/***** HELPER FUNCTIONS *****/

function isValidRaidHelperID(id) {
  return /^\d{15,25}$/.test(id);
}

function resetCheckboxRange(sheet, range) {
  const rangeObj = sheet.getRange(range);
  const numRows = rangeObj.getNumRows();
  const numCols = rangeObj.getNumColumns();
  const falseRow = Array(numCols).fill(false);
  const blankCheckboxes = Array.from({ length: numRows }, () => [...falseRow]);
  rangeObj.setValues(blankCheckboxes);
}

function validateRaidHelperData(importSheet, raidHelperId) {
  const ui = SpreadsheetApp.getUi();
  const value = importSheet.getRange("A2").getValue().toString().trim();

  const knownErrors = {
    "": "Unable to fetch RaidHelper data. Please begin the import process again using the 'Import Using RaidHelper' button.",
    "NO COMP BUILT": "RaidHelper has no comp built for this ID.",
    "ERROR: Unexpected end of JSON input": "Unexpected end of data. Possibly a malformed RaidHelper comp.",
    "ERROR: Raid has too many participants": "This raid has too many participants. Please reduce the number in RaidHelper and try again."
  };

  if (knownErrors[value]) {
    ui.alert(knownErrors[value]);
    showRaidHelperLink(raidHelperId);
    return null;
  }

  if (value.includes("returned code 401")) {
    ui.alert("API Key is missing or doesn't match the server your event is listed in. You need to include both the RaidHelper ID AND the correct API Key to import rosters from RaidHelper.");
    return null;
  }

  if (value.includes("returned code 404")) {
    ui.alert("No RaidHelper data found. Please check your RaidHelper ID is correct. First try to reimport, and then Refresh RaidHelper.");
    return null;
  }

  if (value.includes("returned code 502")) {
    ui.alert("RaidHelper is experiencing a temporary server issue — please try again later.");
    return null;
  }

  return "VALID";
}

function showRaidHelperLink(raidHelperId) {
  const fullUrl = RAIDHELPER_URL_BASE + encodeURIComponent(raidHelperId);
  const template = HtmlService.createTemplateFromFile("RaidHelperPopup");
  template.url = fullUrl;

  const htmlOutput = template.evaluate()
    .setWidth(650)
    .setHeight(200);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, "RaidHelper Troubleshooting");
}

function updateCheckboxesInDefinedRange(sheet, startRow, endRow, checkboxCol, valueCol) {
  const numRows = endRow - startRow + 1;
  const values = sheet.getRange(startRow, valueCol, numRows).getValues();
  const checkboxes = values.map(row => [row[0] !== ""]);
  sheet.getRange(startRow, checkboxCol, numRows).setValues(checkboxes);
}

function copyImportData(raidHelperId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName(SHEET_IMPORT);
  const destinationSheet = ss.getSheetByName(SHEET_ROSTER);

  if (!sourceSheet || !destinationSheet) {
    SpreadsheetApp.getUi().alert("Error: One or both target sheets not found.");
    return;
  }

  const sourceStartRow = 2;
  const destinationStartRow = 5;
  const numRows = 60;

  const sourceData = sourceSheet.getRange(sourceStartRow, 1, numRows, 6).getValues();
  const finalData = {
    29: [],
    30: [],
    31: []
  };

  for (let i = 0; i < numRows; i++) {
    const aVal = sourceData[i][0].toString().trim();
    const eVal = sourceData[i][4].toString().trim();
    const fVal = sourceData[i][5].toString().trim();

    if (aVal && (!eVal || !fVal)) {
      showRaidHelperLink(raidHelperId);
      return;
    }

    finalData[29].push([aVal]);
    finalData[30].push([eVal]);
    finalData[31].push([fVal]);
  }

  for (const [col, data] of Object.entries(finalData)) {
    destinationSheet.getRange(destinationStartRow, parseInt(col), numRows, 1).setValues(data);
  }

  updateCheckboxesInDefinedRange(destinationSheet, destinationStartRow, destinationStartRow + numRows - 1, 28, 29);

  SpreadsheetApp.getActiveSpreadsheet().toast(
    "✅ Raid composition imported successfully!",
    "⚔️ RaidHelper Tools",
    4
  );
}