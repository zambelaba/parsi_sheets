/**
 * Roster page buttons: import a RaidHelper composition, refresh it, clear the
 * roster, and show the instructions.
 *
 * The composition is fetched by an IMPORTJSONAPI formula in External Roster
 * Data!A2, which reads the event ID from K1 and the API key from K2. Writing K1
 * is what triggers the fetch; the result is then copied into the Roster table.
 */
const RAID_HELPER = Object.freeze({
  IMPORT_SHEET: 'External Roster Data',
  ROSTER_SHEET: 'Roster',
  ID_CELL: 'K1',
  API_KEY_CELL: 'K2',
  STATUS_CELL: 'A2',
  PLAN_URL: 'https://raid-helper.xyz/raidplan/',
  API_KEY_MASK: '••••••••',
  // Roster table: "Is Rostered" checkbox, then Player, Class, Spec, Flex Spec.
  ROSTER_ROWS: 60,
  ROSTER_FIRST_ROW: 5,
  ROSTER_CHECKBOX_COLUMN: 28, // AB
  ROSTER_PLAYER_COLUMN: 29,   // AC
  ROSTER_EXTRA_CLEAR_RANGES: ['AM15:AM18', 'AX15:AX22'],
  KNOWN_ERRORS: {
    '': "Unable to fetch RaidHelper data. Please begin the import process again using the 'Import Using RaidHelper' button.",
    'NO COMP BUILT': 'RaidHelper has no comp built for this ID.',
    'ERROR: Unexpected end of JSON input': 'Unexpected end of data. Possibly a malformed RaidHelper comp.',
    'ERROR: Raid has too many participants': 'This raid has too many participants. Please reduce the number in RaidHelper and try again.',
  },
});

/** IMPORT button: asks for the event ID and API key. */
function importUsingRaidHelper() {
  const importSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(RAID_HELPER.IMPORT_SHEET);
  const hasKey = String(importSheet.getRange(RAID_HELPER.API_KEY_CELL).getValue()).trim() !== '';

  const template = HtmlService.createTemplateFromFile('RaidHelperInput');
  template.apiKeyMasked = hasKey ? RAID_HELPER.API_KEY_MASK : '';
  SpreadsheetApp.getUi().showModalDialog(template.evaluate().setWidth(600).setHeight(430), 'RaidHelper Import');
}

/** Called by RaidHelperInput.html. */
function processRaidHelperInput(raidHelperId, apiKeyInput) {
  const ui = SpreadsheetApp.getUi();
  const importSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(RAID_HELPER.IMPORT_SHEET);

  if (!/^\d{15,25}$/.test(raidHelperId)) {
    ui.alert('Error: Invalid RaidHelper ID. Please enter a valid ID.');
    return;
  }

  if (apiKeyInput && apiKeyInput !== RAID_HELPER.API_KEY_MASK) {
    importSheet.getRange(RAID_HELPER.API_KEY_CELL).setValue(apiKeyInput);
  }
  importSheet.getRange(RAID_HELPER.ID_CELL).setValue(raidHelperId);
  SpreadsheetApp.flush();

  if (!raidHelperDataIsValid_(importSheet, raidHelperId)) return;
  clearRoster(false);
  copyRaidHelperRoster_(raidHelperId);
}

/** REFRESH DATA button: re-fetches the last imported event. */
function refreshRaidHelperData() {
  const ui = SpreadsheetApp.getUi();
  const importSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(RAID_HELPER.IMPORT_SHEET);
  if (!importSheet) {
    ui.alert(`Sheet "${RAID_HELPER.IMPORT_SHEET}" not found.`);
    return;
  }

  const idCell = importSheet.getRange(RAID_HELPER.ID_CELL);
  const raidHelperId = String(idCell.getValue()).trim();
  if (!raidHelperId) {
    ui.alert("RaidHelper ID is empty. Cannot refresh data. Try again using the 'Import Using RaidHelper' button.");
    return;
  }

  // Rewriting the same ID does not recalculate the formula; clearing it first does.
  idCell.clearContent();
  SpreadsheetApp.flush();
  idCell.setValue(raidHelperId);
  SpreadsheetApp.flush();

  if (!raidHelperDataIsValid_(importSheet, raidHelperId)) return;
  clearRoster(false);
  if (copyRaidHelperRoster_(raidHelperId)) {
    SpreadsheetApp.getActiveSpreadsheet().toast('✅ RaidHelper Refresh successful!', '⚔️ RaidHelper Tools', 4);
  }
}

/** CLEAR ROSTER button. */
function clearRoster(withConfirmation = true) {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(RAID_HELPER.ROSTER_SHEET);
  if (!sheet) {
    ui.alert(`Sheet "${RAID_HELPER.ROSTER_SHEET}" not found.`);
    return;
  }

  if (withConfirmation) {
    const response = ui.alert('Confirm Clearing Roster Data',
      'Are you sure you wish to clear existing roster data?', ui.ButtonSet.OK_CANCEL);
    if (response !== ui.Button.OK) return;
  }

  rosterColumns_(sheet, RAID_HELPER.ROSTER_PLAYER_COLUMN, 4).clearContent();
  RAID_HELPER.ROSTER_EXTRA_CLEAR_RANGES.forEach(a1 => sheet.getRange(a1).clearContent());
  rosterColumns_(sheet, RAID_HELPER.ROSTER_CHECKBOX_COLUMN, 1)
    .setValues(new Array(RAID_HELPER.ROSTER_ROWS).fill(null).map(() => [false]));

  if (withConfirmation) {
    SpreadsheetApp.getActiveSpreadsheet().toast('Raid composition has been reset.', '⚔️ RaidHelper Tools', 4);
  }
}

/** INSTRUCTIONS button. */
function showRosterTutorialPopup() {
  const html = HtmlService.createTemplateFromFile('RosterInstructions').evaluate().setWidth(800).setHeight(1200);
  SpreadsheetApp.getUi().showModalDialog(html, 'Roster Import Instructions');
}

/** "Open Instructions" on the auto-assign "No Roster Setup!" dialog. */
function openRosterInstructions() {
  const roster = SpreadsheetApp.getActive().getSheetByName(RAID_HELPER.ROSTER_SHEET);
  if (roster) roster.activate();
  showRosterTutorialPopup();
}

/** `width` columns of the Roster table starting at `column`, over every roster row. */
function rosterColumns_(sheet, column, width) {
  return sheet.getRange(RAID_HELPER.ROSTER_FIRST_ROW, column, RAID_HELPER.ROSTER_ROWS, width);
}

/** Alerts and returns false when the IMPORTJSONAPI formula reports an error. */
function raidHelperDataIsValid_(importSheet, raidHelperId) {
  const ui = SpreadsheetApp.getUi();
  const status = String(importSheet.getRange(RAID_HELPER.STATUS_CELL).getValue()).trim();

  if (Object.prototype.hasOwnProperty.call(RAID_HELPER.KNOWN_ERRORS, status)) {
    ui.alert(RAID_HELPER.KNOWN_ERRORS[status]);
    showRaidHelperLink_(raidHelperId);
    return false;
  }
  if (status.includes('returned code 401')) {
    ui.alert("API Key is missing or doesn't match the server your event is listed in. You need to include both the RaidHelper ID AND the correct API Key to import rosters from RaidHelper.");
    return false;
  }
  if (status.includes('returned code 404')) {
    ui.alert('No RaidHelper data found. Please check your RaidHelper ID is correct. First try to reimport, and then Refresh RaidHelper.');
    return false;
  }
  if (status.includes('returned code 502')) {
    ui.alert('RaidHelper is experiencing a temporary server issue — please try again later.');
    return false;
  }
  return true;
}

/**
 * Copies name / corrected class / corrected spec (External Roster Data A, E, F)
 * into the Roster table and ticks "Is Rostered" for every imported player.
 * Stops without writing if any player has no valid class or spec.
 */
function copyRaidHelperRoster_(raidHelperId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const source = ss.getSheetByName(RAID_HELPER.IMPORT_SHEET);
  const roster = ss.getSheetByName(RAID_HELPER.ROSTER_SHEET);
  if (!source || !roster) {
    SpreadsheetApp.getUi().alert('Error: One or both target sheets not found.');
    return false;
  }

  const rows = source.getRange(2, 1, RAID_HELPER.ROSTER_ROWS, 6).getValues()
    .map(row => [row[0], row[4], row[5]].map(value => String(value).trim()));

  if (rows.some(([name, cls, spec]) => name && (!cls || !spec))) {
    showRaidHelperLink_(raidHelperId);
    return false;
  }

  rosterColumns_(roster, RAID_HELPER.ROSTER_PLAYER_COLUMN, 3).setValues(rows);
  rosterColumns_(roster, RAID_HELPER.ROSTER_CHECKBOX_COLUMN, 1).setValues(rows.map(([name]) => [name !== '']));
  ss.toast('✅ Raid composition imported successfully!', '⚔️ RaidHelper Tools', 4);
  return true;
}

/** Explains that the composition is empty or has invalid players, with a link to fix it. */
function showRaidHelperLink_(raidHelperId) {
  const template = HtmlService.createTemplateFromFile('RaidHelperPopup');
  template.url = RAID_HELPER.PLAN_URL + encodeURIComponent(raidHelperId);
  SpreadsheetApp.getUi().showModalDialog(template.evaluate().setWidth(650).setHeight(200), 'RaidHelper Troubleshooting');
}
