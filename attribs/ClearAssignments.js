/**
 * Menu > Clear Boss Assigns: empties every dropdown cell in AN6:CI38 on every
 * sheet not listed below. Formulas and cells holding a link are kept.
 * (Sheet protection is not checked, despite the function's historical name.)
 */
const CLEAR_ASSIGNMENTS = Object.freeze({
  RANGE: 'AN6:CI38',
  IGNORED_SHEETS: [
    'Raider Data', 'Raid Data', 'Class/Spec Data', 'Static Lists', 'Roster Dropdowns', 'Dynamic Lists',
    'Save Data', 'Intro', 'Roster', 'External Roster Data', 'Guild Manager',
    'Group Builder (25)', 'Group Builder (10)',
  ],
});

function clearValidatedUnprotectedCells() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert('Confirmation',
    'Are you sure you wish to clear? This will remove all assignments on the boss pages.', ui.ButtonSet.YES_NO);
  if (response !== ui.Button.YES) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.getSheets()
    .filter(sheet => CLEAR_ASSIGNMENTS.IGNORED_SHEETS.indexOf(sheet.getName()) === -1)
    .forEach(sheet => {
      const range = sheet.getRange(CLEAR_ASSIGNMENTS.RANGE);
      const validations = range.getDataValidations();
      if (validations.every(row => row.every(rule => rule === null))) return;

      const formulas = range.getFormulas();
      const richTexts = range.getRichTextValues();
      const toClear = [];
      validations.forEach((row, r) => row.forEach((rule, c) => {
        const link = richTexts[r][c] && richTexts[r][c].getLinkUrl();
        if (rule && !formulas[r][c] && !link) toClear.push(a1_(range.getRow() + r, range.getColumn() + c));
      }));
      if (toClear.length) sheet.getRangeList(toClear).clearContent();
    });

  ss.toast('All assignments have been cleared.', 'ℹ️ Assigns Helper', 4);
}
