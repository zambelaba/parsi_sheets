function clearValidatedUnprotectedCells() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    "Confirmation",
    "Are you sure you wish to clear? This will remove all assignments on the boss pages.",
    ui.ButtonSet.YES_NO
  );
  if (response !== ui.Button.YES) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetNamesToIgnore = [
    "Raider Data",
    "Raid Data",
    "Class/Spec Data",
    "Static Lists",
    "Raid Helper Import",
    "Roster Dropdowns",
    "CD Planner Validations",
    "Dynamic Lists",
    "Boss Summaries",
    "Save Data",
    "Intro",
    "Guild Manager",
    "Group Builder (25)",
    "Group Builder (10)",
    "Roster",
    "CD Planner",
    "External Roster Data"
  ];

  const sheets = ss.getSheets().filter(s => !sheetNamesToIgnore.includes(s.getName()));
  const targetRangeA1 = "AN6:CI38";

  sheets.forEach(sheet => {
    const range = sheet.getRange(targetRangeA1);
    const values = range.getValues();
    const formulas = range.getFormulas();
    const validations = range.getDataValidations();
    const richTexts = range.getRichTextValues();

    if (validations.flat().every(v => v === null)) return;

    const rows = values.length, cols = values[0].length;
    const mask = new Uint8Array(rows * cols);

    for (let r = 0; r < rows; r++) {
      for (let c = 0; c < cols; c++) {
        if (!formulas[r][c] &&
            validations[r][c] &&
            !richTexts[r][c]?.getLinkUrl()) {
          mask[r * cols + c] = 1;
        }
      }
    }

    const toClear = [];
    for (let r = 0; r < rows; r++) {
      const offset = r * cols;
      if (!mask.subarray(offset, offset + cols).includes(1)) continue;
      for (let c = 0; c < cols; c++) {
        if (mask[offset + c]) {
          toClear.push(range.getCell(r + 1, c + 1).getA1Notation());
        }
      }
    }

    if (toClear.length) sheet.getRangeList(toClear).clearContent();
  });

  ss.toast("All assignments have been cleared.", "ℹ️ Assigns Helper", 4);
}

function clearCDPlanners() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    "Confirmation",
    "Are you sure you wish to clear? This will remove all  assignments on the Raid CD Planner.",
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var validSheets = ["CD Planner"];

  var allNamedRanges = ss.getNamedRanges();
  var planners = {};
  validSheets.forEach(function(name) {
    planners[name] = allNamedRanges
      .filter(function(nr) {
        var r = nr.getRange();
        return r.getSheet().getName() === name && /(HEALTH_AREA|TIME_AREA)$/.test(nr.getName());
      })
      .map(function(nr) { return nr.getName(); });
  });

  SpreadsheetApp.getActiveSpreadsheet().toast(
    `Clearing Planner Page(s).`,
    "⏳ Please Wait...",
    13
  );

  validSheets.forEach(function(name) {
    var sheet = ss.getSheetByName(name);
    if (sheet) clearPlannerRanges(sheet, planners[name]);
  });

  SpreadsheetApp.flush();

  SpreadsheetApp.getActiveSpreadsheet().toast(
    `✅ Planner Page(s) Cleared!`,
    "📋 Planner Helper",
    5
  );
}

function clearPlannerRanges(sheet, rangeNames) {
  var ss = sheet.getParent();
  var allNamedRanges = ss.getNamedRanges();
  var namedRangeMap = {};
  for (var i = 0; i < allNamedRanges.length; i++) {
    namedRangeMap[allNamedRanges[i].getName()] = allNamedRanges[i].getRange();
  }

  var ranges = rangeNames.map(function(n) { return namedRangeMap[n]; }).filter(function(r) { return r; });
  if (!ranges.length) return;

  var minRow = Math.min(...ranges.map(r => r.getRow()));
  var maxRow = Math.max(...ranges.map(r => r.getLastRow()));
  var minCol = Math.min(...ranges.map(r => r.getColumn()));
  var maxCol = Math.max(...ranges.map(r => r.getLastColumn()));

  var numRows = maxRow - minRow + 1;
  var numCols = maxCol - minCol + 1;
  var totalCells = numRows * numCols;

  var block = sheet.getRange(minRow, minCol, numRows, numCols);
  var formulas = block.getFormulas();
  var values = block.getValues();

  var mask = new Uint8Array(totalCells);
  ranges.forEach(function(range) {
    var r0 = range.getRow() - minRow;
    var c0 = range.getColumn() - minCol;
    var w = range.getNumColumns();
    var h = range.getNumRows();
    for (var r = 0; r < h; r++) {
      mask.fill(1, (r0 + r) * numCols + c0, (r0 + r) * numCols + c0 + w);
    }
  });

  for (var r = 0; r < numRows; r++) {
    var rowOffset = r * numCols;
    var rowMask = mask.subarray(rowOffset, rowOffset + numCols);
    if (!rowMask.includes(1)) continue;
    for (var c = 0; c < numCols; c++) {
      if (!rowMask[c]) continue;
      values[r][c] = formulas[r][c] || "";
    }
  }

  var chunkSize = 10000;
  var startRow = 0;
  while (startRow < numRows) {
    var endRow = Math.min(numRows, startRow + Math.floor(chunkSize / numCols));
    var rows = endRow - startRow;
    if (rows > 0) {
      sheet
        .getRange(minRow + startRow, minCol, rows, numCols)
        .setValues(values.slice(startRow, endRow));
    }
    startRow = endRow;
  }
}