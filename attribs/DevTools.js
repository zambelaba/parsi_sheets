// COLUMN ADJUSTER FUNCTION | FOR PERFORMING LOTS OF SPECIFIC COLUMN WIDTH ADJUSTMENTS | TYPE YOUR COLUMN WIDTHS BELOW
function setCustomColumnWidthsInteractive() {
  const ui = SpreadsheetApp.getUi();

  const response = ui.prompt("Set Column Widths", "Enter the starting column letter (e.g., A, B, C):", ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) return;

  const startLetter = response.getResponseText().trim().toUpperCase();
  if (!/^[A-Z]{1,3}$/.test(startLetter)) { // A basic validation for columns A–L
    ui.alert("Invalid input. Please enter a single column letter (A–CZ).");
    return;
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const startIndex = columnLetterToIndex(startLetter);

  const columnWidths = [129, 195, 216, 200, 120, 112, 112, 130, 160, 140, 140]; 

  for (let i = 0; i < columnWidths.length; i++) {
    const colIndex = startIndex + i;
    if (colIndex > sheet.getMaxColumns()) break; // avoid setting past existing columns
    sheet.setColumnWidth(colIndex, columnWidths[i]);
  }

  ui.alert(`Column widths set starting from column "${startLetter}".`);
}

function columnLetterToIndex(letter) {
  return letter.split('').reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0);
}

// FIX THAT DAMN ALTERNATING COLOUR SHENANIGANS THAT CAUSE THE WEIRD CELL BEHAVIOUR
function removeAlternatingColorsFromSpecificSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var targetSheets = [
    "Stoneguards", "Feng", "Gara'jal", "Spirit Kings", "Elegon",
    "Will", "Zor'lok", "Ta'yak", "Garalon", "Mel'jarak",
    "Un'sok", "Shek'zeer", "MSV CD Planner", "HoF CD Planner", "ToeS CD Planner"
  ];

  Logger.log("Starting alternating color cleanup...");
  Logger.log("Target sheets: " + targetSheets.join(", "));

  targetSheets.forEach(function(name) {
    var sheet = ss.getSheetByName(name);
    if (!sheet) {
      Logger.log("❌ Sheet not found: " + name);
      return;
    }

    var bandings = sheet.getBandings();
    if (bandings.length === 0) {
      Logger.log("ℹ️ No bandings found on sheet: " + name);
      return;
    }

    Logger.log("🟢 Processing sheet: " + name + " (" + bandings.length + " bandings)");

    bandings.forEach(function(banding, index) {
      banding.remove();
      Logger.log("   ✅ Removed banding " + (index + 1) + " of " + bandings.length);
    });

    Logger.log("   ✅ Finished sheet: " + name);
  });

  Logger.log("🎉 Cleanup complete!");
}

// REMOVE NAMED RANGES THAT NO LONGER REFERENCE ANYTHING
function removeRefErrorNamedRanges() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const namedRanges = ss.getNamedRanges();
  let removedCount = 0;

  namedRanges.forEach(namedRange => {
    let remove = false;
    try {
      const range = namedRange.getRange();
      const sheet = range.getSheet();

      // Check if sheet exists (null if deleted)
      if (!sheet) {
        remove = true;
      }

      // Check if A1 notation contains '#REF!'
      const a1 = range.getA1Notation();
      if (a1.includes('#REF')) {
        remove = true;
      }

    } catch (e) {
      // getRange() failed → definitely broken
      remove = true;
    }

    if (remove) {
      Logger.log(`Removing broken named range: "${namedRange.getName()}"`);
      namedRange.remove();
      removedCount++;
    }
  });

  SpreadsheetApp.getUi().alert(`${removedCount} broken named range(s) removed.`);
}
