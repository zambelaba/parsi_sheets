function onEdit(e) {
  if (!e) return;
  var sheet = e.range.getSheet();
  var name = sheet.getName();

  if (e.range.getRow() !== 2) return;
  if (e.range.getColumn() > 36 || e.range.getLastColumn() < 31) return;

  if (name === "Shahraz") toggleShahraz();
  else if (name === "Illidan") toggleIllidan();
}

function toggleShahraz() {
  var sheet = SpreadsheetApp.getActive().getSheetByName("Shahraz");
  var value = sheet.getRange("AE2").getValue();
  var standard = sheet.getRange("BC1:CJ1");
  var fishFountain = sheet.getRange("CK1:DR1");

  if (value === "Standard") {
    sheet.showColumns(standard.getColumn(), standard.getNumColumns());
    sheet.hideColumns(fishFountain.getColumn(), fishFountain.getNumColumns());
  } else if (value === "Fish Fountain") {
    sheet.hideColumns(standard.getColumn(), standard.getNumColumns());
    sheet.showColumns(fishFountain.getColumn(), fishFountain.getNumColumns());
  }
}

function toggleIllidan() {
  var sheet = SpreadsheetApp.getActive().getSheetByName("Illidan");
  if (sheet.getRange("AE2").getValue() === "Lower DPS") {
    sheet.showRows(30, 30);
  } else {
    sheet.hideRows(30, 30);
  }
}