function showForm() {
  const html = HtmlService.createHtmlOutputFromFile("CharacterForm")
    .setWidth(600)
    .setHeight(750)
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi().showModalDialog(html, "Character Submission Form");
}

function getSpecs() {
  const range = SpreadsheetApp.getActiveSpreadsheet().getRangeByName("JOINED_CLASS_AND_SPEC");
  return range ? range.getValues().flat() : [];
}

function submitForm(formData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Guild Manager");
  if (!sheet) throw new Error('❌ Sheet "Guild Manager" not found.');

  const lastRow = sheet.getLastRow();
  let nextRow = 7;
  for (let r = 7; r <= lastRow + 1; r++) {
    const val = sheet.getRange("E" + r).getValue();
    if (!val) {
      nextRow = r;
      break;
    }
  }
  Logger.log("Next available row: " + nextRow);

  const pattern = [
    { char: "M", main: "K", off: "O" },
    { char: "R", main: "P", off: "T" },
    { char: "W", main: "U", off: "Y" },
    { char: "AB", main: "Z", off: "AD" },
    { char: "AG", main: "AE", off: "AI" },
    { char: "AL", main: "AJ", off: "AN" },
    { char: "AQ", main: "AO", off: "AS" },
    { char: "AV", main: "AT", off: "AX" },
    { char: "BA", main: "AY", off: "BC" },
    { char: "BF", main: "BD", off: "BH" },
    { char: "BK", main: "BI", off: "BM" },
  ];


  sheet.getRange("E" + nextRow).setValue(formData.nickname || "");
  Logger.log("Nickname written: " + formData.nickname);

  formData.characters.forEach((entry, i) => {
    if (!entry.characterName || !entry.mainSpec) return;
    const cols = pattern[i];
    if (!cols) return;

    try {
      sheet.getRange(cols.char + nextRow).setValue(entry.characterName);
      sheet.getRange(cols.main + nextRow).setValue(entry.mainSpec);
      sheet.getRange(cols.off + nextRow).setValue(entry.offSpec || "");
      Logger.log(
        `Row ${nextRow}: Wrote ${entry.characterName} (${entry.mainSpec}/${entry.offSpec || ""})`
      );
    } catch (err) {
      Logger.log(`❌ Failed on entry ${i + 1}: ${err}`);
    }
  });

  Logger.log("✅ Finished writing data for " + formData.nickname);
  return `✅ Data submitted successfully to row ${nextRow}.`;
}
