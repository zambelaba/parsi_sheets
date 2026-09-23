function checkVersionAndDisplay() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const localNamedRange = "LOCAL_VERSION";
  const remoteNamedRange = "LATEST_VERSION";
  const versionNumberFormula = sheet.getRangeByName(remoteNamedRange).getFormula();

  sheet.getRangeByName(remoteNamedRange).clearContent();
  // SpreadsheetApp.flush();
  sheet.getRangeByName(remoteNamedRange).setValue(versionNumberFormula);
  // SpreadsheetApp.flush();

  const localVersion = sheet.getRangeByName(localNamedRange).getValue().toString().trim();
  const remoteVersion = sheet.getRangeByName(remoteNamedRange).getValue().toString().trim();

  if (!isValidVersion(localVersion) || !isValidVersion(remoteVersion)) {
    ui.alert("Version Check Error","Could not find version ID. Please check the Introduction Page.")
    return;
  }

  const result = compareVersions(localVersion, remoteVersion);

  if (result === "same") return;

  const html = HtmlService.createTemplateFromFile("VersionCheckerDialog");
  html.result = result;
  html.localVersion = localVersion;
  html.remoteVersion = remoteVersion;

  ui.showModalDialog(html.evaluate().setWidth(480).setHeight(260), "Version Update Notice");
}

function isValidVersion(version) {
  return /^\d+\.\d+\.\d+$/.test(version);
}


function compareVersions(local, remote) {
  const [lMajor, lMinor, lPatch] = local.split(".").map(Number);
  const [rMajor, rMinor, rPatch] = remote.split(".").map(Number);

  if (rMajor > lMajor) return "major";
  if (rMajor === lMajor && rMinor > lMinor) return "minor";
  if (rMajor === lMajor && rMinor === lMinor && rPatch > lPatch) return "patch";

  return "same";
}