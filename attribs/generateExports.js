function showExportHub() {
  SpreadsheetApp.getActiveSpreadsheet().toast(
    "Generating Export Data",
    "⏳ Please Wait...",
    5
  );
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = spreadsheet.getActiveSheet();
  const sheetName = activeSheet.getName();
  let data = {
    sheetName: sheetName,
    cubeTeams: ''
  };

  const allNamedRanges = spreadsheet.getNamedRanges();

  const namedRangeToDataType = {
    "MAG_CUBE_TEAMS": "cubeTeams"
  };

  allNamedRanges.forEach(range => {
    const rangeName = range.getName();
    if (rangeName in namedRangeToDataType) {
      const dataType = namedRangeToDataType[rangeName];
      data[dataType] = range.getRange().getValue();
    }
  });

  const template = HtmlService.createTemplateFromFile("ExportHubFojji");
  template.data = data;

  const html = template.evaluate().setWidth(900).setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, "Export String Hub");
}