function showFojjiTutorialPopup() {
  const template = HtmlService.createTemplateFromFile("FojjiInstructions");
  const html = template.evaluate().setWidth(800).setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(html, "Fojji Assigns Instructions");
}

function showRosterTutorialPopup(){
  const template = HtmlService.createTemplateFromFile("RosterInstructions");
  const html = template.evaluate().setWidth(800).setHeight(1200);
  SpreadsheetApp.getUi().showModalDialog(html, "Roster Import Instructions");
}