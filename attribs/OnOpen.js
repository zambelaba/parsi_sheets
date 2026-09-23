function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui
    .createMenu("📒 Spreadsheet Tools")
    .addSubMenu(
      ui.createMenu('Assignment Helper')
      // .addItem("Save Boss Assignments","showSaveModal")
      // .addItem("Load/Import Boss Assignments","showLoadModal")
      .addItem("Export Save Files", "showSaveDataModal")
      .addItem("Clear Boss Assigns", "clearValidatedUnprotectedCells")
      .addItem("Assign Whole Sheet", "AutoAssignGroups")
    )
    // .addSubMenu(
      // ui.createMenu('Fojji Tools')
      // .addItem("Export Fojji Data", "showExportHub")
      // .addItem("Clear CD Planner","clearCDPlanners")
      // .addItem("Import Fojji CD Profile(s)", "showDecodeModal")
      // .addItem("Help/Info", "showFojjiTutorialPopup")
    // )
    .addSubMenu(
      ui.createMenu('Guild Management')
      .addItem("Add Player", "showForm")
      .addItem("Save Guild Data", "saveGuildData")
      .addItem("Load/Import Guild Data", "showLoadGuildModal")
    )
    .addSubMenu(
      ui.createMenu('Dev')
      .addItem("Check for Updates", "checkVersionAndDisplay")
      .addItem("Refresh Editable Ranges", "refreshEditableRegions")
    )
    .addToUi();
    
  // reloadDropdownCacheAndTrigger();
}