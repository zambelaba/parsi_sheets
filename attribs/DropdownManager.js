function reloadDropdownCacheAndTrigger() {
  const fojjiColumnPairs = [
      [8, 12],
  ];
  for (let key in plannerToValidationSheetMap) {
    //Logger.log(key);
    setupValidationRanges(key, fojjiColumnPairs);
    
  }
  SpreadsheetApp.getActiveSpreadsheet().toast(
    `✅ CD Planner Dropdowns Set.`,
    "📋 Dropdown Manager",
    10
    );
}

const plannerToValidationSheetMap = {
"CD Planner": "CD Planner Validations"
}

function testEdit() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const range = sheet.getRange("E5");  
  ss.setActiveRange(range);
  onEdit({ source: ss });
}

// function onEdit(e) {
  //const sheet = e.source.getActiveSheet();
  //Logger.log(`onEdit: ${sheet.getName()}`);
  // Only respond on 'Raid CD Planner' sheet
  //if (sheet.getName() in plannerToValidationSheetMap) {
    //const fojjiColumnPairs = [
      //[8, 12],   // D → F
    //];
    //return onRaidCDEdit(sheet, sheet.getName(key), fojjiColumnPairs);
  //} 
//}

function onRaidCDEdit(sheet, sheetName, columnPairs) {
  const editedRange = sheet.getActiveCell();
  const col = editedRange.getColumn();
  //Logger.log('row rowindex col colindex A1notation %s %s %s %s %s', editedRange.getRow(), editedRange.getRowIndex(), col, editedRange.getColumnIndex(), editedRange.getA1Notation());

  // Break out early if edited cell wasn't a spell cell
  if (!(col in Object.values(columnPairs))) return;
  
  // Get the spell cell to check validation on
  const spellCell = editedRange;
  var rule = spellCell.getDataValidation();
  var currentValidationIndex = 0;
  if (rule) {
    const criteria = rule.getCriteriaType();
    const args = rule.getCriteriaValues();
    if (criteria == "VALUE_IN_RANGE") { // '==' to implict cast the enum to string
      // expect args to be [ range, true ]
      const range = args[0];
      currentValidationIndex = args[0].getRow();
      //Logger.log(`${range.getRow()}, ${range.getA1Notation()}, ${range.getA1Notation()}`);
      //Logger.log(`The data validation range row is ${args[0].getRow()}`);
    }
  }
  const correctValidationIndex = getValidationIndex(spellCell.getA1Notation(), sheetName);
  //Logger.log(`currentValidationIndex=${currentValidationIndex}, correctValidationIndex=${correctValidationIndex}`);
  if (!correctValidationIndex) {
    // FUBAR ;_; can't self-heal it's too broken
    return;
  }

  if (currentValidationIndex === correctValidationIndex) {
    //Logger.log('\o/ validation unchanged, nothing to do');
    return;
  }

  // if we're here, goldilocks has been messing with the sheet... but we can fix it
  const MAX_SPELL_COUNT = 400;
  const validationSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(plannerToValidationSheetMap[sheetName]);
  rule = SpreadsheetApp.newDataValidation() 
          .requireValueInRange(validationSheet.getRange(correctValidationIndex,1,1,MAX_SPELL_COUNT))
          .setAllowInvalid(true)
          .build();

  spellCell.setDataValidation(rule);

  //Logger.log(`Corrected validation from ${currentValidationIndex} to ${correctValidationIndex}`);
  SpreadsheetApp.getActiveSpreadsheet().toast(
    `You mongo. I had to go in and fix the ${editedRange.getA1Notation()} dropdown for you. >:| Don't paste like that!!! Use "Paste Values" (CTRL + SHIFT + V)`,
    "❌ Raid CD Planner",
    10
  );
}

function numToAZ(num) {
  // warning: only works up to Z
  // warning: 1-indexed
  return String.fromCharCode(64 + (num % 26));
}

function setupValidationRanges(sheetName, columnPairs)
{
  // sets up the validation for each of the Spell cells
  // by setting up a separate sheet (Validations) that dynamically
  // populates each list as a horizontal range based on the corresponding Player selected
  // e.g. for cell E5 that needs to have its validation match the player in D5
  // this will set up 1 row in the Validations sheet and setDataValidation to the horizontal range(validationIndex, 1, 1, maximum spell count) 
  // note: the validationIndex will pretty much never match the player/cell row (and for the MRT sheet, it cannot)
  const MAX_SPELL_COUNT = 200; // TODO: should get this number from the named range instead
  const inputSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  const validationSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(plannerToValidationSheetMap[sheetName]);
  
  const lastRow = inputSheet.getLastRow();
  var validationSheetIndex = 1;
  var spellCellsToValidation = {};
  for (let pairIndex = 0; pairIndex < columnPairs.length; ++pairIndex) {
    let playerCol = columnPairs[pairIndex][0];
    let spellCol = columnPairs[pairIndex][1];
    const spellRange = inputSheet.getRange(1, spellCol, lastRow, 1);
    const rules = spellRange.getDataValidations();

    // rules index 0 = cell 1
    for (let j = 0; j < rules.length; j++) {
      //Logger.log(`${rules[j][0]} ${typeof rules[j][0]} ${rules[j][0] == null}`);
      // Some bullshittery magic that just works(TM) If not null then validation exists
      if (rules[j][0] != null ) {
        // e.g. row 5 that has a player and a spell cell in it
        const playerCellA1Notation = numToAZ(playerCol) + String(j+1);
        const spellCellA1Notation = numToAZ(spellCol) + String(j+1);
        // 1. Tries to find player, if error then
        //   2. Tries to find class, if error then
        //      3. Returns all the spells that aren't the header name 
        // TODO: probably don't need the header exclusion with named ranges
        var dropdownRangeCellFormula = 
        `=IF('${sheetName}'!${playerCellA1Notation}="spellName","",IF(OR('${sheetName}'!${playerCellA1Notation}="All",'${sheetName}'!${playerCellA1Notation}="",'${sheetName}'!${playerCellA1Notation}="MeleeDPS",'${sheetName}'!${playerCellA1Notation}="RangedDPS",'${sheetName}'!${playerCellA1Notation}="Healers",'${sheetName}'!${playerCellA1Notation}="Tanks"),TRANSPOSE(SORT(UNIQUE(CD_Index[spellName]))),IF(COUNTIF(Roster_Map[specKey],'${sheetName}'!${playerCellA1Notation})>0,TRANSPOSE(FILTER(CD_Index[spellName],CD_Index[spellClass]=XLOOKUP('${sheetName}'!${playerCellA1Notation},Roster_Map[specKey],Roster_Map[Class]))),IF(COUNTIF(Dynamic_Lists[fojjiOptions],'${sheetName}'!${playerCellA1Notation})>0,TRANSPOSE(FILTER(CD_Index[spellName],CD_Index[spellClass]='${sheetName}'!${playerCellA1Notation})),TRANSPOSE(SORT(UNIQUE(CD_Index[spellName])))))))`;
        const validationCell1 = validationSheet.getRange(validationSheetIndex, 1);
        validationCell1.setValue(dropdownRangeCellFormula); 
        // the TRANSPOSE function makes this value in validationCell1 a horizontal cell array - set that as the dropdown criteria
        const validationRange = validationSheet.getRange(validationSheetIndex,1,1,MAX_SPELL_COUNT);
        const rule = SpreadsheetApp.newDataValidation() 
          .requireValueInRange(validationRange) // warning: possibly expensive
          .setAllowInvalid(true)
          .build();
        rules[j] = [ rule ];
        // const critVals = rule.getCriteriaValues();
        // associate this validationSheetIndex with this spellRange in case some mongo c&ps
        //Logger.log(`adding val index ${validationSheetIndex} to rules[${j}] and mapped to ${spellCellA1Notation}. range criteria=${critVals[0].getA1Notation()}`);
        spellCellsToValidation[spellCellA1Notation] = validationSheetIndex;
        validationSheetIndex++;
      } else {
        const validationCell1 = validationSheet.getRange(validationSheetIndex, 1);
        validationCell1.setValue("this cell intentionally left blank"); 
        validationSheetIndex++;
        //Logger.log(`adding null rule to rules[${j}]`);
        rules[j] = [ null ];
      }
    }
    spellRange.setDataValidations(rules);
  }

  updateValidationIndexMap(sheetName, spellCellsToValidation);
}

function updateValidationIndexMap(sheetName, spellCellsToValidation)
{
  // stores the 2D array mapping spell cell coordinates to the Validations sheet row index
  const key = sheetName + " Validations"; // e.g. MRT CD Planner Validations
  const mapStr = JSON.stringify(spellCellsToValidation);
  //Logger.log(`caching: ${mapStr}`);
  const cache = CacheService.getScriptCache();
  cache.put(key, mapStr, 3600);
}

function getValidationIndex(spellColA1Notation, sheetName) 
{
  // reads the 2D array mapping spell cell coordinates to the Validations sheet row index
  const cache = CacheService.getScriptCache();
  const key = sheetName + " Validations"; // e.g. MRT CD Planner Validations
  var validationMapStr = cache.get(key);
  if (!validationMapStr) {
    // toast? should never get here
    return 0;
  }
  const validationMap = JSON.parse(validationMapStr);  
  if (!validationMap[spellColA1Notation]) {
    //Logger.log(`getValidationIndex: could not find entry for cell ${spellColA1Notation} in map: ${validationMapStr}`)
    return 0;
  }
  //Logger.log(`getValidationIndex: returning ${validationMap[spellColA1Notation]} for cell ${spellColA1Notation} in map: ${validationMapStr}`)
  return validationMap[spellColA1Notation];
}