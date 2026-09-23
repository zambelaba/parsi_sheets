/** Generate ratio for import */
function writeJsonToCell() {
  writeRatioForRCLCInAdmin();
  writeRatioForGargulInAdmin();
}

function writeRatioForRCLCInAdmin() {
  const watchedSheetName = "Ratio Présence/Loot (Préloot)";
  const targetSheetName = "Admin";
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const watchedSheet = spreadsheet.getSheetByName(watchedSheetName);
  const data = watchedSheet.getDataRange().getValues();
  let obj = {};

  // Loop through rows, skipping the header row
  for (let i = 2; i < data.length; i++) {
    const key = data[i][0];    // Column A (index 0)
    const value = (Number(data[i][9]) || 0).toFixed(3); // Column J (index 9)

    Logger.log(`Ligne ${i+1} | key="${key}" | value="${value}"`);

    // Make sure the key exists, but allow value = 0
    if (key !== "" && key !== null && value !== "" && value !== null) {
      obj[key] = value;
    }
  }

  const jsonString = JSON.stringify(obj, null, 2);
  
  // Create blob and base64 encode directly
  const blob = Utilities.newBlob(jsonString, 'text/plain');
  const base64String = Utilities.base64Encode(blob.getBytes());

  // Write the JSON into cell U3
  Logger.log("Base64 généré : " + base64String);
  const targetSheet = spreadsheet.getSheetByName(targetSheetName);
  targetSheet.getRange("C2").setValue(base64String);
  Logger.log("Écriture RCLC effectuée");
}

/** Generate Gargul ratio for import */
function writeRatioForGargulInAdmin() {
  const watchedSheetName = "Ratio Présence/Loot (Préloot)";
  const targetSheetName = "Admin";
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const watchedSheet = spreadsheet.getSheetByName(watchedSheetName);
  const data = watchedSheet.getDataRange().getValues();
  let obj = {};

  // Loop through rows, skipping the header row
  for (let i = 2; i < data.length; i++) {
    const key = data[i][0]; // Column A (index 0)
    const attendanceValue = (Number(data[i][6]) || 0) + (Number(data[i][7]) || 0); // Columns G + H
    const lootCountValue = Number(data[i][9]) || 0; // Column J (index 9)

    Logger.log(`Ligne ${i+1} | key="${key}" | attendance="${attendanceValue}" | lootCount="${lootCountValue}"`);

    // Make sure the key exists, but allow values = 0
    if (key !== "" && key !== null) {
      obj[key] = [attendanceValue, lootCountValue];
    }
  }

  const jsonString = JSON.stringify(obj, null, 2);
  
  // Create blob and base64 encode directly
  const blob = Utilities.newBlob(jsonString, 'text/plain');
  const base64String = Utilities.base64Encode(blob.getBytes());

  // Write the JSON into cell U3
  Logger.log("Base64 généré : " + base64String);
  const targetSheet = spreadsheet.getSheetByName(targetSheetName);
  targetSheet.getRange("D2").setValue(base64String);
  Logger.log("Écriture Gargul effectuée");
}

function resetAdminImportRatioCell() {
  const targetSheetName = "Admin";

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const targetSheet = spreadsheet.getSheetByName(targetSheetName);

  // Reset the cell
  targetSheet.getRange("C2").clearContent();
  targetSheet.getRange("D2").clearContent();
  targetSheet.getRange("E2").clearContent();
}

/** Process raid logs sheets export */
function process_raid_logs_loots() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rclcResult = process_rclc_raid_logs_loots({ skipOutput: true });
  const gargulResult = process_gargul_raid_logs_loots({ skipOutput: true });
  const allItems = rclcResult.allItems.concat(gargulResult.allItems);
  const playerSet = new Set();

  for (let i = 0; i < allItems.length; i++) {
    playerSet.add(allItems[i].player);
  }

  populateLootDetails(ss, allItems, Array.from(playerSet).sort());
  updateLootCounts(ss, allItems);

  Logger.log(
    'Processed ' + allItems.length +
    ' total items from RCLC and Gargul. RCLC: ' + rclcResult.allItems.length +
    ', Gargul: ' + gargulResult.allItems.length + '.'
  );

  resetAdminImportRatioCell();

  return { allItems: allItems, playerSet: playerSet };
}

function process_rclc_raid_logs_loots(options) {
  options = options || {};
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('raid_logs');
  
  if (!sheet) {
    Logger.log('raid_logs sheet not found.');
    return { allItems: [], playerSet: new Set() };
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    Logger.log('No data rows found.');
    return { allItems: [], playerSet: new Set() };
  }

  const colD = 4; // column D (RCLC JSON payload)

  // Load prio item IDs (items with coeff 0)
  const prioIds = getPrioItemIds();
  const rosterPlayers = getRatioPresenceRoster(ss);
  if (rosterPlayers === null) {
    Logger.log('Skipping RCLC loot processing: roster sheet not found.');
    return { allItems: [], playerSet: new Set() };
  }
  if (rosterPlayers === null) {
    Logger.log('Skipping loot processing: Ratio Présence roster sheet not found.');
    return;
  }

  // Read all rows from row 2 to lastRow
  const dataRange = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn());
  const rows = dataRange.getValues();

  const allItems = []; // array of {date, player, itemID, itemName} objects
  const playerSet = new Set(); // to collect unique player names
  let skippedNonRosterItems = 0;

  for (let i = 0; i < rows.length; i++) {
    const rowIndex = i + 2; // actual sheet row
    const payloadCell = rows[i][colD - 1];
    if (!payloadCell) {
      // nothing to parse
      continue;
    }

    let parsed = null;
    try {
      parsed = JSON.parse(payloadCell);
    } catch (e) {
      try {
        // fallback: payload may be a comma-separated list of objects without
        // the wrapping array brackets
        parsed = JSON.parse('[' + String(payloadCell).replace(/,\s*$/, '') + ']');
      } catch (e2) {
        // unable to parse, skip
        Logger.log('Skipping row ' + rowIndex + ': invalid JSON');
        continue;
      }
    }

    if (!parsed) continue;

    // Handle both array and object
    const records = Array.isArray(parsed) ? parsed : [parsed];

    for (let j = 0; j < records.length; j++) {
      const record = records[j];
      
      if (!shouldProcessRecord(record)) {
        continue;
      }

      const playerKey = getRosterPlayerKey(record.player || record.owner || '');
      if (!rosterPlayers.has(playerKey)) {
        skippedNonRosterItems++;
        continue;
      }

      const player = rosterPlayers.get(playerKey);
      
      const itemName = record.itemName || record.item || '';
      const itemID = record.itemID || '';
      const date = formatDate(record.date || '');

      // determine coefficient based on instance
      let coeff = 1;
      if (record.instance && typeof record.instance === 'string') {
        if (record.instance.includes('Kara') ||
            record.instance.includes('Gruul') ||
            record.instance.includes('Magh')
          ) {
          const recordDate = new Date(record.date);
          const p2StartDate = new Date('2026-05-12');

          if (recordDate > p2StartDate) {
            coeff = 0.33;
          } else {
            coeff = 1;
          }
        }
        else{
          coeff = 1
        }
      }

      // Check if item is in prio list (coefficient 0)
      if (prioIds.has(String(itemID))) {
        coeff = 0;
      }

      allItems.push({
        date: date,
        player: player,
        itemID: itemID,
        itemName: itemName,
        coeff: coeff
      });

      playerSet.add(player);
    }
  }

  if (options.skipOutput) {
    return { allItems: allItems, playerSet: playerSet };
  }

  // Populate loot_details sheet
  populateLootDetails(ss, allItems, Array.from(playerSet).sort());

  // Update loot counts in ratio sheet
  updateLootCounts(ss, allItems);

  Logger.log('Processed ' + allItems.length + ' items. Skipped ' + skippedNonRosterItems + ' non-roster items.');

  resetAdminImportRatioCell();

  return { allItems: allItems, playerSet: playerSet };
}

function process_gargul_raid_logs_loots(options) {
  options = options || {};
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('raid_logs');
  
  if (!sheet) {
    Logger.log('raid_logs sheet not found.');
    return { allItems: [], playerSet: new Set() };
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    Logger.log('No data rows found.');
    return { allItems: [], playerSet: new Set() };
  }

  const colC = 3; // column C (Gargul JSON payload)

  // Load prio item IDs (items with coeff 0)
  const prioIds = getPrioItemIds();
  const rosterPlayers = getRatioPresenceRoster(ss);
  if (rosterPlayers === null) {
    Logger.log('Skipping Gargul loot processing: roster sheet not found.');
    return { allItems: [], playerSet: new Set() };
  }
  if (rosterPlayers === null) {
    Logger.log('Skipping RCLC loot processing: roster sheet not found.');
    return { allItems: [], playerSet: new Set() };
  }
  if (rosterPlayers === null) {
    Logger.log('Skipping Gargul loot processing: Ratio Présence roster sheet not found.');
    return;
  }

  if (rosterPlayers === null) {
    Logger.log('Skipping loot processing: Ratio Présence roster sheet not found.');
    return { allItems: [], playerSet: new Set() };
  }

  // Read all rows from row 2 to lastRow
  const dataRange = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn());
  const rows = dataRange.getValues();

  const allItems = []; // array of {date, player, itemID, itemName, coeff} objects
  const playerSet = new Set(); // to collect unique player names
  const warnings = [];
  let skippedNonRosterItems = 0;
  let skippedMissingRequiredFields = 0;

  for (let i = 0; i < rows.length; i++) {
    const rowIndex = i + 2; // actual sheet row
    const payloadCell = rows[i][colC - 1];
    if (!payloadCell) {
      // nothing to parse
      continue;
    }

    const records = parseGargulRecords(payloadCell, rowIndex, warnings);

    for (let j = 0; j < records.length; j++) {
      const record = records[j];
      const recordLabel = 'row ' + rowIndex + ', record ' + (j + 1);
      
      if (!shouldProcessGargulRecord(record)) {
        continue;
      }

      const playerName = record.player || '';
      const itemName = record.itemName || record.item || '';
      const itemID = record.id || record.itemID || '';
      const date = formatDate(record.date || '');
      const coeff = getGargulRecordCoeff(record, recordLabel, warnings);

      if (!playerName) {
        warnings.push('Missing player on ' + recordLabel + '. Skipping record.');
      }
      if (!itemID) {
        warnings.push('Missing id/itemID on ' + recordLabel + '. Skipping record.');
      }
      if (!itemName) {
        warnings.push('Missing itemName/item on ' + recordLabel + '. Skipping record.');
      }
      if (!record.date) {
        warnings.push('Missing date on ' + recordLabel + '. Keeping record with an empty date.');
      }

      if (!playerName || !itemID || !itemName) {
        skippedMissingRequiredFields++;
        continue;
      }

      const playerKey = getRosterPlayerKey(playerName);
      if (!rosterPlayers.has(playerKey)) {
        skippedNonRosterItems++;
        continue;
      }

      const player = rosterPlayers.get(playerKey);
      const finalCoeff = prioIds.has(String(itemID)) ? 0 : coeff;

      allItems.push({
        date: date,
        player: player,
        itemID: itemID,
        itemName: itemName,
        coeff: finalCoeff
      });

      playerSet.add(player);
    }
  }

  logGargulProcessingWarnings(warnings);

  if (options.skipOutput) {
    return { allItems: allItems, playerSet: playerSet };
  }

  // Populate loot_details sheet
  populateLootDetails(ss, allItems, Array.from(playerSet).sort());

  // Update loot counts in ratio sheet
  updateLootCounts(ss, allItems);

  Logger.log(
    'Processed ' + allItems.length +
    ' Gargul items. Skipped ' + skippedNonRosterItems +
    ' non-roster items and ' + skippedMissingRequiredFields +
    ' items with missing required fields.'
  );

  resetAdminImportRatioCell();

  return { allItems: allItems, playerSet: playerSet };
}

function getRatioPresenceSheet(ss) {
  const sheet = ss.getSheetByName('Ratio Présence/Loot (Préloot)');
  if (!sheet) {
    Logger.log('Ratio Présence/Loot (Préloot) sheet not found.');
  }

  return sheet;
}

function normalizePlayerName(playerName) {
  const player = String(playerName || '')
    .trim()
    .replace(/-Thunderstrike$/i, '');

  if (!player) {
    return '';
  }

  return player.charAt(0).toUpperCase() + player.slice(1);
}

function getRosterPlayerKey(playerName) {
  return normalizePlayerName(playerName).toLowerCase();
}

function getRatioPresenceRoster(ss) {
  const ratioSheet = getRatioPresenceSheet(ss);

  if (!ratioSheet) {
    return null;
  }

  const lastRow = ratioSheet.getLastRow();
  if (lastRow < 3) {
    Logger.log('No player rows found in ratio sheet roster (need at least row 3).');
    return new Map();
  }

  const playerRange = ratioSheet.getRange(3, 1, lastRow - 2, 1);
  const playerNames = playerRange.getValues();
  const rosterPlayers = new Map();

  for (let i = 0; i < playerNames.length; i++) {
    const playerName = normalizePlayerName(playerNames[i][0]);
    const playerKey = getRosterPlayerKey(playerName);

    if (playerName && !rosterPlayers.has(playerKey)) {
      rosterPlayers.set(playerKey, playerName);
    }
  }

  Logger.log('Loaded ' + rosterPlayers.size + ' roster players from Ratio Présence.');
  return rosterPlayers;
}

/** Get all prio item IDs from the "Prio items" sheet (column X) */
function getPrioItemIds() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const prioSheet = ss.getSheetByName('Prio items');
  
  if (!prioSheet) {
    Logger.log('Prio items sheet not found.');
    return new Set();
  }

  const colX = 24; // column X
  const lastRow = prioSheet.getLastRow();
  
  if (lastRow < 2) {
    return new Set();
  }

  // Read column X from row 2 onwards
  const range = prioSheet.getRange(2, colX, lastRow - 1, 1);
  const values = range.getValues();
  
  const prioIds = new Set();
  for (let i = 0; i < values.length; i++) {
    const itemId = values[i][0];
    if (itemId && itemId !== '') {
      prioIds.add(String(itemId)); // convert to string for comparison
    }
  }
  
  Logger.log('Loaded ' + prioIds.size + ' prio item IDs');
  return prioIds;
}

/** Check if a record should be processed */
function shouldProcessRecord(record) {
  // Must have responseID = 1
  if (String(record.responseID) !== '1') {
    return false;
  }
  
  // Skip disenchant responses
  if (record.response === 'Désenchantement' || record.response === 'Disenchant') {
    return false;
  }
  
  // Skip patterns
  const itemName = record.itemName || record.item || '';
  if (itemName.startsWith('Pattern') || itemName.startsWith('Design')) {
    return false;
  }
  
  return true;
}

function parseGargulRecords(payload, rowIndex, warnings) {
  const text = String(payload || '').trim();
  if (!text) {
    return [];
  }

  let parsed = null;

  try {
    parsed = JSON.parse(text);
  } catch (e) {
    try {
      const arrayText = '[' + text.replace(/,\s*$/, '') + ']';
      parsed = JSON.parse(arrayText);
    } catch (e2) {
      warnings.push('Skipping row ' + rowIndex + ': invalid Gargul JSON export.');
      return [];
    }
  }

  return Array.isArray(parsed) ? parsed : [parsed];
}

function shouldProcessGargulRecord(record) {
  if (String(record.OS) === '1') {
    return false;
  }

  const itemName = record.itemName || record.item || '';
  if (itemName.startsWith('Pattern') || itemName.startsWith('Design')) {
    return false;
  }
  
  return true;
}

function getGargulRecordCoeff(record, recordLabel, warnings) {
  if (record.coeff === '' || record.coeff === null || record.coeff === undefined) {
    warnings.push('Missing coeff on ' + recordLabel + '. Defaulting to 1.');
    return 1;
  }

  const coeff = Number(String(record.coeff).replace(',', '.'));
  if (isNaN(coeff)) {
    warnings.push('Invalid coeff "' + record.coeff + '" on ' + recordLabel + '. Defaulting to 1.');
    return 1;
  }

  return coeff;
}

function logGargulProcessingWarnings(warnings) {
  if (warnings.length === 0) {
    Logger.log('No Gargul processing warnings.');
    return;
  }

  Logger.log('Gargul processing warnings (' + warnings.length + '):');

  const maxWarningsToLog = 50;
  for (let i = 0; i < warnings.length && i < maxWarningsToLog; i++) {
    Logger.log('- ' + warnings[i]);
  }

  if (warnings.length > maxWarningsToLog) {
    Logger.log('- ... ' + (warnings.length - maxWarningsToLog) + ' more warnings not shown.');
  }
}

/** Format date to dd/MM/yyyy format */
function formatDate(dateStr) {
  if (!dateStr) return '';
  
  try {
    // Handle yyyy/MM/dd and yyyy-MM-dd formats
    const parts = String(dateStr).split(/[\/-]/);
    if (parts.length === 3) {
      const year = parts[0];
      const month = parts[1];
      const day = parts[2];
      return day + '/' + month + '/' + year;
    }
  } catch (e) {
    Logger.log('Error formatting date: ' + dateStr);
  }
  
  return dateStr;
}

function populateLootDetails(ss, allItems, sortedPlayers) {
  // Get or create loot_details sheet
  let detailsSheet = ss.getSheetByName('loot_details');
  if (!detailsSheet) {
    detailsSheet = ss.insertSheet('loot_details');
  } else {
    // Clear existing content and remove any merges/formats
    detailsSheet.clear();
  }

  if (sortedPlayers.length === 0 || allItems.length === 0) {
    Logger.log('No players or items to populate.');
    return;
  }

  // Row 1: Player names (merged across 3 columns each)
  // Row 2: Fixed headers ('raid_date', 'item_name', 'item_id')
  // Row 3+: One row per item

  // Each player takes 4 columns now (raid_date, item_name, item_id, coeff)
  const numCols = sortedPlayers.length * 4;

  // Write row 1: player names with 4-column merge
  const row1Values = [];
  for (let i = 0; i < sortedPlayers.length; i++) {
    row1Values.push(sortedPlayers[i]);
    row1Values.push('');
    row1Values.push('');
    row1Values.push('');
  }
  detailsSheet.getRange(1, 1, 1, numCols).setValues([row1Values]);

  // Merge cells for player names (4 cells per player)
  for (let i = 0; i < sortedPlayers.length; i++) {
    const startCol = i * 4 + 1;
    detailsSheet.getRange(1, startCol, 1, 4).merge();
  }

  // Write row 2: fixed headers
  const row2Values = [];
  for (let i = 0; i < sortedPlayers.length; i++) {
    row2Values.push('raid_date');
    row2Values.push('item_name');
    row2Values.push('item_id');
    row2Values.push('coeff');
  }
  detailsSheet.getRange(2, 1, 1, numCols).setValues([row2Values]);

  // Style headers (rows 1-2): bold and light grey background
  detailsSheet.getRange(1, 1, 2, numCols)
    .setFontWeight('bold')
    .setBackground('#073763') // dark blue background for better contrast
    .setFontColor('#FFFFFF') // white font color for better contrast
    .setHorizontalAlignment('center');
  // Set player name row1 font size 18, rest font 10
  detailsSheet.getRange(1, 1, 1, numCols).setFontSize(18);
  detailsSheet.getRange(2, 1, 1, numCols).setFontSize(10);

  // Optionally adjust column widths for readability
  // raid_date narrower, item_name wide, item_id normal, coeff smaller
  for (let p = 0; p < sortedPlayers.length; p++) {
    const base = p * 4;
    detailsSheet.setColumnWidth(base + 1, 75); // raid_date
    detailsSheet.setColumnWidth(base + 2, 200); // item_name
    detailsSheet.setColumnWidth(base + 3, 50);  // item_id
    detailsSheet.setColumnWidth(base + 4, 35);  // coeff
  }

  // Build row data (starting from row 3)
  const rowData = [];
  for (let itemIdx = 0; itemIdx < allItems.length; itemIdx++) {
    const item = allItems[itemIdx];
    const rowValues = [];

    for (let playerIdx = 0; playerIdx < sortedPlayers.length; playerIdx++) {
      const playerName = sortedPlayers[playerIdx];

      // Find all items for this player
      const playerItems = allItems.filter(
        x => x.player === playerName && x.itemID !== '' && x.itemName !== ''
      );

      // Get the item for this column quadruplet (if exists)
      if (itemIdx < playerItems.length) {
        const playerItem = playerItems[itemIdx];
        rowValues.push(playerItem.date || '');
        rowValues.push(playerItem.itemName || '');
        rowValues.push(playerItem.itemID || '');
        rowValues.push(playerItem.coeff != null ? playerItem.coeff : '');
      } else {
        rowValues.push('');
        rowValues.push('');
        rowValues.push('');
        rowValues.push('');
      }
    }

    rowData.push(rowValues);
  }

  // Write all data rows
  if (rowData.length > 0) {
    detailsSheet.getRange(3, 1, rowData.length, numCols).setValues(rowData);

    // Apply formatting: raid_date as plain text, item_id and coeff as numbers
    for (let p = 0; p < sortedPlayers.length; p++) {
      const raidDateCol = p * 4 + 1;
      const itemNameCol = p * 4 + 2;
      const itemIdCol = p * 4 + 3;
      const coeffCol = p * 4 + 4;
      // set raid_date column as text to prevent automatic date parsing
      detailsSheet.getRange(3, raidDateCol, rowData.length, 1).setNumberFormat('@');
      // set item_id and coeff columns as integer
      detailsSheet.getRange(3, itemIdCol, rowData.length, 1).setNumberFormat('0');
      detailsSheet.getRange(3, coeffCol, rowData.length, 1).setNumberFormat('0.00');
      // add vertical separator right border after each player block
      const sepCol = coeffCol;
      detailsSheet.getRange(1, sepCol, rowData.length + 2, 1)
        // apply right border (after coeff column) instead of left
        .setBorder(null, null, null, true, null, null, 'black', SpreadsheetApp.BorderStyle.SOLID);
    }
  }

  Logger.log('Loot details sheet populated with ' + rowData.length + ' rows.');
}

/** Update loot counts in the Ratio Présence/Loot (Préloot) sheet */
function updateLootCounts(ss, allItems) {
  const ratioSheet = getRatioPresenceSheet(ss);
  
  if (!ratioSheet) {
    return;
  }

  // Create a map of player -> weighted loot count (coeff applied)
  const lootCountMap = {};
  for (let i = 0; i < allItems.length; i++) {
    const player = allItems[i].player;
    const coeff = allItems[i].coeff != null ? allItems[i].coeff : 1;
    lootCountMap[player] = (lootCountMap[player] || 0) + coeff;
  }

  // Get all data from the ratio sheet
  const lastRow = ratioSheet.getLastRow();
  if (lastRow < 3) {
    Logger.log('No player rows found in ratio sheet (need at least row 3).');
    return;
  }

  // Read player names from column A starting at row 3 (skip rows 1 and 2)
  const playerRange = ratioSheet.getRange(3, 1, lastRow - 2, 1);
  const playerNames = playerRange.getValues();

  // Update column J with loot counts
  for (let i = 0; i < playerNames.length; i++) {
    const playerName = playerNames[i][0];
    if (playerName && playerName.trim() !== '') {
      const lootCount = lootCountMap[playerName] || 0;
      const rowIndex = i + 3; // actual sheet row (starting from row 3)
      ratioSheet.getRange(rowIndex, 10).setValue(lootCount); // Column J = column 10
    }
  }

  Logger.log('Updated loot counts in ratio sheet.');
}
