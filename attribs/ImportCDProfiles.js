function showDecodeModal() {
  const template = HtmlService.createTemplateFromFile("ImportCDProfilesPopup");
  const html = template.evaluate().setWidth(550).setHeight(420);
  SpreadsheetApp.getUi().showModalDialog(html, 'Import CD Planner Profile');
}

function importWithToasts(inputString) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast('Importing Fojji Profile…', '⏳ Please Wait...', 10);
  const keyMap = getLookupMap_("KEY_LOOKUP");
  const spellMap = getLookupMap_("SPELL_LOOKUP");
  const custIconMap = getLookupMap_("CUSTOM_ICON_LOOKUP");
  const { importedCount, skippedCount } = decodeAndPasteMultipleRanges_(inputString, keyMap, spellMap, custIconMap);
  ss.toast(`Imported ${importedCount} rows, skipped ${skippedCount} rows.`, '⚔️ Raid CD Planner', 5);
}

function decodeAndPasteFromInput(inputString) {
  const keyMap = getLookupMap_("KEY_LOOKUP");
  const spellMap = getLookupMap_("SPELL_LOOKUP");
  const custIconMap = getLookupMap_("CUSTOM_ICON_LOOKUP");
  decodeAndPasteMultipleRanges_(inputString, keyMap, spellMap, custIconMap);
}

function getLookupMap_(namedRange) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const range = ss.getRangeByName(namedRange);
  if (!range) throw new Error(`Named range "${namedRange}" not found.`);
  const values = range.getValues();
  return new Map(values.map(([k, v]) => [String(k).trim(), String(v).trim()]));
}

function invertMap_(map) {
  const inv = new Map();
  for (const [k, v] of map.entries()) {
    if (v !== undefined && v !== null && v !== "") inv.set(String(v).trim(), String(k).trim());
  }
  return inv;
}

function decodeAndPasteMultipleRanges_(formattedString, keyMap, spellMap, custIconMap) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const bossAliases = { ELDERS: "COUNCIL", IRONQON: "QON", CONSORTS: "TWINS" };
  const namedRanges = ss.getNamedRanges();
  const rangeByName = {};
  for (let i = 0; i < namedRanges.length; i++) {
    rangeByName[namedRanges[i].getName().toUpperCase()] = namedRanges[i].getRange();
  }

  const invertedKeyMap = invertMap_(keyMap);
  const invertedSpellMap = invertMap_(spellMap);
  const invertedCustIconMap = invertMap_(custIconMap);
  if (!formattedString) return { importedCount: 0, skippedCount: 0 };

  const rowStrs = formattedString.split("*").filter(s => s !== "");
  const bossesInData = new Set();
  for (let i = 0; i < rowStrs.length; i++) {
    const key = String(rowStrs[i].split("/")[0] || "").toUpperCase();
    const parts = key.split("_");
    if (parts.length > 1) bossesInData.add(parts[0]);
  }

  const bossRanges = {};
  bossesInData.forEach(boss => {
    const bossUpper = boss.toUpperCase();
    const healthKey = bossUpper + "_HEALTH_AREA";
    const timeKey = bossUpper + "_TIME_AREA";
    if (rangeByName[healthKey]) {
      bossRanges[bossUpper] = bossRanges[bossUpper] || {};
      bossRanges[bossUpper].health = healthKey;
    }
    if (rangeByName[timeKey]) {
      bossRanges[bossUpper] = bossRanges[bossUpper] || {};
      bossRanges[bossUpper].time = timeKey;
    }
  });

  const buckets = {};
  Object.values(bossRanges).forEach(b => {
    if (b.health) buckets[b.health] = [];
    if (b.time) buckets[b.time] = [];
  });

  const COLS = { KEY: 0, COUNT: 3, NAME: 5, TIME: 4, SPELL: 9, NPC: 1, OPT_START: 10, OPT_END: 12, CUST_ICON: 14 };
  const skippedCounter = { count: 0 };

  for (let i = 0; i < rowStrs.length; i++) {
    const parts = rowStrs[i].split("/");
    const encodedKeyUpper = String(parts[0] || "").toUpperCase();
    let boss = null;
    for (const [alias, canonical] of Object.entries(bossAliases)) {
      if (encodedKeyUpper.startsWith(alias + "_")) { boss = canonical; break; }
    }
    if (!boss) {
      for (const b of Object.keys(bossRanges)) {
        if (encodedKeyUpper.startsWith(b + "_")) { boss = b; break; }
      }
    }
    if (!boss) { skippedCounter.count++; continue; }
    const lastToken = encodedKeyUpper.split("_").pop();
    const isHealth = lastToken === "HEALTH";
    const targetRangeName = isHealth ? bossRanges[boss]?.health : bossRanges[boss]?.time;
    if (!targetRangeName) { skippedCounter.count++; continue; }
    const row = decodeRowAligned_(parts, isHealth, COLS, invertedKeyMap, invertedSpellMap, invertedCustIconMap, skippedCounter);
    if (row) buckets[targetRangeName].push(row);
  }

  let importedCount = 0;

  for (const [rangeName, rows] of Object.entries(buckets)) {
    if (!rows.length) continue;
    const range = rangeByName[rangeName.toUpperCase()];
    if (!range) { skippedCounter.count += rows.length; continue; }

    const neededCols = rows[0].length;
    const neededRows = rows.length;
    const totalRows = range.getNumRows();
    const totalCols = range.getNumColumns();
    if (totalCols < neededCols) { skippedCounter.count += rows.length; continue; }

    const targetRange = range.offset(0, 0, totalRows, neededCols);
    const formulas = targetRange.getFormulas();
    const values = targetRange.getValues();

    for (let r = 0; r < totalRows; r++) {
      for (let c = 0; c < neededCols; c++) {
        const f = formulas[r][c];
        if (f) {
          values[r][c] = f;
        } else if (r < neededRows) {
          const v = rows[r][c];
          values[r][c] = v !== undefined && v !== null ? v : "";
        } else {
          values[r][c] = "";
        }
      }
    }

    targetRange.setValues(values);
    importedCount += rows.length;
  }

  return { importedCount, skippedCount: skippedCounter.count };
}

function decodeRowAligned_(parts, isHealth, COLS, invertedKeyMap, invertedSpellMap, invertedCustIconMap) {
  const maxIndex = Math.max(COLS.OPT_END, COLS.CUST_ICON);
  const row = new Array(maxIndex + 1).fill("");
  const trimmedParts = parts.map(p => (p || "").toString().trim());
  row[COLS.KEY] = trimmedParts[0];
  row[COLS.COUNT] = trimmedParts[1];
  row[COLS.NAME] = trimmedParts[2];
  row[COLS.TIME] = trimmedParts[3];
  row[COLS.SPELL] = trimmedParts[4];
  if (isHealth) {
    row[COLS.NPC] = trimmedParts[5];
    row[COLS.CUST_ICON] = trimmedParts[9];
    const opts = trimmedParts.slice(6, 9);
    for (let i = 0; i < opts.length; i++) row[COLS.OPT_START + i] = (!opts[i] || opts[i].toLowerCase() === "nil") ? "" : opts[i];
  } else {
    row[COLS.NPC] = "";
    row[COLS.CUST_ICON] = trimmedParts[8];
    const opts = trimmedParts.slice(5, 8);
    for (let i = 0; i < opts.length; i++) row[COLS.OPT_START + i] = (!opts[i] || opts[i].toLowerCase() === "nil") ? "" : opts[i];
  }
  if (row[COLS.COUNT] === "-1") row[COLS.COUNT] = "ALL";
  if (invertedKeyMap.has(row[COLS.KEY])) row[COLS.KEY] = invertedKeyMap.get(row[COLS.KEY]);
  if (!row[COLS.SPELL] || row[COLS.SPELL].toLowerCase() === "nil") row[COLS.SPELL] = "Custom Spell Assignment";
  else if (invertedSpellMap.has(row[COLS.SPELL])) row[COLS.SPELL] = invertedSpellMap.get(row[COLS.SPELL]);
  const rawKey = row[COLS.CUST_ICON];
  const lookupKey = String(rawKey).trim();
  if (invertedCustIconMap.has(lookupKey)) row[COLS.CUST_ICON] = String(invertedCustIconMap.get(lookupKey)).trim();
  else {
    const numericKey = Number(rawKey);
    if (!isNaN(numericKey) && invertedCustIconMap.has(numericKey)) row[COLS.CUST_ICON] = String(invertedCustIconMap.get(numericKey)).trim();
    else row[COLS.CUST_ICON] = "";
  }
  return row.map(v => (typeof v === "string" ? v.trim() : v));
}
