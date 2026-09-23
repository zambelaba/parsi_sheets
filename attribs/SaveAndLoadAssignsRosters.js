const ENABLE_LOG = true
const RANGE_A1 = 'AN6:CI38'
const ROSTER_RANGE = 'AB5:AF64'
const ROSTER_EXTRA_RANGE = 'AH13:BB22'
const ROSTER_EXTRA_ALLOWED = ['AM15:AM18','AX15:AX22']
const TARGET_SHEETS = [
  'Roster', 'Group Builder (25)', 'Group Builder (10)', 'Maulgar','Gruul','Magtheridon', 'Hydross', 'Lurker', 'Leotheras', 'Karathress', 'Morogrim', 'Vashj[Alt]', 'Vashj', 'Void Reaver', 'Al\'ar', 'Solarian', 'KT'
]
const SAVE_SHEET_NAME = 'Save Data'
const SAVE_LABELS = ['Split 1','Split 2','Split 3','Split 4','Split 5','Custom']

function showSaveModal() {
  const slots = getSaveSlots()
  const template = HtmlService.createTemplateFromFile('SaveAssignmentsPopup')
  template.slotsData = slots
  const html = template.evaluate().setWidth(500).setHeight(300)
  SpreadsheetApp.getUi().showModalDialog(html, 'Save Assignments')
}

function showLoadModal() {
  const saveSlots = getSaveSlots().filter(s => s.hasData)
  const targetSheets = getTargetSheets().filter(n => n !== 'Group Builder (25)')
  const template = HtmlService.createTemplateFromFile("LoadAssignmentsPopup")
  template.slotsData = saveSlots
  template.sheetNames = targetSheets
  const html = template.evaluate().setWidth(550).setHeight(500)
  SpreadsheetApp.getUi().showModalDialog(html, "Load Saved Assignments")
}

function toastMessage(title, msg, secs) {
  SpreadsheetApp.getActiveSpreadsheet().toast(msg, title, secs)
}

function compress(str) {
  const bytes = Utilities.newBlob(str).getBytes()
  const gzipped = Utilities.gzip(Utilities.newBlob(bytes))
  return Utilities.base64Encode(gzipped.getBytes())
}

function decompress(base64Str) {
  const bytes = Utilities.base64Decode(base64Str)
  const blob = Utilities.newBlob(bytes, 'application/x-gzip')
  const decompressed = Utilities.ungzip(blob)
  return decompressed.getDataAsString()
}

function getSaveSlots() {
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const saveSheet = ss.getSheetByName(SAVE_SHEET_NAME)
  if (!saveSheet) return []
  const slots = []
  for (let i=1; i<=6; i++) {
    const val = saveSheet.getRange(`A${i}`).getValue()
    slots.push({id:i, label:SAVE_LABELS[i-1], hasData:!!val})
  }
  return slots
}

function getTargetSheets() {
  return TARGET_SHEETS
}

function refreshEditableRegions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const map = {}
  TARGET_SHEETS.forEach(name => {
    const sheet = ss.getSheetByName(name)
    if (!sheet) return
    const range = sheet.getDataRange()
    const validations = range.getDataValidations()
    const editable = []
    for (let r = 0; r < validations.length; r++) {
      for (let c = 0; c < validations[0].length; c++) {
        if (validations[r][c]) editable.push({ row: r + 1, col: c + 1 })
      }
    }
    map[name] = editable
  })
  const json = JSON.stringify(map)
  const compressed = compress(json)
  const saveSheet = ss.getSheetByName(SAVE_SHEET_NAME)
  saveSheet.getRange('C1').setValue(compressed)
  toastMessage('📋 Assignments Helper', 'Editable regions refreshed', 4)
}

function getEditableRegions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const saveSheet = ss.getSheetByName(SAVE_SHEET_NAME)
  const compressed = saveSheet.getRange('C1').getValue()
  return compressed ? JSON.parse(decompress(compressed)) : {}
}

function saveAssignmentsToSlot(slotId) {
  const label = SAVE_LABELS[slotId-1] || `Slot ${slotId}`
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const saveSheet = ss.getSheetByName(SAVE_SHEET_NAME)
  const saveCell = `A${slotId}`
  const editableMap = getEditableRegions()
  toastMessage('📋 Assignments Helper', `Saving Assignments to ${label}`, 6)
  const startAll = Date.now()
  const data = {}
  const summary = []

  TARGET_SHEETS.forEach(name => {
    if (name === 'Group Builder (25)') return
    const sheetStart = Date.now()
    const ssheet = ss.getSheetByName(name)
    if (!ssheet) return

    let cells = []

    if (name === 'Roster') {
      const mainRange = ssheet.getRange(ROSTER_RANGE)
      const mainValues = mainRange.getValues()
      const mainFormulas = mainRange.getFormulas()

      for (let r = 0; r < mainValues.length; r++) {
        for (let c = 0; c < mainValues[0].length; c++) {
          const globalRow = mainRange.getRow() + r
          const globalCol = mainRange.getColumn() + c
          const val = mainValues[r][c]
          const f = mainFormulas[r][c]
          if (typeof val === 'object' && val && val.valueType === 'IMAGE') continue
          if (val !== '' || f) {
            cells.push({
              row: globalRow,
              col: globalCol,
              value: val,
              hasFormula: !!f
            })
          }
        }
      }

      ROSTER_EXTRA_ALLOWED.forEach(a1 => {
        const range = ssheet.getRange(a1)
        const values = range.getValues()
        const formulas = range.getFormulas()
        for (let r = 0; r < values.length; r++) {
          for (let c = 0; c < values[0].length; c++) {
            const globalRow = range.getRow() + r
            const globalCol = range.getColumn() + c
            const val = values[r][c]
            const f = formulas[r][c]
            if (typeof val === 'object' && val && val.valueType === 'IMAGE') continue
            if (val !== '' || f) {
              cells.push({
                row: globalRow,
                col: globalCol,
                value: val,
                hasFormula: !!f
              })
            }
          }
        }
      })

    } else {
      const editableCells = editableMap[name] || []
      if (!editableCells.length) return
      let minRow = Infinity, maxRow = -Infinity
      let minCol = Infinity, maxCol = -Infinity
      for (let i = 0; i < editableCells.length; i++) {
        const rc = editableCells[i]
        if (rc.row < minRow) minRow = rc.row
        if (rc.row > maxRow) maxRow = rc.row
        if (rc.col < minCol) minCol = rc.col
        if (rc.col > maxCol) maxCol = rc.col
      }
      const numRows = maxRow - minRow + 1
      const numCols = maxCol - minCol + 1
      const range = ssheet.getRange(minRow, minCol, numRows, numCols)
      const values = range.getValues()
      const formulas = range.getFormulas()
      const posSet = new Set()
      for (let i = 0; i < editableCells.length; i++) {
        const rc = editableCells[i]
        posSet.add(rc.row + ',' + rc.col)
      }
      for (let r = 0; r < numRows; r++) {
        for (let c = 0; c < numCols; c++) {
          const globalRow = minRow + r
          const globalCol = minCol + c
          if (!posSet.has(globalRow + ',' + globalCol)) continue
          const val = values[r][c]
          const f = formulas[r][c]
          if (val !== '' || f) {
            cells.push({ row: globalRow, col: globalCol, value: val, hasFormula: !!f })
          }
        }
      }
    }

    if (cells.length) {
      data[name] = { cells }
      const elapsed = ((Date.now() - sheetStart) / 1000).toFixed(2)
      summary.push({ sheet: name, cells: cells.length, time: elapsed })
    }
  })

  const compressed = compress(JSON.stringify(data))
  saveSheet.getRange(saveCell).setValue(compressed)
  const totalElapsed = ((Date.now()-startAll)/1000).toFixed(2)
  toastMessage('📋 Assignments Helper', `Assignments saved to ${label}`, 6)
}

function loadAssignmentsFromSlot(slotId, chosenSheets, rawString) {
  toastMessage('📋 Assignments Helper', `Loading/Importing...`, 6)
  const label = SAVE_LABELS[slotId-1] || `Slot ${slotId}`
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const saveSheet = ss.getSheetByName(SAVE_SHEET_NAME)
  const raw = rawString || saveSheet.getRange(`A${slotId}`).getValue()
  if (!raw) {
    toastMessage('📋 Assignments Helper', 'No saved assignments found.', 6)
    return {ok:false,msg:'Empty slot'}
  }
  let data
  try { data = JSON.parse(decompress(raw)) }
  catch(e) { return {ok:false,msg:'Decode failed'} }

  if (chosenSheets.includes('Roster')) {
    loadSingleSheet_('Roster', data['Roster'], ss)
    SpreadsheetApp.flush()
  }

  for (const [name, sheetData] of Object.entries(data)) {
    if (name === 'Roster' || !chosenSheets.includes(name)) continue
    loadSingleSheet_(name, sheetData, ss)
  }

  toastMessage('📋 Assignments Helper', `Assignments loaded from ${label}`, 6)
  return {ok:true}
}

function loadSingleSheet_(name, sheetData, ss) {
  const sheet = ss.getSheetByName(name)
  if (!sheet || !sheetData || !sheetData.cells || !sheetData.cells.length) return

  let filteredCells

  if (name === 'Roster') {
    const allowedRanges = [
      sheet.getRange(ROSTER_RANGE),
      ...ROSTER_EXTRA_ALLOWED.map(a1 => sheet.getRange(a1))
    ]
    const bounds = allowedRanges.map(r => ({
      minRow: r.getRow(),
      maxRow: r.getLastRow(),
      minCol: r.getColumn(),
      maxCol: r.getLastColumn()
    }))
    filteredCells = []
    for (let i = 0; i < sheetData.cells.length; i++) {
      const c = sheetData.cells[i]
      let inAny = false
      for (let j = 0; j < bounds.length; j++) {
        const b = bounds[j]
        if (c.row >= b.minRow && c.row <= b.maxRow && c.col >= b.minCol && c.col <= b.maxCol) {
          inAny = true
          break
        }
      }
      if (inAny) filteredCells.push(c)
    }
    if (!filteredCells.length) return
  } else {
    const editableMap = getEditableRegions()
    const editableCells = editableMap[name] || []
    if (!editableCells.length) return
    const mapSet = new Set()
    for (let i = 0; i < editableCells.length; i++) {
      const rc = editableCells[i]
      mapSet.add(rc.row + ',' + rc.col)
    }
    filteredCells = []
    for (let i = 0; i < sheetData.cells.length; i++) {
      const c = sheetData.cells[i]
      if (mapSet.has(c.row + ',' + c.col)) filteredCells.push(c)
    }
    if (!filteredCells.length) return
  }

  const writes = {}
  for (let i = 0; i < filteredCells.length; i++) {
    const c = filteredCells[i]
    if (!writes[c.row]) writes[c.row] = {}
    writes[c.row][c.col] = c.value
  }
  for (const r in writes) {
    const cols = Object.keys(writes[r]).map(Number).sort(function(a, b) { return a - b })
    let start = 0
    for (let i = 1; i <= cols.length; i++) {
      if (cols[i] !== cols[i - 1] + 1) {
        const blockCols = cols.slice(start, i)
        const values = [blockCols.map(function(col) { return writes[r][col] })]
        sheet.getRange(parseInt(r, 10), blockCols[0], 1, blockCols.length).setValues(values)
        start = i
      }
    }
  }
}

const GUILD_SAVE_SHEET_NAME = 'Save Data';
const GUILD_SAVE_CELL = 'B1';
const GUILD_TARGET_SHEETS = ['Guild Manager', 'Group Builder (25)', 'Group Builder (10)'];

const GUILD_RANGES = {
  'Guild Manager': 'E7:BM156',
  'Group Builder (25)': 'C7:AK47',
  'Group Builder (10)': 'C7:AE47'
};

const GUILD_ALLOWED_COLS = [5,11,13,15,16,18,20,21,23,25,26,27,29,30,32,34,35,37,39,41,43,45,46,48,50,51,53,55,56,58,60,62,63,65];

function saveGuildData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const saveSheet = ss.getSheetByName(GUILD_SAVE_SHEET_NAME);
  const startAll = Date.now();
  const data = {};

  GUILD_TARGET_SHEETS.forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) return;

    const range = sheet.getRange(GUILD_RANGES[name]);
    const values = range.getValues();
    const formulas = range.getFormulas();
    const cells = [];

    for (let r = 0; r < values.length; r++) {
      for (let c = 0; c < values[0].length; c++) {
        const val = values[r][c];
        const hasFormula = formulas[r][c] !== '';
        if (typeof val === 'object' && val?.valueType === 'IMAGE') continue;
        if (val !== '' || hasFormula) {
          cells.push({
            row: r + range.getRow(),
            col: c + range.getColumn(),
            value: val
          });
        }
      }
    }

    if (cells.length) {
      data[name] = { cells };
    }
  });

  const compressed = compress(JSON.stringify(data));
  saveSheet.getRange(GUILD_SAVE_CELL).setValue(compressed);

  if (ENABLE_LOG) Logger.log(`Guild save complete in ${((Date.now() - startAll)/1000).toFixed(2)}s`);
  toastMessage('📋 Guild Data Helper', 'Guild data saved!', 5);
}

function loadGuildData(chosenSheets, rawString) {
  toastMessage('📥 Roster Helper', `Loading/Importing...`, 6)
  chosenSheets = Array.isArray(chosenSheets) ? chosenSheets : []
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const saveSheet = ss.getSheetByName(GUILD_SAVE_SHEET_NAME)
  const raw = rawString || saveSheet.getRange(GUILD_SAVE_CELL).getValue()
  if (!raw) {
    toastMessage('📋 Guild Data Helper', 'No saved guild data found.', 6)
    if (ENABLE_LOG) Logger.log('Guild data empty')
    return { ok: false, msg: 'Empty slot' }
  }
  let data
  try { data = JSON.parse(decompress(raw)) }
  catch(e) {
    toastMessage('📋 Guild Data Helper', 'Could not read or decode the provided data.', 6)
    if (ENABLE_LOG) Logger.log('Decode error: ' + e)
    return { ok: false, msg: 'Decode failed' }
  }

  if (chosenSheets.includes('Guild Manager') && data['Guild Manager']) {
    loadGuildManager('Guild Manager', data['Guild Manager'], ss)
    SpreadsheetApp.flush()
  }

  chosenSheets
  .filter(name => name.includes('Group Builder') && data[name])
  .forEach(name => {
    loadGroupBuilder(name, data[name], ss)
  })

  toastMessage('📋 Guild Data Helper', rawString ? 'Guild data loaded from string' : 'Guild data loaded successfully', 6)
  return { ok: true }
}

function loadGuildManager(sheetName, sheetData, ss) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet || !sheetData || !sheetData.cells || !sheetData.cells.length) return;

  const allowedLetters = ['E','K','M','O','P','R','T','U','W','Y','Z','AB','AD','AE','AG','AI','AJ','AL','AN','AO','AQ','AS','AT','AV','AX','AY','BA','BC','BD','BF','BH','BI','BK','BM'];
  const allowedCols = allowedLetters.map(colLetterToNumber);
  const allowedSet = new Set(allowedCols);

  const bounds = GUILD_RANGES[sheetName];
  const bRange = sheet.getRange(bounds);
  const minRow = bRange.getRow();
  const maxRow = bRange.getLastRow();
  const minCol = bRange.getColumn();
  const maxCol = bRange.getLastColumn();

  const filtered = [];
  for (let i = 0; i < sheetData.cells.length; i++) {
    const c = sheetData.cells[i];
    if (typeof c.value === 'object') continue;
    if (c.row < minRow || c.row > maxRow) continue;
    if (c.col < minCol || c.col > maxCol) continue;
    if (!allowedSet.has(c.col)) continue;
    filtered.push(c);
  }
  if (!filtered.length) return;

  const byRow = {};
  for (let i = 0; i < filtered.length; i++) {
    const c = filtered[i];
    if (!byRow[c.row]) byRow[c.row] = [];
    byRow[c.row].push(c);
  }

  const rows = Object.keys(byRow).map(Number).sort((a,b)=>a-b);
  for (let rIdx = 0; rIdx < rows.length; rIdx++) {
    const r = rows[rIdx];
    const cells = byRow[r].sort((a,b)=>a.col-b.col);
    let start = 0;
    for (let i = 1; i <= cells.length; i++) {
      const isBreak = i === cells.length || cells[i].col !== cells[i-1].col + 1;
      if (isBreak) {
        const block = cells.slice(start, i);
        const startCol = block[0].col;
        const width = block.length;
        const values = [block.map(x => x.value === undefined || x.value === null ? '' : x.value)];
        sheet.getRange(r, startCol, 1, width).setValues(values);
        start = i;
      }
    }
  }
}

function loadGroupBuilder(sheetName, sheetData, ss) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    Logger.log(`Sheet not found: ${sheetName}`);
    return;
  }
  if (!sheetData || !sheetData.cells || !sheetData.cells.length) {
    Logger.log(`No sheet data available for ${sheetName}`);
    return;
  }

  Logger.log(`Starting loadGroupBuilder for ${sheetName}`);
  Logger.log(`Raw saved cell count: ${sheetData.cells.length}`);

  const editableMap = getEditableRegions();
  const editable = editableMap[sheetName] || [];

  Logger.log(`Editable cells loaded for ${sheetName}: ${editable.length}`);
  if (!editable.length) {
    Logger.log(`No editable regions for ${sheetName}, aborting write.`);
    return;
  }

  const editableSet = new Set(editable.map(e => `${e.row},${e.col}`));
  const filtered = [];
  for (let i = 0; i < sheetData.cells.length; i++) {
    const c = sheetData.cells[i];
    const key = `${c.row},${c.col}`;
    if (!editableSet.has(key)) continue;
    if (typeof c.value === 'object') continue;
    filtered.push(c);
  }

  Logger.log(`Filtered to ${filtered.length} restorable cells for ${sheetName}`);
  if (!filtered.length) {
    Logger.log(`No cells to write!`);
    return;
  }

  Logger.log(`Sample filtered cells: ${JSON.stringify(filtered.slice(0, 5), null, 2)}`);

  const byRow = {};
  for (let i = 0; i < filtered.length; i++) {
    const c = filtered[i];
    if (!byRow[c.row]) byRow[c.row] = [];
    byRow[c.row].push(c);
  }

  const rowKeys = Object.keys(byRow);
  Logger.log(`Rows with at least 1 valid cell: ${rowKeys.length}`);

  const rows = rowKeys.map(Number).sort((a,b)=>a-b);
  for (let rIdx = 0; rIdx < rows.length; rIdx++) {
    const r = rows[rIdx];
    const cells = byRow[r].sort((a,b)=>a.col-b.col);
    let start = 0;
    for (let i = 1; i <= cells.length; i++) {
      const isBreak = i === cells.length || cells[i].col !== cells[i-1].col + 1;
      if (isBreak) {
        const block = cells.slice(start, i);
        const startCol = block[0].col;
        const width = block.length;
        const values = [block.map(x => x.value == null ? '' : x.value)];
        Logger.log(`Writing row=${r}, startCol=${startCol}, width=${width}, values=${JSON.stringify(values[0])}`);
        try {
          sheet.getRange(r, startCol, 1, width).setValues(values);
        } catch(e) {
          Logger.log(`ERROR writing row=${r}, col=${startCol}: ${e.message}`);
        }
        start = i;
      }
    }
  }

  Logger.log(`loadGroupBuilder completed for ${sheetName}`);
}

function colLetterToNumber(letter) {
  let col = 0;
  for (let i = 0; i < letter.length; i++) {
    col = col * 26 + (letter.charCodeAt(i) - 64);
  }
  return col;
}

function compress(str) {
  const bytes = Utilities.newBlob(str).getBytes();
  const gz = Utilities.gzip(Utilities.newBlob(bytes));
  return Utilities.base64Encode(gz.getBytes());
}

function decompress(base64Str) {
  const bytes = Utilities.base64Decode(base64Str);
  const blob = Utilities.newBlob(bytes, 'application/x-gzip');
  const decompressed = Utilities.ungzip(blob);
  return decompressed.getDataAsString();
}

function getEditableRegions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const saveSheet = ss.getSheetByName('Save Data');
  const raw = saveSheet.getRange('C1').getValue();
  return raw ? JSON.parse(decompress(raw)) : {};
}

function toastMessage(title, msg, secs) {
  SpreadsheetApp.getActiveSpreadsheet().toast(msg, title, secs);
}


function getGuildTargetSheets() {
  return GUILD_TARGET_SHEETS
}

function showLoadGuildModal() {
  const targetSheets = getGuildTargetSheets()
  const template = HtmlService.createTemplateFromFile("LoadGuildDataPopup")
  template.sheetNames = targetSheets
  const html = template.evaluate().setWidth(550).setHeight(500)
  SpreadsheetApp.getUi().showModalDialog(html, "Load Guild Data")
}

function getExportData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const sheet = ss.getSheetByName('Save Data')
  const guild = sheet.getRange('B1').getValue() || 'No Data Found'
  const assignVals = sheet.getRangeList(['A1','A2','A3','A4','A5','A6']).getRanges().map(r=>r.getValue() || 'No Data Found')
  return {
    guild,
    assignments: [
      {label:'Split 1', value:assignVals[0]},
      {label:'Split 2', value:assignVals[1]},
      {label:'Split 3', value:assignVals[2]},
      {label:'Split 4', value:assignVals[3]},
      {label:'Split 5', value:assignVals[4]},
      {label:'Custom', value:assignVals[5]},
    ]
  }
}

function showSaveDataModal() {
  const html = HtmlService.createTemplateFromFile('SaveDataHub')
  html.data = getExportData()
  const output = html.evaluate().setWidth(800).setHeight(1200)
  SpreadsheetApp.getUi().showModalDialog(output, 'Export Hub')
}

function debugViewGuildSaveData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const saveSheet = ss.getSheetByName(GUILD_SAVE_SHEET_NAME)
  const raw = saveSheet.getRange(GUILD_SAVE_CELL).getValue()
  if (!raw) {
    Logger.log('No guild save data found')
    return
  }
  try {
    const data = JSON.parse(decompress(raw))
    Logger.log(`Guild save dump:\n${JSON.stringify(data, null, 2)}`)
  } catch (e) {
    Logger.log(`Guild save decode error: ${e}`)
  }
}


function debugEditableMap() {
  const map = getEditableRegions();
  Logger.log(JSON.stringify(map['Roster'].slice(0, 10), null, 2));
  Logger.log(`Roster editable cells: ${map['Roster'].length}`);
}