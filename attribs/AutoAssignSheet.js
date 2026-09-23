/**
 * Reading and writing the boss sheets.
 *
 * Every block's sheet is read once up front (one bounding-box snapshot per
 * sheet, taken after the tank/healer broadcast). That is safe because a run
 * only writes plain values into the blocks, and the rows above them are
 * either other blocks (skipped) or formulas nothing here touches.
 */

/**
 * Resolves every block's ranges (cached on the encounter as .sheet / .ranges)
 * and snapshots each sheet. Returns { owned, snapshots }: owned[sheet]["col,row"]
 * marks cells written by some block, so reading them back is never mistaken for
 * someone already placed.
 */
function prepareSheets_(ss, encounters) {
  const owned = {};
  const boxes = {};

  encounters.forEach(encounter => {
    if (!encounter.groupCount || !encounter.anchors.length) return;
    const sheet = ss.getSheetByName(encounter.sheetName);
    if (!sheet) return;

    let ranges;
    try {
      ranges = resolveTargetRanges_(sheet, encounter);
    } catch (err) {
      return; // reported when the row itself runs
    }
    encounter.sheet = sheet;
    encounter.ranges = ranges;

    const cells = owned[encounter.sheetName] = owned[encounter.sheetName] || {};
    ranges.forEach(range => {
      const column = range.getColumn();
      const firstRow = range.getRow();
      const lastRow = firstRow + range.getNumRows() - 1;
      for (let row = firstRow; row <= lastRow; row++) cells[`${column},${row}`] = true;

      const top = Math.max(1, firstRow - AUTO_ASSIGN.ABOVE_ANCHOR_ROWS);
      const box = boxes[encounter.sheetName];
      if (!box) {
        boxes[encounter.sheetName] = { sheet: sheet, minRow: top, maxRow: lastRow, minCol: column, maxCol: column };
        return;
      }
      box.minRow = Math.min(box.minRow, top);
      box.maxRow = Math.max(box.maxRow, lastRow);
      box.minCol = Math.min(box.minCol, column);
      box.maxCol = Math.max(box.maxCol, column);
    });
  });

  const snapshots = {};
  Object.keys(boxes).forEach(name => {
    const box = boxes[name];
    const rows = box.maxRow - box.minRow + 1;
    const cols = box.maxCol - box.minCol + 1;
    if (rows * cols > AUTO_ASSIGN.MAX_SNAPSHOT_CELLS) return;

    const range = box.sheet.getRange(box.minRow, box.minCol, rows, cols);
    snapshots[name] = {
      firstRow: box.minRow,
      firstCol: box.minCol,
      values: range.getDisplayValues(),
      formulas: range.getFormulas(),
    };
  });

  return { owned: owned, snapshots: snapshots };
}

/** A cell from a snapshot grid ('values' or 'formulas'), or null outside it. */
function snapshotCell_(snapshot, grid, row, column) {
  if (!snapshot) return null;
  const r = row - snapshot.firstRow;
  const c = column - snapshot.firstCol;
  const cells = snapshot[grid];
  if (r < 0 || c < 0 || r >= cells.length || c >= cells[0].length) return null;
  return cells[r][c];
}

/**
 * One single-column range per group: either one anchor per group, or a single
 * anchor repeated every ColumnStep columns.
 */
function resolveTargetRanges_(sheet, encounter) {
  const { anchors, groupSizes: sizes, groupCount } = encounter;
  if (sizes.length !== groupCount) {
    throw new Error(`GroupSize gives ${sizes.length} size(s) but GroupAmount is ${groupCount}`);
  }

  if (anchors.length > 1) {
    if (anchors.length !== groupCount) {
      throw new Error(`CellAnchor lists ${anchors.length} anchor(s) but GroupAmount is ${groupCount}`);
    }
    return anchors.map((anchor, i) => sheet.getRange(anchor).offset(0, 0, sizes[i], 1));
  }

  const first = sheet.getRange(anchors[0]);
  return sizes.map((size, i) => first.offset(0, i * encounter.columnStep, size, 1));
}

/**
 * The non-empty cells in the few rows directly above a block, nearest first,
 * skipping cells some other block writes: { text, row, column, formula }.
 */
function readAboveAnchor_(sheet, range, owned, snapshot) {
  const anchorRow = range.getRow();
  if (anchorRow <= 1) return [];

  const height = Math.min(AUTO_ASSIGN.ABOVE_ANCHOR_ROWS, anchorRow - 1);
  const startRow = anchorRow - height;
  const column = range.getColumn();
  const direct = snapshot ? null : sheet.getRange(startRow, column, height, 1);
  const values = direct ? direct.getDisplayValues() : null;
  const formulas = direct ? direct.getFormulas() : null;

  const found = [];
  for (let i = 0; i < height; i++) {
    const row = startRow + i;
    const raw = snapshot ? snapshotCell_(snapshot, 'values', row, column) : values[i][0];
    const formula = snapshot ? snapshotCell_(snapshot, 'formulas', row, column) : formulas[i][0];
    const text = String(raw === null ? '' : raw).trim();
    if (!text || owned[`${column},${row}`]) continue;
    found.push({ text: text, row: row, column: column, formula: String(formula || '').trim() });
  }
  return found.reverse();
}

/**
 * Lowercased name -> cell, for people a formula above the block already puts
 * in the group (a marker's healer row, "=AV18"). Typed names belong to some
 * other table, and formulas reading FROM the block only mirror it.
 */
function prePlacedNames_(sheet, ranges, owned, snapshot) {
  const byName = new Map();
  ranges.forEach(range => {
    readAboveAnchor_(sheet, range, owned, snapshot).forEach(cell => {
      if (!cell.formula || formulaReadsFromRanges_(cell.formula, ranges)) return;
      const key = cell.text.toLowerCase();
      if (!byName.has(key)) byName.set(key, a1_(cell.row, cell.column));
    });
  });
  return byName;
}

/** The names above each block (its healers), nearest first, however they got there. */
function markerHealers_(sheet, ranges, owned, snapshot) {
  return ranges.map(range => readAboveAnchor_(sheet, range, owned, snapshot).map(cell => cell.text));
}

/** First target cell holding a formula, as { cell, onAnchorRow }, or null. */
function findFormulaCell_(ranges, snapshot) {
  for (const range of ranges) {
    const column = range.getColumn();
    const firstRow = range.getRow();
    const formulas = snapshot ? null : range.getFormulas();

    for (let r = 0; r < range.getNumRows(); r++) {
      const formula = snapshot ? snapshotCell_(snapshot, 'formulas', firstRow + r, column) : formulas[r][0];
      if (formula) return { cell: a1_(firstRow + r, column), onAnchorRow: r === 0 };
    }
  }
  return null;
}

/**
 * The groups another row has written this run, for InheritFrom. Read from the
 * sheet, so that row must sit above this one in AutoAssignData.
 */
function readInheritedGroups_(ss, encounters, label) {
  const wanted = label.trim().toLowerCase();
  const source = encounters.filter(e => e.encounter.trim().toLowerCase() === wanted)[0];

  if (!source) return { error: `no Encounters row is named "${label}"` };
  if (!source.groupCount || !source.anchors.length) return { error: `"${label}" has no groups configured to copy` };

  const sheet = ss.getSheetByName(source.sheetName);
  if (!sheet) return { error: `"${label}" writes to a tab called "${source.sheetName}", which does not exist` };

  let ranges;
  try {
    ranges = resolveTargetRanges_(sheet, source);
  } catch (err) {
    return { error: `"${label}" has a CellAnchor that could not be read` };
  }

  return {
    markerNames: source.markerNames,
    groups: ranges.map(range => range.getDisplayValues().flat()
      .map(value => String(value).trim())
      .filter(value => value !== '')),
  };
}

/** Evaluates a SkipWhen condition: { matched, actual } or { error }. */
function evaluateCondition_(sheet, condition) {
  let actual;
  try {
    actual = String(sheet.getRange(condition.cell).getDisplayValue()).trim();
  } catch (err) {
    return { error: `could not read the cell "${condition.cell}" (${(err && err.message) || err})` };
  }
  const same = actual.toLowerCase() === condition.value.toLowerCase();
  const matched = condition.negate ? !same : same;
  Logger.log(`SkipWhen on "${sheet.getName()}": ${condition.cell} is "${actual}" -> ${matched ? 'SKIP' : 'run'}`);
  return { matched: matched, actual: actual };
}

/**
 * Clears each block and writes its group. Returns { rejected, overflow }.
 *
 * A dropdown listing "players not already assigned" refuses someone still
 * sitting in the cell being rewritten unless the list recalculates between
 * clearing and writing. That recalculation costs a flush per block, so it is
 * only paid when the fast attempt is refused. If the retry is refused too, the
 * block is put back as it was.
 */
function writeGroups_(ranges, groups, snapshot, label) {
  const previous = ranges.map(range => {
    if (!snapshot) return range.getValues();
    const column = range.getColumn();
    const firstRow = range.getRow();
    const copy = [];
    for (let i = 0; i < range.getNumRows(); i++) {
      const value = snapshotCell_(snapshot, 'values', firstRow + i, column);
      copy.push([value === null ? '' : value]);
    }
    return copy;
  });

  const overflow = [];
  const plans = ranges.map((range, i) => {
    const names = groups[i].map(p => p.player);
    const capacity = range.getNumRows();
    if (names.length > capacity) overflow.push({ group: i, capacity: capacity, dropped: names.slice(capacity) });
    return names.slice(0, capacity).map(name => [name]);
  });

  Logger.log(`writeGroups_ ${label}: ` + ranges.map((range, i) =>
    `${range.getA1Notation()}=[${plans[i].map(row => row[0]).join(', ')}]`).join(', '));

  const clear = () => ranges.forEach(range => range.clearContent());
  const write = () => ranges.forEach((range, i) => {
    if (plans[i].length) range.offset(0, 0, plans[i].length, 1).setValues(plans[i]);
  });

  // flush() makes a refused write throw here, rather than later on some
  // unrelated read that would then be blamed for it.
  try {
    clear();
    write();
    SpreadsheetApp.flush();
    return { rejected: [], overflow: overflow };
  } catch (err) {
    Logger.log(`writeGroups_ ${label}: write refused (${err.message}), retrying after a recalculation`);
  }

  try {
    clear();
    SpreadsheetApp.flush();
    write();
    SpreadsheetApp.flush();
    return { rejected: [], overflow: overflow };
  } catch (err) {
    Logger.log(`writeGroups_ ${label}: retry refused too (${err.message})`);
    const intended = {};
    ranges.forEach((range, i) => plans[i].forEach((row, r) => {
      intended[a1_(range.getRow() + r, range.getColumn())] = row[0];
    }));
    return { rejected: diagnoseWriteFailure_(ranges, groups, previous, err.message, intended), overflow: [] };
  }
}

/**
 * Works out which names a dropdown refused, then restores the block. The cells
 * are cleared first because the dropdown lists are built from who is unassigned.
 */
function diagnoseWriteFailure_(ranges, groups, previous, message, intended) {
  let rejected = [];
  try {
    ranges.forEach(range => range.clearContent());
    SpreadsheetApp.flush();
    rejected = findValidationRejects_(ranges, groups);
  } catch (err) {
    Logger.log(`Could not work out which names were refused: ${err.message}`);
  }

  try {
    ranges.forEach((range, i) => range.setValues(previous[i]));
    SpreadsheetApp.flush();
  } catch (err) {
    Logger.log(`Could not restore the previous assignment: ${err.message}`);
  }

  if (rejected.length) return rejected;

  // Custom-formula rules have no list to read, but Google's message names the cell.
  Logger.log(`No dropdown list could be read. Raw error: ${message}`);
  const match = /cell ([A-Z]+[0-9]+)/i.exec(String(message));
  const cell = match ? match[1].toUpperCase() : '';
  if (cell && intended[cell]) return [`${intended[cell]} (into ${cell})`];
  if (cell) return [`the name going into ${cell}`];
  return ['a name on this sheet'];
}

/** Names the (rejecting) dropdowns in these cells would refuse. */
function findValidationRejects_(ranges, groups) {
  const rejected = [];
  ranges.forEach((range, i) => {
    const names = groups[i].map(p => p.player).slice(0, range.getNumRows());
    if (!names.length) return;

    const rules = range.getDataValidations();
    names.forEach((name, row) => {
      const rule = rules[row] && rules[row][0];
      if (!rule || rule.getAllowInvalid()) return;
      const allowed = allowedValuesFor_(rule);
      if (allowed && allowed.indexOf(name.trim().toLowerCase()) === -1) rejected.push(name);
    });
  });
  return rejected;
}

/** Lowercased values a list/range dropdown accepts, or null for any other rule. */
function allowedValuesFor_(rule) {
  const criteria = SpreadsheetApp.DataValidationCriteria;
  try {
    const type = rule.getCriteriaType();
    const values = rule.getCriteriaValues();
    if (type === criteria.VALUE_IN_LIST) return values[0].map(v => String(v).trim().toLowerCase());
    if (type === criteria.VALUE_IN_RANGE) {
      return values[0].getDisplayValues().flat().map(v => String(v).trim().toLowerCase()).filter(v => v !== '');
    }
  } catch (err) {
    // A dropdown whose source range was deleted throws when read; treat it as unrestricted.
    Logger.log(`Could not read a dropdown's allowed values: ${err.message}`);
  }
  return null;
}

/**
 * Whether a formula refers to a cell inside these ranges (same sheet only).
 * Plain string matching: it only needs "=AV18" and "=IF(BO19=..,BO18,BO19)".
 */
function formulaReadsFromRanges_(formula, ranges) {
  const local = String(formula)
    .replace(/('[^']*'|[A-Za-z_][A-Za-z0-9_.]*)!\$?[A-Z]{1,3}\$?[0-9]+(:\$?[A-Z]{1,3}\$?[0-9]+)?/g, ' ');
  const refs = local.match(/\$?[A-Z]{1,3}\$?[0-9]+/g) || [];

  return refs.some(ref => {
    const parts = /^\$?([A-Z]{1,3})\$?([0-9]+)$/.exec(ref);
    const column = columnNumber_(parts[1]);
    const row = Number(parts[2]);
    return ranges.some(range => column === range.getColumn() &&
      row >= range.getRow() && row < range.getRow() + range.getNumRows());
  });
}

/** 1 -> "A", 27 -> "AA". */
function columnLetter_(column) {
  let letter = '';
  for (let n = column; n > 0; n = Math.floor((n - 1) / 26)) {
    letter = String.fromCharCode(65 + (n - 1) % 26) + letter;
  }
  return letter;
}

/** "A" -> 1, "AA" -> 27. */
function columnNumber_(letters) {
  let column = 0;
  for (let i = 0; i < letters.length; i++) column = column * 26 + (letters.charCodeAt(i) - 64);
  return column;
}

function a1_(row, column) {
  return columnLetter_(column) + row;
}
