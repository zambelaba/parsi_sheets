/**
 * AutoAssignData: one row per assignment block, read by header name, so columns
 * can be added or reordered freely.
 *
 * Required:
 *   Encounter        label other rows refer to; also the tab name unless Sheet is set
 *   CellAnchor       first player cell of the block, or one per group: "BC10, BH10, BM10"
 *   GroupAmount      number of groups; 0 (or a blank CellAnchor) turns the row off
 *   GroupSize        rows per group, one number or one per group: "4, 5, 4, 4, 2"
 *
 * Optional:
 *   Sheet            tab(s) to write to. "*" = every tab the other rows name, "-X" removes X
 *   Raid             MH / BT, for the per-instance buttons. Blank, "*" or "all" = both
 *   ColumnStep       columns between groups when CellAnchor is a single cell (default 1)
 *   AutoRoles        ordered list of Roles, Classes or ClassSpecs, most preferred first.
 *                    A bare Class never takes that class's tank or healer; "Any X" does.
 *                    "Flex X" matches on the off-spec. Blank = everyone.
 *   EnsureEach       every group gets at least one of these (same vocabulary)
 *   Preferred        names placed first, bypassing AutoRoles; absent players are skipped
 *   AllowDuplicates  TRUE lets a Preferred name appear more than once
 *   WriteBackTo      named range seeded with the eligible pool while completely empty
 *   LimitTo          never place more people than that row did (it must sit above)
 *   ExcludeFrom      skip people that row placed (it must sit above)
 *   InheritFrom      keep people on the marker they had in that row (it must sit above)
 *   MarkerNames      group names, used by InheritFrom and in messages
 *   MarkerRoles      soft Role / Class / ClassSpec preference per marker
 *   TotemParties     TRUE builds each marker around a shaman's raid party
 *   BalanceGroups    with TotemParties, caps each marker at an even share of the pool
 *   BackupRows       bottom rows of each group that are reserves (reporting only)
 *   SkipWhen         "AE2 = Good DPS" or "AE2 != X": the row is skipped when the cell matches
 *   Selection        TRUE when the row picks a few people out of many
 *   Critical         TRUE turns "nobody can do this" into a red warning
 *
 * Every target cell must be a plain value: a block holding a formula is refused.
 */

/** Every enabled row of AutoAssignData, one entry per target tab. */
function loadEncounters_(ss, scope) {
  const sheetName = AUTO_ASSIGN.ENCOUNTERS_SHEET;
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error(`Sheet "${sheetName}" not found`);

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const header = locateHeaderRow_(sheet, ['Encounter', 'CellAnchor', 'GroupAmount', 'GroupSize'], sheetName);
  if (lastRow <= header.row) return [];

  const values = sheet.getRange(header.row + 1, 1, lastRow - header.row, sheet.getLastColumn()).getValues();
  const read = (row, name) => {
    const index = header.index[normalizeHeader_(name)];
    return index === undefined ? '' : row[index];
  };

  const isWildcard = tab => tab === '*' || tab.toLowerCase() === 'all';

  const allRows = values
    .filter(row => String(read(row, 'Encounter')).trim() !== '')
    .map(row => {
      const label = String(read(row, 'Encounter')).trim();
      const groupCount = Number(read(row, 'GroupAmount')) || 0;
      const anchors = parseAnchors_(read(row, 'CellAnchor'));
      const tabEntries = parseList_(read(row, 'Sheet'));
      const tabs = tabEntries.filter(t => !isWildcard(t) && t.charAt(0) !== '-');

      return {
        tabs: {
          wildcard: tabEntries.some(isWildcard),
          excluded: tabEntries.filter(t => t.charAt(0) === '-').map(t => t.slice(1).trim().toLowerCase()),
          named: tabs.length ? tabs : [label],
        },
        encounter: label,
        raid: normalizeRaidScope_(read(row, 'Raid')),
        anchor: anchors.join(', '),
        anchors: anchors,
        groupCount: groupCount,
        groupSizes: parseSizes_(read(row, 'GroupSize'), groupCount),
        columnStep: Number(read(row, 'ColumnStep')) || 1,
        backupRows: Number(read(row, 'BackupRows')) || 0,
        eligibleRoles: parseRoleList_(read(row, 'AutoRoles')),
        ensureEach: parseRoleList_(read(row, 'EnsureEach')),
        preferred: parseList_(read(row, 'Preferred')),
        allowDuplicates: parseBoolean_(read(row, 'AllowDuplicates')),
        writeBackTo: String(read(row, 'WriteBackTo')).trim(),
        limitTo: String(read(row, 'LimitTo')).trim(),
        excludeFrom: parseList_(read(row, 'ExcludeFrom')),
        inheritFrom: String(read(row, 'InheritFrom')).trim(),
        markerNames: parseList_(read(row, 'MarkerNames')),
        markerRoles: parseSlotList_(read(row, 'MarkerRoles')),
        totemParties: parseBoolean_(read(row, 'TotemParties')),
        balanceGroups: parseBoolean_(read(row, 'BalanceGroups')),
        skipWhen: parseCondition_(read(row, 'SkipWhen')),
        selection: parseBoolean_(read(row, 'Selection')),
        critical: parseBoolean_(read(row, 'Critical')),
      };
    });

  // Scoped before "*" is expanded, so a wildcard row follows the scope for free.
  const rows = scope ? allRows.filter(row => !row.raid || row.raid === scope) : allRows;
  if (scope) Logger.log(`Scoped to "${scope}": kept ${rows.length} of ${allRows.length} row(s)`);

  const namedTabs = [];
  rows.filter(row => !row.tabs.wildcard).forEach(row => row.tabs.named.forEach(name => {
    if (namedTabs.indexOf(name) === -1) namedTabs.push(name);
  }));

  // One entry per target tab, so everything downstream is "assign this block here".
  const expanded = [];
  rows.forEach(row => {
    const sheetNames = row.tabs.wildcard
      ? namedTabs.filter(name => row.tabs.excluded.indexOf(name.toLowerCase()) === -1)
      : row.tabs.named;
    sheetNames.forEach(name => {
      const encounter = Object.assign({}, row, { sheetName: name, repeated: sheetNames.length > 1 });
      delete encounter.tabs;
      expanded.push(encounter);
    });
  });
  return expanded;
}

/**
 * Finds the first row (within the top few) carrying every required header, so
 * config sheets can be Tables or sit below a title. Returns { row, index }.
 */
function locateHeaderRow_(sheet, required, label) {
  const scanTo = Math.min(AUTO_ASSIGN.HEADER_SCAN_ROWS, sheet.getLastRow());

  for (let row = 1; row <= scanTo; row++) {
    const index = {};
    sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0].forEach((header, i) => {
      const key = normalizeHeader_(header);
      if (key && index[key] === undefined) index[key] = i;
    });
    if (required.every(name => index[normalizeHeader_(name)] !== undefined)) return { row: row, index: index };
  }

  throw new Error(`Could not find a header row in "${label}" containing: ${required.join(', ')}. ` +
    `Checked the first ${scanTo} row(s).`);
}

function normalizeHeader_(header) {
  return String(header).trim().toLowerCase().replace(/[^a-z0-9]/g, '');
}

/** "Star, Circle, Square" -> ['Star', 'Circle', 'Square']. */
function parseList_(value) {
  return String(value).split(',').map(s => s.trim()).filter(s => s !== '');
}

/** Like parseList_, lowercased: the form every AutoRoles-style match expects. */
function parseRoleList_(value) {
  return parseList_(value).map(s => s.toLowerCase());
}

/** Keeps empty entries so "Ranged, , Melee" still lines up with groups 1 and 3. */
function parseSlotList_(value) {
  const text = String(value).trim();
  return text === '' ? [] : text.split(',').map(s => s.trim());
}

/** "4, 5, 4" -> [4, 5, 4]. A single number applies to every group. */
function parseSizes_(value, groupCount) {
  const sizes = String(value).split(',').map(part => Number(part.trim())).filter(n => !isNaN(n) && n > 0);
  if (sizes.length === 1) return new Array(groupCount).fill(sizes[0]);
  return sizes;
}

/** "Najentus!BC10, BH10" -> ['BC10', 'BH10']. The row already says which sheet. */
function parseAnchors_(value) {
  return String(value).split(',').map(stripSheetPrefix_).filter(part => part !== '');
}

function stripSheetPrefix_(reference) {
  const text = String(reference).trim();
  const bang = text.lastIndexOf('!');
  return bang === -1 ? text : text.slice(bang + 1).trim();
}

/** "AE2 = Good DPS" -> { cell: 'AE2', negate: false, value: 'Good DPS' }. Also accepts "!=". */
function parseCondition_(value) {
  const parts = String(value).trim().match(/^(.+?)\s*(!=|=)\s*(.*)$/);
  if (!parts) return null;
  return { cell: stripSheetPrefix_(parts[1]), negate: parts[2] === '!=', value: parts[3].trim() };
}

function parseBoolean_(value) {
  const text = String(value).trim().toLowerCase();
  return text === 'true' || text === 'yes' || text === '1';
}

/** Blank, "*" and "all" all mean "every instance". */
function normalizeRaidScope_(value) {
  const text = String(value === null || value === undefined ? '' : value).trim().toLowerCase();
  return text === '*' || text === 'all' ? '' : text;
}
