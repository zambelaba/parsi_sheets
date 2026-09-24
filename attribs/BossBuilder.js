/**
 * Menu > Add Boss Tab…: builds a boss tab from blocks and wires every block
 * into the auto-assign run.
 *
 *   Tanks / Healers   named ranges <Boss>_tanks / <Boss>_healers, filled by the
 *                     tank and healer broadcast from Quick Assigns
 *   Marker groups     one AutoAssignData row, optionally with a healer row above
 *                     each marker linked to the Healers block
 *   Pick list         one AutoAssignData row (interrupts, kiters, soakers…)
 *   Seat grid         one AutoAssignData row, one column per seat (vehicles…)
 *   Raid-wide duties  copies of the "*" rows (Bloodlust, PI…) placed on this tab
 *
 * The tab copies Supremus's frame and cell styles. The "*" rows write at fixed
 * cells that only exist on the original boss tabs, so the new tab is excluded
 * from them; ticking a raid-wide duty copies that row onto the tab instead.
 */
const BOSS_BUILDER = Object.freeze({
  STYLE_TAB: 'Supremus',
  RAID_DATA_SHEET: 'Raid Data',
  AREA: { firstColumn: 42, lastColumn: 88, firstRow: 7 }, // AP7:CJ, under the Assignments banner
  GAP_COLUMNS: 2,
  GAP_ROWS: 1,
  CELL_WIDTH: 5,
  LIST_WIDTH: 11,
  MAX_ROWS: 25,
  MAX_COLUMNS: 9,
  LINK_ROW: 31,
  // Supremus cells whose style each part of a block copies.
  STYLES: {
    icon: 'AP7', title: 'AQ7', header: 'AQ8', label: 'AP9', note: 'AR9', player: 'AV9',
    groupIcon: 'BC7', groupTitle: 'BD7', markerIcon: 'BC8', markerLabel: 'BD8',
    healerRow: 'BC9', groupPlayer: 'BC10', filler: 'A40', link: 'AD31',
  },
  // Named ranges holding the raid marker icons.
  MARKER_ICONS: ['star', 'circle', 'diamond', 'triangle', 'moon', 'square', 'cross', 'skull'],
});

function showBossBuilder() {
  const template = HtmlService.createTemplateFromFile('BossBuilderDialog');
  // Sheet text ends up inside a <script>; escaping "<" keeps it from closing it.
  template.optionsJson = JSON.stringify(bossBuilderOptions_(SpreadsheetApp.getActiveSpreadsheet())).replace(/</g, '\\u003c');
  SpreadsheetApp.getUi().showModalDialog(template.evaluate().setWidth(760).setHeight(760), 'Add Boss Tab');
}

/** Called by BossBuilderDialog.html. Returns what was created, or throws a message worth showing. */
function createBossTab(spec) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = readEncounterConfig_(ss);
  const plan = planBossTab_(ss, spec, config);

  const sheet = copyBossFrame_(ss, plan, config);
  const roster = ss.getRangeByName('playerNames') || ss.getSheetByName(AUTO_ASSIGN.RAIDER_DATA_SHEET).getRange('A2:A');
  const dropdown = SpreadsheetApp.newDataValidation().requireValueInRange(roster, true).setAllowInvalid(true).build();
  plan.blocks.forEach(block => drawBlock_(sheet, plan, block, dropdown));

  const namedRanges = addBlockNamedRanges_(ss, sheet, plan);
  const images = addBossImageLabels_(ss, plan.name);
  appendBlockRows_(config, plan);
  const excluded = excludeTabFromWildcardRows_(config, plan.name);
  sheet.activate();

  const outside = plan.blocks.filter(b => b.bottom > 38 || b.right > 87).map(b => b.title);
  return {
    tab: plan.name,
    rows: plan.blocks.filter(b => b.label).map(b => b.label),
    namedRanges: namedRanges,
    images: images,
    excludedFrom: excluded,
    outsideClearRange: outside,
  };
}

/** What the dialog offers: known raid tags and the "*" rows that can be copied as duties. */
function bossBuilderOptions_(ss) {
  const config = readEncounterConfig_(ss);
  const tags = [];
  config.rows.forEach(row => {
    const tag = String(row.get('Raid')).trim();
    if (tag && tag !== '*' && tag.toLowerCase() !== 'all' && tags.indexOf(tag) === -1) tags.push(tag);
  });
  return {
    raidTags: tags,
    duties: wildcardDuties_(config).map(row => ({ label: row.label, rows: row.rows, autoRoles: String(row.get('AutoRoles')) })),
  };
}

/**
 * AutoAssignData as { sheet, header, rows }, each row answering get(header)
 * with its value and formula(header) with its formula.
 */
function readEncounterConfig_(ss) {
  const sheet = ss.getSheetByName(AUTO_ASSIGN.ENCOUNTERS_SHEET);
  if (!sheet) throw new Error(`Sheet "${AUTO_ASSIGN.ENCOUNTERS_SHEET}" not found`);
  const header = locateHeaderRow_(sheet, ['Encounter', 'CellAnchor', 'GroupAmount', 'GroupSize'], AUTO_ASSIGN.ENCOUNTERS_SHEET);
  const width = sheet.getLastColumn();
  const count = Math.max(0, sheet.getLastRow() - header.row);
  const range = count ? sheet.getRange(header.row + 1, 1, count, width) : null;
  const values = range ? range.getValues() : [];
  const formulas = range ? range.getFormulas() : [];
  const column = name => header.index[normalizeHeader_(name)];

  const rows = values.map((row, i) => ({
    sheetRow: header.row + 1 + i,
    values: row,
    formulas: formulas[i],
    get: name => (column(name) === undefined ? '' : row[column(name)]),
    formula: name => (column(name) === undefined ? '' : formulas[i][column(name)]),
  }));
  return { sheet: sheet, header: header, width: width, column: column, rows: rows };
}

/** Enabled one-group "*" rows, first row per label. */
function wildcardDuties_(config) {
  const seen = {};
  return config.rows.filter(row => {
    const label = String(row.get('Encounter')).trim();
    const tabs = parseList_(row.get('Sheet'));
    const wildcard = tabs.some(t => t === '*' || t.toLowerCase() === 'all');
    const single = Number(row.get('GroupAmount')) === 1 && parseSizes_(row.get('GroupSize'), 1).length === 1;
    if (!label || !wildcard || !single || !String(row.get('CellAnchor')).trim() || seen[label.toLowerCase()]) return false;
    seen[label.toLowerCase()] = true;
    return true;
  }).map(row => Object.assign(row, {
    label: String(row.get('Encounter')).trim(),
    rows: parseSizes_(row.get('GroupSize'), 1)[0],
  }));
}

/**
 * Validates the dialog's spec and lays the blocks out. Returns
 * { name, raid, raidPlanUrl, prefix, blocks } where each block has its
 * type-specific settings plus top/left/bottom/right on the sheet.
 */
function planBossTab_(ss, spec, config) {
  const name = String(spec.name || '').trim();
  if (!name) throw new Error('Give the boss a name.');
  if (ss.getSheetByName(name)) throw new Error(`A tab called "${name}" already exists.`);
  if (!ss.getSheetByName(BOSS_BUILDER.STYLE_TAB)) throw new Error(`The "${BOSS_BUILDER.STYLE_TAB}" tab is needed for the look.`);

  const titles = {};
  const uniqueTitle = (title, fallback) => {
    const base = String(title || '').trim() || fallback;
    let unique = base;
    for (let n = 2; titles[unique.toLowerCase()]; n++) unique = `${base} ${n}`;
    titles[unique.toLowerCase()] = true;
    return unique;
  };

  const blocks = (spec.blocks || []).map((raw, i) => normalizeBlock_(raw, name, uniqueTitle, i));

  // Raid-wide duties, plus any duty they depend on through LimitTo / ExcludeFrom.
  const duties = wildcardDuties_(config);
  const byLabel = {};
  duties.forEach(duty => { byLabel[duty.label.toLowerCase()] = duty; });
  const wanted = [];
  const want = label => {
    const duty = byLabel[String(label).trim().toLowerCase()];
    if (!duty || wanted.indexOf(duty) !== -1) return;
    [String(duty.get('LimitTo'))].concat(parseList_(duty.get('ExcludeFrom'))).filter(Boolean).forEach(want);
    wanted.push(duty);
  };
  (spec.duties || []).forEach(want);
  wanted.forEach(duty => blocks.push({
    type: 'duty', title: uniqueTitle(duty.label), label: duty.label, rows: duty.rows, source: duty,
  }));

  if (!blocks.length) throw new Error('Add at least one block.');
  resolveExclusions_(blocks);
  layoutBlocks_(blocks);

  return {
    name: name,
    raid: String(spec.raid || '').trim(),
    raidPlanUrl: String(spec.raidPlanUrl || '').trim(),
    prefix: name.replace(/[^A-Za-z0-9]+/g, '_').replace(/^(?=[0-9_])/, 'Boss'),
    blocks: blocks,
  };
}

function normalizeBlock_(raw, bossName, uniqueTitle, index) {
  const count = (value, fallback, max) => {
    const n = Math.floor(Number(value));
    if (!n || n < 1) return fallback;
    if (n > max) throw new Error(`Block ${index + 1}: at most ${max} allowed, got ${n}.`);
    return n;
  };
  const text = value => String(value || '').trim();

  switch (raw.type) {
    case 'tanks':
    case 'healers': {
      const rows = count(raw.rows, raw.type === 'tanks' ? 2 : 6, 12);
      const labels = raw.type === 'tanks'
        ? ['MT', 'OT'].concat(Array.from({ length: 10 }, (_, i) => `T${i + 3}`)).slice(0, rows)
        : Array.from({ length: rows }, (_, i) => `H${i + 1}`);
      return { type: raw.type, title: uniqueTitle(raw.title, raw.type === 'tanks' ? 'Tank Assignments' : 'Healer Assignments'), rows: rows, labels: labels };
    }
    case 'markers': {
      const markers = parseList_(raw.markers);
      if (!markers.length) throw new Error(`Block ${index + 1}: list the markers, e.g. "Star, Triangle, Square".`);
      if (markers.length > BOSS_BUILDER.MAX_COLUMNS) throw new Error(`Block ${index + 1}: at most ${BOSS_BUILDER.MAX_COLUMNS} markers.`);
      const title = uniqueTitle(raw.title, 'Positions');
      return {
        type: 'markers', title: title, label: `${bossName} ${title}`, markers: markers,
        rows: count(raw.rows, 5, BOSS_BUILDER.MAX_ROWS),
        autoRoles: text(raw.autoRoles), markerRoles: text(raw.markerRoles), ensureEach: text(raw.ensureEach),
        totemParties: !!raw.totemParties, balanceGroups: !!raw.balanceGroups,
        firstHealer: raw.healerRow ? count(raw.firstHealer, 1, 12) : 0,
      };
    }
    case 'list': {
      const title = uniqueTitle(raw.title, 'Assignments');
      return {
        type: 'list', title: title, label: `${bossName} ${title}`, rows: count(raw.rows, 3, BOSS_BUILDER.MAX_ROWS),
        autoRoles: text(raw.autoRoles), preferred: text(raw.preferred), excludeTitles: raw.excludeFrom || [],
        critical: !!raw.critical, selection: raw.selection !== false,
      };
    }
    case 'seats': {
      const columns = parseList_(raw.columns).map(entry => {
        const colon = entry.indexOf(':');
        return colon === -1
          ? { group: '', label: entry }
          : { group: entry.slice(0, colon).trim(), label: entry.slice(colon + 1).trim() };
      });
      if (!columns.length) throw new Error(`Block ${index + 1}: list the seats, e.g. "Siege Engine: Driver, Siege Engine: Gunner".`);
      if (columns.length > BOSS_BUILDER.MAX_COLUMNS) throw new Error(`Block ${index + 1}: at most ${BOSS_BUILDER.MAX_COLUMNS} seat columns.`);
      const title = uniqueTitle(raw.title, 'Seats');
      return {
        type: 'seats', title: title, label: `${bossName} ${title}`, columns: columns,
        rows: count(raw.rows, 5, BOSS_BUILDER.MAX_ROWS),
        autoRoles: text(raw.autoRoles), markerRoles: text(raw.markerRoles),
      };
    }
    default:
      throw new Error(`Block ${index + 1}: unknown block type "${raw.type}".`);
  }
}

/** Turns each pick list's ExcludeFrom titles into the labels of blocks placed before it. */
function resolveExclusions_(blocks) {
  blocks.forEach((block, i) => {
    if (block.type !== 'list') return;
    block.excludeFrom = block.excludeTitles.map(title => {
      const source = blocks.slice(0, i).find(b => b.label && b.title.toLowerCase() === String(title).trim().toLowerCase());
      if (!source) throw new Error(`"${block.title}" excludes "${title}", which is not a block above it.`);
      return source.label;
    });
  });
}

/** Width and height of a block in cells. */
function blockSize_(block) {
  const W = BOSS_BUILDER.CELL_WIDTH;
  if (block.type === 'markers') return { width: W * block.markers.length, height: 2 + (block.firstHealer ? 1 : 0) + block.rows };
  if (block.type === 'seats') {
    const grouped = block.columns.some(c => c.group);
    return { width: 1 + W * block.columns.length, height: (grouped ? 3 : 2) + block.rows };
  }
  return { width: BOSS_BUILDER.LIST_WIDTH, height: 2 + block.rows };
}

/** Packs blocks left to right in bands under the Assignments banner. */
function layoutBlocks_(blocks) {
  const area = BOSS_BUILDER.AREA;
  let column = area.firstColumn;
  let row = area.firstRow;
  let bandHeight = 0;

  blocks.forEach(block => {
    const size = blockSize_(block);
    if (size.width > area.lastColumn - area.firstColumn + 1) throw new Error(`"${block.title}" is too wide for the tab.`);
    if (column + size.width - 1 > area.lastColumn) {
      row += bandHeight + BOSS_BUILDER.GAP_ROWS;
      column = area.firstColumn;
      bandHeight = 0;
    }
    Object.assign(block, { top: row, left: column, bottom: row + size.height - 1, right: column + size.width - 1 });
    column += size.width + BOSS_BUILDER.GAP_COLUMNS;
    bandHeight = Math.max(bandHeight, size.height);
  });
}

/** Copies Supremus, clears its assignments and notes, and sets the title, diagram and link row. */
function copyBossFrame_(ss, plan, config) {
  const template = ss.getSheetByName(BOSS_BUILDER.STYLE_TAB);
  const sheet = template.copyTo(ss).setName(plan.name);

  // Sit with the boss tabs: after the last tab any AutoAssignData row writes to.
  const bossTabs = {};
  config.rows.forEach(row => {
    const named = parseList_(row.get('Sheet')).filter(t => t !== '*' && t.toLowerCase() !== 'all' && t.charAt(0) !== '-');
    (named.length ? named : [String(row.get('Encounter')).trim()]).forEach(tab => { bossTabs[tab] = true; });
  });
  const lastBoss = ss.getSheets().filter(s => bossTabs[s.getName()]).pop();
  ss.setActiveSheet(sheet);
  if (lastBoss) ss.moveActiveSheet(lastBoss.getIndex() + 1);

  const needed = Math.max.apply(null, plan.blocks.map(b => b.bottom)) + 2;
  if (sheet.getMaxRows() < needed) sheet.insertRowsAfter(sheet.getMaxRows(), needed - sheet.getMaxRows());

  sheet.getImages().forEach(image => image.remove());
  sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(protection => protection.remove());
  sheet.setConditionalFormatRules([]);
  sheet.getRange('H1').setValue(plan.name);
  sheet.getRange('A6').setFormula(`=IFNA(XLOOKUP("${plan.name} Map", 'Raid Data'!$J:$J, 'Raid Data'!$K:$K),"")`);

  const filler = template.getRange(BOSS_BUILDER.STYLES.filler);
  const lastRow = sheet.getMaxRows();
  const lastColumn = sheet.getMaxColumns();
  [
    sheet.getRange(30, 1, lastRow - 29, 40),             // A30:AN, notes under the diagram
    sheet.getRange(6, 41, lastRow - 5, lastColumn - 40), // AO6 onwards, the assignment area
  ].forEach(range => {
    range.breakApart();
    range.clearContent().clearNote().clearDataValidations();
    filler.copyTo(range, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  });

  const link = sheet.getRange(BOSS_BUILDER.LINK_ROW, 1, 1, 40);
  template.getRange(BOSS_BUILDER.STYLES.link).copyTo(link, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  link.merge();
  const text = `Raid plan: paste its picture into Raid Data next to "${plan.name} Map"`;
  const rich = SpreadsheetApp.newRichTextValue().setText(plan.raidPlanUrl ? `${text} (open the plan)` : text);
  if (plan.raidPlanUrl) rich.setLinkUrl(plan.raidPlanUrl);
  sheet.getRange(BOSS_BUILDER.LINK_ROW, 1).setRichTextValue(rich.build());
  return sheet;
}

/** Draws one block with Supremus's styles; player cells get a warn-only roster dropdown. */
function drawBlock_(sheet, plan, block, dropdown) {
  const template = sheet.getParent().getSheetByName(BOSS_BUILDER.STYLE_TAB);
  const S = BOSS_BUILDER.STYLES;
  const W = BOSS_BUILDER.CELL_WIDTH;
  const paint = (style, row, column, width, content) => {
    const target = sheet.getRange(row, column, 1, width);
    template.getRange(style).copyTo(target, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    if (width > 1) target.merge();
    const cell = sheet.getRange(row, column);
    if (content && content.formula) cell.setFormula(content.formula);
    else if (content !== undefined && content !== '') cell.setValue(content);
  };
  const players = (row, column, rows) => sheet.getRange(row, column, rows, 1).setDataValidation(dropdown);
  const { top, left } = block;

  if (block.type === 'markers') {
    paint(S.groupIcon, top, left, 1);
    paint(S.groupTitle, top, left + 1, W * block.markers.length - 1, block.title);
    const healers = plan.blocks.find(b => b.type === 'healers');
    block.markers.forEach((marker, i) => {
      const column = left + W * i;
      const icon = BOSS_BUILDER.MARKER_ICONS.indexOf(marker.toLowerCase()) !== -1 ? { formula: `=${marker.toLowerCase()}` } : undefined;
      paint(S.markerIcon, top + 1, column, 1, icon);
      paint(S.markerLabel, top + 1, column + 1, W - 1, marker);
      let row = top + 2;
      if (block.firstHealer) {
        const healerIndex = block.firstHealer - 1 + i;
        const source = healers && healerIndex < healers.rows ? { formula: `=${a1_(healers.top + 2 + healerIndex, healers.left + 6)}` } : undefined;
        paint(S.healerRow, row++, column, W, source);
      }
      for (let r = 0; r < block.rows; r++) paint(S.groupPlayer, row + r, column, W);
      players(row, column, block.rows);
    });
    block.anchorRow = top + 2 + (block.firstHealer ? 1 : 0);
    block.anchorColumn = left;
    return;
  }

  if (block.type === 'seats') {
    paint(S.icon, top, left, 1);
    paint(S.title, top, left + 1, W * block.columns.length, block.title);
    let row = top + 1;
    if (block.columns.some(c => c.group)) {
      paint(S.header, row, left, 1);
      let start = 0;
      block.columns.forEach((column, i) => {
        const next = block.columns[i + 1];
        if (next && next.group === column.group) return;
        paint(S.groupTitle, row, left + 1 + W * start, W * (i - start + 1), column.group);
        start = i + 1;
      });
      row++;
    }
    paint(S.header, row, left, 1, '#');
    block.columns.forEach((column, i) => paint(S.header, row, left + 1 + W * i, W, column.label));
    row++;
    for (let r = 0; r < block.rows; r++) {
      paint(S.label, row + r, left, 1, r + 1);
      block.columns.forEach((column, i) => paint(S.groupPlayer, row + r, left + 1 + W * i, W));
    }
    block.columns.forEach((column, i) => players(row, left + 1 + W * i, block.rows));
    block.anchorRow = row;
    block.anchorColumn = left + 1;
    return;
  }

  // Tanks, healers, pick lists and duties share the Tank Assignments table shape:
  // label | note (5) | player (5).
  paint(S.icon, top, left, 1);
  paint(S.title, top, left + 1, BOSS_BUILDER.LIST_WIDTH - 1, block.title);
  paint(S.header, top + 1, left, 1);
  paint(S.header, top + 1, left + 1, W, 'Assignment');
  paint(S.header, top + 1, left + 1 + W, W, 'Player');
  for (let r = 0; r < block.rows; r++) {
    paint(S.label, top + 2 + r, left, 1, block.labels ? block.labels[r] : r + 1);
    paint(S.note, top + 2 + r, left + 1, W);
    paint(S.player, top + 2 + r, left + 1 + W, W);
  }
  players(top + 2, left + 1 + W, block.rows);
  block.anchorRow = top + 2;
  block.anchorColumn = left + 1 + W;
}

/** <Boss>_tanks / <Boss>_healers over the player cells, so the broadcast fills them. */
function addBlockNamedRanges_(ss, sheet, plan) {
  const taken = {};
  ss.getNamedRanges().forEach(nr => { taken[nr.getName().toLowerCase()] = true; });
  const created = [];
  plan.blocks.filter(b => b.type === 'tanks' || b.type === 'healers').forEach(block => {
    const base = `${plan.prefix}_${block.type === 'tanks' ? 'Tanks' : 'Healers'}`;
    let name = base;
    for (let n = 2; taken[name.toLowerCase()]; n++) name = `${plan.prefix}${n}_${block.type === 'tanks' ? 'Tanks' : 'Healers'}`;
    taken[name.toLowerCase()] = true;
    ss.setNamedRange(name, sheet.getRange(block.anchorRow, block.anchorColumn, block.rows, 1));
    created.push(name);
  });
  return created;
}

/** Adds "<Boss>" (icon) and "<Boss> Map" (diagram) to the Raid Data image table; returns their K cells. */
function addBossImageLabels_(ss, name) {
  const sheet = ss.getSheetByName(BOSS_BUILDER.RAID_DATA_SHEET);
  const firstRow = 27;
  const lastRow = Math.max(sheet.getLastRow(), 68);
  const pairs = sheet.getRange(firstRow, 10, lastRow - firstRow + 1, 2).getValues();

  return [name, `${name} Map`].map(label => {
    let index = pairs.findIndex(([key]) => String(key).trim() === label);
    if (index === -1) {
      // The boss image table ends at row 68; past it, append below everything.
      index = pairs.slice(0, 68 - firstRow + 1).findIndex(([key, image]) => key === '' && image === '');
      if (index === -1) {
        pairs.push(['', '']);
        index = pairs.length - 1;
      }
      pairs[index] = [label, ''];
      sheet.getRange(firstRow + index, 10).setValue(label);
    }
    return `K${firstRow + index}`;
  });
}

/** One AutoAssignData row per filled block, appended below the existing rows. */
function appendBlockRows_(config, plan) {
  const rows = plan.blocks.filter(b => b.label).map(block => {
    const row = new Array(config.width).fill('');
    const set = (name, value) => {
      const index = config.column(name);
      if (index !== undefined) row[index] = value;
    };

    if (block.type === 'duty') {
      // Keep the "*" row's settings and formulas (Preferred reads Quick Assigns).
      block.source.values.forEach((value, i) => { row[i] = block.source.formulas[i] || value; });
    }
    set('Encounter', block.label);
    set('Sheet', plan.name);
    set('Raid', plan.raid);
    set('CellAnchor', a1_(block.anchorRow, block.anchorColumn));
    set('GroupSize', block.rows);

    if (block.type === 'markers') {
      set('GroupAmount', block.markers.length);
      set('ColumnStep', BOSS_BUILDER.CELL_WIDTH);
      set('MarkerNames', block.markers.join(', '));
      set('MarkerRoles', block.markerRoles);
      set('AutoRoles', block.autoRoles);
      set('EnsureEach', block.ensureEach);
      set('TotemParties', block.totemParties);
      set('BalanceGroups', block.balanceGroups);
    } else if (block.type === 'seats') {
      set('GroupAmount', block.columns.length);
      set('ColumnStep', BOSS_BUILDER.CELL_WIDTH);
      set('MarkerNames', block.columns.map(c => [c.group, c.label].filter(Boolean).join(' ')).join(', '));
      set('MarkerRoles', block.markerRoles);
      set('AutoRoles', block.autoRoles);
    } else if (block.type === 'list') {
      set('GroupAmount', 1);
      set('AutoRoles', block.autoRoles);
      set('Preferred', block.preferred);
      set('ExcludeFrom', block.excludeFrom.join(', '));
      set('Selection', block.selection);
      set('Critical', block.critical);
    } else {
      set('GroupAmount', 1);
    }
    return row;
  });
  if (!rows.length) return;

  const last = Math.max(config.sheet.getLastRow(), config.header.row);
  config.sheet.insertRowsAfter(last, rows.length);
  config.sheet.getRange(last + 1, 1, rows.length, config.width).setValues(rows);
}

/** Adds "-<Boss>" to every "*" row, since their fixed cells don't exist on this tab. */
function excludeTabFromWildcardRows_(config, name) {
  const column = config.column('Sheet');
  if (column === undefined) return 0;
  let changed = 0;
  config.rows.forEach(row => {
    const value = row.get('Sheet');
    const entries = parseList_(value);
    const wildcard = entries.some(t => t === '*' || t.toLowerCase() === 'all');
    const excluded = entries.some(t => t.charAt(0) === '-' && t.slice(1).trim().toLowerCase() === name.toLowerCase());
    if (!wildcard || excluded || row.formula('Sheet')) return;
    config.sheet.getRange(row.sheetRow, column + 1).setValue(`${String(value).trim()}, -${name}`);
    changed++;
  });
  return changed;
}
