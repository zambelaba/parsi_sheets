/**
 * Who is raiding and what they can do.
 *
 * Raider Data is the only player list (already filtered by "Is Rostered" on the
 * Roster sheet). Class/Spec Data maps Class + Spec to a Role, and this week's
 * TANKS_GLOBAL / HEALERS_GLOBAL picks override that role.
 */

/** Raider Data rows [name, class, spec, flexSpec] with a name, as displayed. */
function readRaiders_(ss) {
  const sheet = ss.getSheetByName(AUTO_ASSIGN.RAIDER_DATA_SHEET);
  if (!sheet) throw new Error(`Sheet "${AUTO_ASSIGN.RAIDER_DATA_SHEET}" not found`);

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  return sheet.getRange(2, 1, lastRow - 1, 4).getDisplayValues()
    .filter(([name]) => String(name).trim() !== '');
}

/** Players with their effective role. `tanks` / `healers` are lowercased name Sets. */
function buildRoster_(raiders, roleMap, tanks, healers) {
  return raiders.map(([name, cls, spec, flexSpec]) => {
    const player = String(name).trim();
    const key = player.toLowerCase();
    const className = String(cls).trim();
    const specName = String(spec).trim();
    const flexName = String(flexSpec).trim();

    let role = roleMap.roles[roleKey_(className, specName)] || '';
    if (tanks.has(key)) role = 'Tank';
    else if (healers.has(key)) role = 'Healer';

    return {
      player: player,
      class: className,
      spec: specName,
      classspec: specName && className ? `${specName} ${className}` : '',
      role: role,
      flexSpec: flexName,
      flexclassspec: flexName && className ? `${flexName} ${className}` : '',
      flexrole: flexName ? (roleMap.roles[roleKey_(className, flexName)] || '') : '',
    };
  });
}

/**
 * { roles, terms } from Class/Spec Data (A = Class, B = Spec, F = Role).
 * `terms` holds every valid class, spec, "Spec Class" and role, so a misspelled
 * AutoRoles entry can be told apart from a spec nobody plays this week.
 */
function loadRoleMap_(ss) {
  const sheet = ss.getSheetByName(AUTO_ASSIGN.ROLE_MAP_SHEET);
  if (!sheet) throw new Error(`Sheet "${AUTO_ASSIGN.ROLE_MAP_SHEET}" not found`);

  const roles = {};
  const terms = {};
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { roles: roles, terms: terms };

  const remember = value => {
    const key = String(value).trim().toLowerCase();
    if (key) terms[key] = true;
  };

  sheet.getRange(2, 1, lastRow - 1, 6).getValues().forEach(([cls, spec, , , , role]) => {
    if (!cls && !spec) return;
    roles[roleKey_(cls, spec)] = String(role).trim();
    remember(cls);
    remember(spec);
    remember(role);
    if (cls && spec) remember(`${String(spec).trim()} ${String(cls).trim()}`);
  });
  return { roles: roles, terms: terms };
}

function roleKey_(cls, spec) {
  return `${String(cls).trim().toLowerCase()}|${String(spec).trim().toLowerCase()}`;
}

/** Lowercased player name -> raid party (1-5), from Dynamic Lists ActiveRosterG1..G5. */
function loadPartyMembership_(ss) {
  const byPlayer = new Map();
  const sheet = ss.getSheetByName(AUTO_ASSIGN.PARTIES_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return byPlayer;

  let header;
  try {
    header = locateHeaderRow_(sheet, ['ActiveRosterG1'], AUTO_ASSIGN.PARTIES_SHEET);
  } catch (err) {
    return byPlayer; // only TotemParties rows need it, and they report its absence
  }

  const lastRow = sheet.getLastRow();
  if (lastRow <= header.row) return byPlayer;
  const values = sheet.getRange(header.row + 1, 1, lastRow - header.row, sheet.getLastColumn()).getDisplayValues();

  for (let party = 1; party <= AUTO_ASSIGN.PARTY_COUNT; party++) {
    const column = header.index[normalizeHeader_(`ActiveRosterG${party}`)];
    if (column === undefined) continue;
    values.forEach(row => {
      const name = String(row[column]).trim();
      if (name) byPlayer.set(name.toLowerCase(), party);
    });
  }
  return byPlayer;
}

/** Lowercased names in a named range, or null when the range does not exist. */
function namesInRange_(ss, rangeName) {
  const range = ss.getRangeByName(rangeName);
  if (!range) return null;

  const names = new Set();
  range.getDisplayValues().flat().forEach(value => {
    const name = String(value).trim();
    if (name) names.add(name.toLowerCase());
  });
  return names;
}

/** An ordered priority list: one name per cell, or one cell holding a comma list. */
function loadPriority_(ss, rangeName, fallback) {
  const range = ss.getRangeByName(rangeName);
  if (!range) return fallback;

  const entries = range.getDisplayValues().flat().map(v => String(v).trim()).filter(v => v !== '');
  const parsed = entries.length === 1 ? parseList_(entries[0]) : entries;
  return parsed.length ? parsed : fallback;
}

/**
 * Clears anyone not raiding out of a picks range, then fills every empty cell
 * from the priority list, best first. A cell only ever holds a name or nothing.
 *
 * Returns { removed, added, firstFilled }, firstFilled being who landed in the
 * first cell (the Main Tank for TANKS_GLOBAL).
 */
function reconcilePicks_(ss, rangeName, rosteredKeys, candidates, priority) {
  const result = { removed: [], added: [], firstFilled: '' };
  const range = ss.getRangeByName(rangeName);
  if (!range) return result;

  const values = range.getDisplayValues();
  const empty = [];
  const present = new Set();

  values.forEach((row, r) => row.forEach((value, c) => {
    const name = String(value).trim();
    if (name && rosteredKeys.has(name.toLowerCase())) {
      present.add(name.toLowerCase());
      return;
    }
    if (name) {
      values[r][c] = '';
      result.removed.push(name);
    }
    empty.push({ r: r, c: c, isFirst: r === 0 && c === 0 });
  }));

  const wanted = priority.map(p => String(p).trim().toLowerCase());
  const replacements = candidates
    .map(player => ({ player: player, rank: eligibilityRank_(player, wanted) }))
    .filter(entry => entry.rank !== -1 && !present.has(entry.player.player.toLowerCase()))
    .sort((a, b) => a.rank - b.rank);

  empty.forEach(cell => {
    const next = replacements.shift();
    if (!next) return;
    values[cell.r][cell.c] = next.player.player;
    result.added.push(next.player.player);
    if (cell.isFirst) result.firstFilled = next.player.player;
    present.add(next.player.player.toLowerCase());
  });

  if (!result.removed.length && !result.added.length) return result;
  range.setValues(values);
  Logger.log(`${rangeName}: removed [${result.removed.join(', ')}], filled [${result.added.join(', ')}]`);
  return result;
}

/**
 * Copies HEALERS_GLOBAL and TANKS_GLOBAL into every named range ending in
 * _healers / _tanks. Each target starts again from the top of the list, so a
 * one-cell range gets the Main Tank and a two-cell range gets MT, OT.
 */
function assignHealersAndTanks() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const healers = ss.getRangeByName(AUTO_ASSIGN.HEALERS_RANGE);
  const tanks = ss.getRangeByName(AUTO_ASSIGN.TANKS_RANGE);
  if (!healers && !tanks) return;

  const namesIn = range => range ? range.getValues().flat().filter(n => n.toString().trim() !== '') : [];
  const namedRanges = ss.getNamedRanges();
  const broadcast = (suffix, names, message) => {
    const targets = namedRanges.filter(nr => suffix.test(nr.getName()));
    if (!targets.length || !names.length) return;
    targets.forEach(nr => fillRange_(nr.getRange(), names));
    ss.toast(message, '🧹 Auto Assigns Helper', 10);
  };

  broadcast(/_healers$/i, namesIn(healers), 'Main Spec healers assigned to boss pages.');
  broadcast(/_tanks$/i, namesIn(tanks), 'Tank Roles assigned to boss pages.');
}

/**
 * Seeds an empty setup range (Quick Assigns) with everyone eligible, so next
 * run they come back through Preferred as a visible, editable choice. Never
 * touches a range that already holds a name or a formula.
 */
function writeBackPicks_(ss, rangeName, candidates, warnings, filled) {
  const range = ss.getRangeByName(rangeName);
  if (!range) {
    warnings.push({
      internal: true,
      text: `WriteBackTo names "${rangeName}", which is not a named range`,
      fix: 'Create it on your setup page (Data > Named ranges), or clear the WriteBackTo cell.',
    });
    return;
  }

  if (range.getDisplayValues().flat().some(value => String(value).trim() !== '')) return;

  if (range.getFormulas().flat().some(formula => formula !== '')) {
    warnings.push({
      internal: true,
      text: `"${rangeName}" holds a formula, so nothing was written back to it`,
      fix: 'Point WriteBackTo at plain cells people can type in, or drop the column.',
    });
    return;
  }

  const names = candidates.slice(0, range.getNumRows() * range.getNumColumns()).map(p => p.player);
  if (!names.length) return;

  try {
    fillRange_(range, names);
    SpreadsheetApp.flush(); // a refused dropdown throws here, not on some later row
  } catch (err) {
    const one = names.length === 1;
    warnings.push({
      text: `a dropdown on "${rangeName}" would not accept ${names.join(', ')}`,
      fix: `${one ? 'That name is' : 'Those names are'} missing from the list of options those cells allow. ` +
        `Click one, check Data > Data validation, and add ${one ? 'the name' : 'them'} to its list.`,
    });
    return;
  }
  filled.push(`${rangeName}: ${names.join(', ')}`);
}

/** Writes names into a range left to right, top to bottom, blanking the rest. */
function fillRange_(range, names) {
  let index = 0;
  const values = [];
  for (let r = 0; r < range.getNumRows(); r++) {
    const row = [];
    for (let c = 0; c < range.getNumColumns(); c++) row.push(names[index++] || '');
    values.push(row);
  }
  range.setValues(values);
}
