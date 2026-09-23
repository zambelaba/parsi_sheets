/**
 * Archimonde's Decurse Assignments table: Star, Triangle, Square, Tank Group,
 * Melee Group (in priority order).
 *
 * Star / Triangle / Square take the decurser actually standing in that marker.
 * Every other slot goes to the decurser covering the fewest groups so far, ties
 * going to whoever ranks higher in the Archimonde row's EnsureEach. Unlike any
 * other table, one name may repeat here: a decurse is cheap enough to cover
 * several groups.
 */
const ARCHIMONDE_DECURSE = Object.freeze({
  SHEET: 'Archimonde',
  MARKER_RANGES: ['BC10:BC14', 'BH10:BH14', 'BM10:BM14'],
  TARGET_CELLS: ['BH17', 'BH18', 'BH19', 'BH20', 'BH21'],
  // Used only when no Archimonde row has an EnsureEach.
  FALLBACK_DECURSERS: ['any mage', 'any druid'],
});

function assignArchimondeDecurse_(ss, players, encounters, problems, notices) {
  const sheet = ss.getSheetByName(ARCHIMONDE_DECURSE.SHEET);
  if (!sheet) return;

  const source = encounters.filter(e => e.sheetName === ARCHIMONDE_DECURSE.SHEET && e.ensureEach.length)[0];
  const entries = source ? source.ensureEach : ARCHIMONDE_DECURSE.FALLBACK_DECURSERS;
  const rankOf = player => eligibilityRank_(player, entries);
  const isDecurser = player => rankOf(player) !== -1;
  const byName = new Map(players.map(p => [p.player.toLowerCase(), p]));

  const standing = ARCHIMONDE_DECURSE.MARKER_RANGES.map(a1 => {
    const decurser = sheet.getRange(a1).getDisplayValues().flat()
      .map(value => String(value).trim())
      .filter(Boolean)
      .map(name => byName.get(name.toLowerCase()))
      .find(player => player && isDecurser(player));
    return decurser ? decurser.player : null;
  });

  const pool = players.filter(isDecurser).sort((a, b) => rankOf(a) - rankOf(b));
  if (!pool.length) {
    problems.push({
      where: 'Archimonde Decurse Assignments',
      text: 'no Mage, Balance Druid or Restoration Druid is rostered, so nothing can decurse',
      fix: 'This fight cannot be safely done without a decurser. Get one rostered or reassign specs before the pull.',
    });
    return;
  }

  // Real placements count first, so nobody standing in a marker is handed a
  // second group while another decurser sits idle.
  const load = new Map(pool.map(p => [p.player, 0]));
  standing.forEach(name => {
    if (name && load.has(name)) load.set(name, load.get(name) + 1);
  });
  const leastLoaded = () => {
    let pick = pool[0];
    pool.forEach(p => { if (load.get(p.player) < load.get(pick.player)) pick = p; });
    load.set(pick.player, load.get(pick.player) + 1);
    return pick.player;
  };

  const cells = ARCHIMONDE_DECURSE.TARGET_CELLS;
  const values = cells.map((cell, i) => standing[i] || leastLoaded());

  if (pool.length < cells.length) {
    const covers = {};
    values.forEach(name => { covers[name] = (covers[name] || 0) + 1; });
    const stretched = Object.keys(covers).filter(name => covers[name] > 1).map(name => `${name} has ${covers[name]}`);
    notices.push({
      info: true,
      text: `Only ${pool.length} decurser${pool.length === 1 ? '' : 's'} for ${cells.length} groups on Archimonde` +
        (stretched.length ? `, so ${stretched.join(' and ')}` : ''),
    });
  }

  const blocked = [];
  cells.forEach((cell, i) => {
    const range = sheet.getRange(cell);
    if (range.getFormula()) blocked.push(cell);
    else range.setValue(values[i]);
  });
  SpreadsheetApp.flush(); // a refused dropdown is reported against this table

  if (blocked.length) {
    pushProblem_(problems, 'Archimonde Decurse Assignments', {
      internal: true,
      text: `${blocked.join(', ')} still ${blocked.length === 1 ? 'has' : 'have'} a formula in it, so nothing was written there`,
      fix: 'Clear the formula in that cell and re-run to have it auto-filled too.',
    });
  }
}
