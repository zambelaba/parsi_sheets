/**
 * Assigns one AutoAssignData row on one tab. Returns one of:
 *   { skipped: reason }      the row is off tonight (SkipWhen, nobody eligible)
 *   { problem: item }        nothing was written; see pushProblem_
 *   { names, warnings }      what was written, and what is worth saying about it
 * One bad row never stops the rest of the run.
 */
function assignEncounter_(run, encounter) {
  const block = resolveBlock_(run, encounter);
  if (block.skipped || block.problem) return block;

  const warnings = [];
  const prePlaced = prePlacedNames_(block.sheet, block.ranges, block.owned, block.snapshot);
  const eligible = buildPool_(run, encounter, prePlaced, warnings);
  if (!eligible.pool.length) return emptyPoolOutcome_(encounter, eligible);

  const groups = buildGroups_(run, encounter, eligible.pool, block, warnings);

  const written = writeGroups_(block.ranges, groups, block.snapshot, encounter.encounter);
  if (written.rejected.length) {
    const one = written.rejected.length === 1;
    const names = written.rejected.slice(0, 6).join(', ') +
      (written.rejected.length > 6 ? ` and ${written.rejected.length - 6} more` : '');
    return { problem: {
      text: `a dropdown here would not accept ${names}, so this block was left as it was`,
      fix: `${one ? 'That name is' : 'Those names are'} missing from the list of options the cell allows. ` +
        `Click the cell, check Data > Data validation, and add ${one ? 'the name' : 'them'} to its list.`,
    } };
  }

  // Once per run per range: a wildcard row would otherwise re-check it on every tab.
  const writeBackKey = encounter.writeBackTo.toLowerCase();
  if (writeBackKey && !run.writtenBack[writeBackKey]) {
    run.writtenBack[writeBackKey] = true;
    writeBackPicks_(run.ss, encounter.writeBackTo, eligible.pool, warnings, run.writeBackFilled);
  }

  written.overflow.forEach(o => warnings.push({
    text: `${o.dropped.join(', ')} had nowhere to go in ${markerLabel_(encounter, o.group)}. ` +
      `It only has room for ${o.capacity}`,
    fix: 'That table needs more rows before everyone fits.',
  }));

  const names = [];
  groups.forEach(group => group.forEach(player => names.push(player.player)));
  return { names: names, warnings: warnings };
}

/**
 * The tab and target ranges for this row, as { sheet, ranges, owned, snapshot },
 * or { skipped } / { problem } when the row cannot run.
 */
function resolveBlock_(run, encounter) {
  const sheet = encounter.sheet || run.ss.getSheetByName(encounter.sheetName);
  if (!sheet) {
    return { problem: {
      internal: true,
      text: `there is no tab called "${encounter.sheetName}"`,
      fix: 'Make Encounter match the tab name exactly, or add a Sheet column naming the tab. Set GroupAmount to 0 to skip it.',
    } };
  }

  if (encounter.skipWhen) {
    const skip = checkSkipWhen_(sheet, encounter, run.previousLabel);
    if (skip) return skip;
  }

  let ranges = encounter.ranges;
  if (!ranges) {
    try {
      ranges = resolveTargetRanges_(sheet, encounter);
    } catch (err) {
      return { problem: {
        internal: true,
        text: `the CellAnchor "${encounter.anchor}" could not be read`,
        fix: 'Use a single cell like BC10, or one per group: BC10, BH10, BM10.',
      } };
    }
  }

  const snapshot = run.sheets.snapshots[encounter.sheetName] || null;
  const formulaCell = findFormulaCell_(ranges, snapshot);
  if (formulaCell) {
    return { problem: {
      internal: true,
      text: `${formulaCell.cell} contains a formula, so none of this block was written`,
      fix: formulaCell.onAnchorRow
        ? 'CellAnchor is probably one row too high. Point it at the first player slot, below the healer.'
        : `That cell is inside the block but is not a plain dropdown. Clear the formula in ${formulaCell.cell}, or shorten GroupSize so the block stops above it.`,
    } };
  }

  return { sheet: sheet, ranges: ranges, owned: run.sheets.owned[encounter.sheetName] || {}, snapshot: snapshot };
}

/**
 * Seats the pool into groups (Preferred, InheritFrom, EnsureEach, TotemParties,
 * then everyone else), applies LimitTo, and notes what is worth reporting.
 */
function buildGroups_(run, encounter, pool, block, warnings) {
  let inherited = null;
  if (encounter.inheritFrom) {
    const result = readInheritedGroups_(run.ss, run.encounters, encounter.inheritFrom);
    if (result.error) {
      warnings.push({
        internal: true,
        text: `nothing was inherited: ${result.error}`,
        fix: 'Rows sharing a tab need a Sheet column naming it, with Encounter as their label.',
      });
    } else {
      inherited = result;
    }
  }

  const groups = encounter.groupSizes.map(() => []);
  const assigned = new Set();

  const missing = seedPreferred_(groups, encounter, run.playersByName, assigned);
  if (missing.length) {
    const one = missing.length === 1;
    warnings.push({
      info: true,
      text: `${missing.join(', ')} ${one ? 'is' : 'are'} picked for this job but ` +
        `${one ? 'is' : 'are'} not raiding tonight, so it was filled automatically instead`,
      fix: 'Nothing to do unless you expected them to be in.',
    });
  }

  const kept = seedFromInheritance_(groups, encounter, inherited, pool, assigned);
  if (encounter.ensureEach.length) seedEnsureEach_(groups, encounter, pool, assigned);
  if (encounter.totemParties) {
    const healersAbove = markerHealers_(block.sheet, block.ranges, block.owned, block.snapshot);
    seedTotemParties_(run, encounter, groups, pool, assigned, healersAbove, warnings);
  }
  fillRemaining_(groups, encounter, pool, assigned);

  noteUncoveredGroups_(groups, encounter, warnings);
  if (inherited && !kept && !run.skippedRows[rowKey_(encounter.inheritFrom, encounter.sheetName)]) {
    warnings.push({
      info: true,
      text: `people did not keep their positions from "${encounter.inheritFrom}". ` +
        'Nothing is filled in there, so fresh spots were picked here',
      fix: 'Normal if that part has not been assigned yet.',
    });
  }
  if (encounter.limitTo) applyLimitTo_(groups, encounter, run, warnings);
  noteFallbacksInMainSlots_(groups, encounter, warnings);
  return groups;
}

/** { skipped } or { problem } when SkipWhen stops this row, otherwise null. */
function checkSkipWhen_(sheet, encounter, previousLabel) {
  // A write refused earlier in the run can surface on the next read, which is
  // this one. Flushing first blames the right row.
  try {
    SpreadsheetApp.flush();
  } catch (err) {
    return { problem: {
      internal: true,
      text: 'a write from an earlier assignment in this run failed to save, and only ' +
        `surfaced here when this SkipWhen cell was read: ${err.message}`,
      fix: `This is not about ${encounter.skipWhen.cell}. Check ` +
        `${previousLabel ? `"${previousLabel}", the assignment run just before this one` : 'the assignment run just before this one'}. ` +
        'A name it tried to write was refused by a dropdown.',
    } };
  }

  const check = evaluateCondition_(sheet, encounter.skipWhen);
  if (check.error) {
    return { problem: {
      internal: true,
      text: `SkipWhen ${check.error}`,
      fix: 'Write it as a cell and a value, e.g. "AE2 = Good DPS".',
    } };
  }
  return check.matched ? { skipped: `${encounter.skipWhen.cell} is "${check.actual}"` } : null;
}

/**
 * Everyone this row may use, most preferred first: AutoRoles matches, minus
 * people ExcludeFrom rows placed and people already in the group from above.
 */
function buildPool_(run, encounter, prePlaced, warnings) {
  let pool = run.players
    .map(player => ({ player: player, rank: eligibilityRank_(player, encounter.eligibleRoles) }))
    .filter(entry => entry.rank !== -1)
    .sort((a, b) => a.rank - b.rank)
    .map(entry => entry.player);

  // An entry matching nobody is only suspicious when it is not a real term either.
  const unknown = encounter.eligibleRoles.filter(entry =>
    !run.players.some(p => eligibilityRank_(p, [entry]) === 0) && !isKnownTerm_(run.knownTerms, entry));
  if (unknown.length) {
    warnings.push({
      internal: true,
      text: `AutoRoles lists ${unknown.map(e => `"${e}"`).join(', ')}, which is not a class, spec or role`,
      fix: 'Check the spelling against Class/Spec Data: columns A, B and F list every valid term.',
    });
  }

  const matched = pool.map(p => p.player);

  const busy = [];
  encounter.excludeFrom.forEach(label => {
    const names = run.placed[rowKey_(label, encounter.sheetName)];
    if (names === undefined) {
      warnings.push({
        internal: true,
        text: `ExcludeFrom "${label}" did not run before this row, so nobody was excluded`,
        fix: 'Move that row above this one in AutoAssignData, since the exclusion reads what it assigned.',
      });
      return;
    }
    names.forEach(name => busy.push(name.toLowerCase()));
  });

  const removedAsBusy = [];
  const removedAsPlaced = [];
  pool = pool.filter(p => {
    const key = p.player.toLowerCase();
    if (busy.indexOf(key) !== -1) {
      removedAsBusy.push(p.player);
      return false;
    }
    if (prePlaced.has(key)) {
      removedAsPlaced.push(`${p.player} in ${prePlaced.get(key)}`);
      return false;
    }
    return true;
  });

  Logger.log(`${encounter.encounter} on ${encounter.sheetName}: pool [${pool.map(p => p.player).join(', ') || 'EMPTY'}]`);
  return { pool: pool, matched: matched, removedAsBusy: removedAsBusy, removedAsPlaced: removedAsPlaced };
}

/** Why the pool came out empty, as { skipped } or { problem }. */
function emptyPoolOutcome_(encounter, eligible) {
  const roles = encounter.eligibleRoles;
  const wanted = roles.length ? ` "${roles.join(', ')}"` : '';
  const needs = roles.length ? roles.map(stripMatchPrefix_).join(' or ') : 'anyone';

  if (!eligible.matched.length) {
    if (!encounter.critical) return { skipped: `nobody on the roster matches AutoRoles${wanted}` };
    return { problem: {
      text: `no ${needs} is rostered, and this fight needs one`,
      fix: 'Get someone who can do it into the raid, or go in knowing this is uncovered.',
    } };
  }

  if (eligible.removedAsBusy.length) {
    const who = eligible.removedAsBusy.join(', ');
    const claimedBy = encounter.excludeFrom.join(', ');
    if (encounter.selection || !encounter.critical) {
      return { skipped: `${who} match AutoRoles${wanted}, but ${claimedBy} already used them — no spares this week` };
    }
    return { problem: {
      text: `${who} could cover this, but ${claimedBy} already used them. Nobody is free, and this fight needs it`,
      fix: 'Roster another one, or take somebody off that other job.',
    } };
  }

  return { problem: {
    internal: true,
    text: `${eligible.removedAsPlaced.join(', ')} matched AutoRoles${wanted}, but that cell ` +
      'sits just above this block, so they count as already placed',
    fix: 'Clear that cell, move the block down, or point CellAnchor somewhere the name is not directly above.',
  } };
}

/** "any mage" -> "mage": the prefixes change matching, not what the thing is called. */
function stripMatchPrefix_(entry) {
  return entry.replace(/^(any|flex)\s+/i, '');
}

/**
 * Reports groups left without an EnsureEach match. Archimonde's decurse table
 * covers its markers from elsewhere, so it is only logged there.
 */
function noteUncoveredGroups_(groups, encounter, warnings) {
  if (!encounter.ensureEach.length) return;

  const uncovered = groups
    .map((group, i) => group.some(p => eligibilityRank_(p, encounter.ensureEach) !== -1) ? null : markerLabel_(encounter, i))
    .filter(label => label !== null);
  if (!uncovered.length) return;

  const wanted = encounter.ensureEach.map(stripMatchPrefix_).join(' or ');
  if (encounter.sheetName === ARCHIMONDE_DECURSE.SHEET) {
    Logger.log(`${encounter.encounter}: no ${wanted} standing in ${uncovered.join(', ')}, covered by the decurse table`);
    return;
  }
  warnings.push({
    info: !encounter.critical,
    text: `no ${wanted} in ${uncovered.join(', ')}`,
    fix: encounter.critical
      ? 'Not enough of them raiding to cover every group, and this row is marked Critical.'
      : 'Not enough of them raiding to cover every group.',
  });
}

/**
 * Never places more people than the LimitTo row did: two Power Infusion targets
 * with one priest raiding would promise a buff nobody casts.
 */
function applyLimitTo_(groups, encounter, run, warnings) {
  const key = rowKey_(encounter.limitTo, encounter.sheetName);
  const source = run.placed[key];

  if (source === undefined) {
    // A skipped source row means nobody can do that job tonight: the limit is zero.
    if (run.skippedRows[key]) {
      groups.forEach(group => { group.length = 0; });
      return;
    }
    warnings.push({
      internal: true,
      text: `LimitTo "${encounter.limitTo}" did not run before this row, so no limit was applied`,
      fix: 'Move that row above this one in AutoAssignData, since the limit reads what it assigned.',
    });
    return;
  }

  const cap = source.length;
  const before = groups.reduce((n, group) => n + group.length, 0);
  let remaining = cap;
  groups.forEach(group => {
    group.length = Math.max(0, Math.min(group.length, remaining));
    remaining -= group.length;
  });

  if (before > cap && !encounter.selection) {
    warnings.push({
      info: true,
      text: `${before - cap} slot(s) left empty: only ${cap} "${encounter.limitTo}" to go round`,
      fix: 'Deliberate: better an empty cell than promising something nobody can cast.',
    });
  }
}

/**
 * With BackupRows, flags later-choice AutoRoles picks sitting in an active slot:
 * there were not enough first-choice players for the slots doing the mechanic.
 */
function noteFallbacksInMainSlots_(groups, encounter, warnings) {
  if (!encounter.backupRows || encounter.eligibleRoles.length < 2) return;

  const found = [];
  let mainSlotTotal = 0;
  groups.forEach((group, i) => {
    const mainSlots = encounter.groupSizes[i] - encounter.backupRows;
    if (mainSlots <= 0) return;
    mainSlotTotal += mainSlots;
    group.slice(0, mainSlots).forEach(player => {
      if (eligibilityRank_(player, encounter.eligibleRoles) > 0) {
        found.push(`${player.player} (${player.classspec || player.role})`);
      }
    });
  });
  if (!found.length) return;

  const firstChoice = encounter.eligibleRoles[0];
  warnings.push({
    info: true,
    text: `not enough ${firstChoice} to fill the main slots, so ${found.join(', ')} ` +
      `${found.length === 1 ? 'is' : 'are'} in an active slot rather than the backup row`,
    fix: `Fine if you expect it. ${mainSlotTotal} main slots need that many ${firstChoice}.`,
  });
}

/** "Star" when MarkerNames is set, otherwise "group 3". */
function markerLabel_(encounter, index) {
  return encounter.markerNames[index] || `group ${index + 1}`;
}
