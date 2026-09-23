/**
 * Auto assignments: one run fills every boss sheet from the AutoAssignData config.
 *
 * A run, in order:
 *   1. Refreshes TANKS_GLOBAL / HEALERS_GLOBAL against tonight's roster.
 *   2. Broadcasts them to every *_tanks / *_healers range (assignHealersAndTanks).
 *   3. Assigns every AutoAssignData row in sheet order, so a row can read what
 *      the rows above it placed (ExcludeFrom, LimitTo, InheritFrom).
 *   4. Fills the Archimonde decurse table and opens the T6 export sidebar.
 */

const AUTO_ASSIGN = Object.freeze({
  TITLE: 'Auto Assignments Overview',
  ENCOUNTERS_SHEET: 'AutoAssignData',
  ROLE_MAP_SHEET: 'Class/Spec Data',
  RAIDER_DATA_SHEET: 'Raider Data',
  PARTIES_SHEET: 'Dynamic Lists',
  TANKS_RANGE: 'TANKS_GLOBAL',
  HEALERS_RANGE: 'HEALERS_GLOBAL',
  TANK_PRIORITY_RANGE: 'TANK_PRIORITY',
  HEALER_PRIORITY_RANGE: 'HEALER_PRIORITY',
  DEFAULT_TANK_PRIORITY: ['Tank'],
  DEFAULT_HEALER_PRIORITY: ['Healer'],
  PARTY_COUNT: 5,
  HEADER_SCAN_ROWS: 10,
  // Rows above an anchor that belong to its group (the healer rows). Kept small
  // so a table stacked higher up the same columns is not mistaken for it.
  ABOVE_ANCHOR_ROWS: 4,
  // Above this many cells, a sheet's blocks are read one by one instead of as
  // one bounding-box snapshot.
  MAX_SNAPSHOT_CELLS: 40000,
});

function autoAssignGroups() {
  runAutoAssignSafely_(null);
}

function autoAssignMountHyjal() {
  runAutoAssignSafely_('mh');
}

function autoAssignBlackTemple() {
  runAutoAssignSafely_('bt');
}

function runAutoAssignSafely_(scope) {
  const ss = SpreadsheetApp.getActive();
  try {
    runAutoAssign_(ss, scope);
  } catch (err) {
    const message = String((err && err.message) || err);
    const transient = /service unavailable|internal error|try again|timed out|too many/i.test(message);
    // Not "stopped before writing": some sheets may already be assigned.
    showResult_(ss, {
      ok: false,
      headline: transient ? 'Google was briefly unavailable' : 'Stopped: setup problem',
      issues: [{
        text: message,
        fix: transient
          ? 'Nothing to fix on your side. Run it again — any sheets already done will simply be rewritten.'
          : undefined,
      }],
    });
  }
}

function runAutoAssign_(ss, scope) {
  const started = new Date();
  announce_(ss, 'Starting — reading your roster…', 30);

  const raiders = readRaiders_(ss);
  if (!raiders.length) {
    showResult_(ss, {
      ok: false,
      headline: 'No Roster Setup!',
      issues: [{ text: 'Run "Instructions" in Roster.' }],
      action: { label: 'Open Instructions', fn: 'openRosterInstructions' },
    });
    return;
  }

  const roleMap = loadRoleMap_(ss);
  const pickNotices = refreshTankAndHealerPicks_(ss, buildRoster_(raiders, roleMap, new Set(), new Set()));
  const tankPicks = namesInRange_(ss, AUTO_ASSIGN.TANKS_RANGE);
  const healerPicks = namesInRange_(ss, AUTO_ASSIGN.HEALERS_RANGE);

  const missingPicks = describeMissingPicks_(tankPicks, healerPicks);
  if (missingPicks.length) {
    showResult_(ss, {
      ok: false,
      headline: 'Stopped before writing: tanks and healers are not set',
      issues: missingPicks.map(text => ({
        text: text,
        fix: 'Pick your tanks and healers in Quick Assigns, then run this again.',
      })),
      notes: pickNotices.map(notice => notice.text).concat([
        'The healer row above each group is filled by formula from those picks, so without them every group would be built on blank input.',
      ]),
    });
    return;
  }

  // The healer rows above each group are formulas reading these tables, so they
  // must be current before any group is built.
  const setupNotes = [];
  try {
    assignHealersAndTanks();
    SpreadsheetApp.flush();
  } catch (err) {
    setupNotes.push(`Could not run assignHealersAndTanks: ${err.message}`);
  }

  const encounters = loadEncounters_(ss, scope);
  if (!encounters.length) {
    showResult_(ss, {
      ok: false,
      headline: 'Nothing to do',
      issues: [{
        text: `no boss rows were found in the "${AUTO_ASSIGN.ENCOUNTERS_SHEET}" sheet`,
        fix: 'Add a row per boss with Encounter, CellAnchor, GroupAmount and GroupSize filled in.',
      }],
    });
    return;
  }

  const players = buildRoster_(raiders, roleMap, tankPicks, healerPicks);
  const run = {
    ss: ss,
    encounters: encounters,
    knownTerms: roleMap.terms,
    players: players,
    playersByName: new Map(players.map(p => [p.player.toLowerCase(), p])),
    parties: loadPartyMembership_(ss),
    sheets: prepareSheets_(ss, encounters),
    placed: {},        // row key -> names that row placed
    skippedRows: {},   // row key -> true when its SkipWhen matched
    writtenBack: {},   // WriteBackTo range -> true once handled this run
    writeBackFilled: [],
    previousLabel: null,
  };
  Logger.log(`Roles used for matching: ${players.map(p =>
    `${p.player}=${p.classspec || '?'}/${p.role || 'no role'}`).join(', ')}`);

  const active = encounters.filter(e => e.groupCount && e.anchor);
  announce_(ss,
    `Assigning ${active.length} rows — this takes a few minutes. ` +
    `Please don't edit the sheet until it finishes.`, 300);

  const problems = [];
  const notices = pickNotices.slice();
  const assignedCount = assignRows_(run, active, problems, notices);

  if (active.some(e => e.sheetName === ARCHIMONDE_DECURSE.SHEET)) {
    try {
      assignArchimondeDecurse_(ss, players, encounters, problems, notices);
    } catch (err) {
      pushProblem_(problems, 'Archimonde Decurse Assignments', { internal: true, text: err.message });
    }
  }

  if (run.writeBackFilled.length) {
    Logger.log(`Quick Assigns seeded this run: ${run.writeBackFilled.join(', ')}`);
    notices.push({ info: true, text: 'We auto filled it for you, adjust if needed in Quick Assigns.' });
  }

  // Every export string belongs to Black Temple. Opened before the result
  // dialog so the modal lands on top of it.
  if (!scope || scope === 'bt') openT6ExportHubAfterRun_(ss);

  showRunSummary_(ss, { scope, started, assignedCount, problems, notices, setupNotes });
}

function showRunSummary_(ss, summary) {
  const seconds = ((new Date() - summary.started) / 1000).toFixed(1);
  const scopeLabel = summary.scope === 'mh' ? ' (Mount Hyjal only)' : summary.scope === 'bt' ? ' (Black Temple only)' : '';
  const clean = !summary.problems.length && !summary.setupNotes.length;
  const issues = collapseIssues_(summary.problems);
  const written = `${summary.assignedCount} assignments written in ${seconds}s${scopeLabel}`;

  showResult_(ss, {
    ok: clean,
    headline: clean ? `All done: ${written}` : written + (issues.length ? `, ${issues.length} to look at` : ''),
    issues: issues,
    notices: collapseIssues_(summary.notices),
    notes: summary.setupNotes,
  });
}

/** Assigns each row in order, collecting what to report. Returns how many rows were written. */
function assignRows_(run, rows, problems, notices) {
  const progressEvery = Math.max(1, Math.ceil(rows.length / 10));
  let assignedCount = 0;

  rows.forEach((encounter, index) => {
    const done = index + 1;
    if (done % progressEvery === 0 || done === rows.length) {
      const filled = Math.min(10, Math.round(done / rows.length * 10));
      announce_(run.ss, `${'▓'.repeat(filled)}${'░'.repeat(10 - filled)}  ` +
        `${Math.round(done / rows.length * 100)}%`, 300);
    }

    // A wildcard row runs on every tab, so name the tab when it repeats.
    const where = encounter.repeated ? `${encounter.encounter} on ${encounter.sheetName}` : encounter.encounter;
    const key = rowKey_(encounter.encounter, encounter.sheetName);

    try {
      const outcome = assignEncounter_(run, encounter);
      run.previousLabel = where;

      if (outcome.skipped) {
        run.skippedRows[key] = true;
        Logger.log(`Skipped ${where}: ${outcome.skipped}`);
      } else if (outcome.problem) {
        pushProblem_(problems, where, outcome.problem);
      } else {
        run.placed[key] = outcome.names;
        assignedCount++;
        outcome.warnings.forEach(warning => {
          if (warning.info) notices.push({ where: where, text: warning.text, fix: warning.fix });
          else pushProblem_(problems, where, warning);
        });
      }
    } catch (err) {
      pushProblem_(problems, where, {
        internal: true,
        text: err.message,
        fix: /range not found/i.test(err.message)
          ? `That cell does not exist on this sheet, which may be narrower or shorter than the others. Exclude it with "-${encounter.sheetName}" in Sheet if it has no such table.`
          : undefined,
      });
    }
  });
  return assignedCount;
}

/**
 * Takes anyone not raiding out of the tank and healer picks and refills empty
 * slots from the priority lists. Returns the notices describing what changed.
 */
function refreshTankAndHealerPicks_(ss, roster) {
  const rosteredKeys = new Set(roster.map(p => p.player.toLowerCase()));
  const jobs = [
    { range: AUTO_ASSIGN.TANKS_RANGE, label: 'tank', namesMainTank: true,
      priority: loadPriority_(ss, AUTO_ASSIGN.TANK_PRIORITY_RANGE, AUTO_ASSIGN.DEFAULT_TANK_PRIORITY) },
    { range: AUTO_ASSIGN.HEALERS_RANGE, label: 'healer', namesMainTank: false,
      priority: loadPriority_(ss, AUTO_ASSIGN.HEALER_PRIORITY_RANGE, AUTO_ASSIGN.DEFAULT_HEALER_PRIORITY) },
  ];

  const notices = [];
  jobs.forEach(job => {
    const { removed, added, firstFilled } =
      reconcilePicks_(ss, job.range, rosteredKeys, roster, job.priority);

    if (removed.length) {
      notices.push({
        info: true,
        text: `${removed.join(', ')} ${removed.length === 1 ? 'is' : 'are'} ` +
          `not raiding, so they were taken out of your ${job.label} picks` +
          (added.length ? `. ${added.join(', ')} filled the ${added.length === 1 ? 'gap' : 'gaps'}` : ''),
      });
    } else if (added.length) {
      notices.push({
        info: true,
        text: `Empty ${job.label} slot${added.length === 1 ? '' : 's'} filled with ${added.join(', ')}.`,
      });
    }

    // The first tank cell is the Main Tank everywhere, so a new one is worth naming.
    if (job.namesMainTank && firstFilled) {
      notices.push({
        info: true,
        text: `${firstFilled} was picked as your main tank. Change it in Quick Assigns if that is not who you want.`,
      });
    }
  });
  return notices;
}

function describeMissingPicks_(tanks, healers) {
  const problems = [];
  [[tanks, AUTO_ASSIGN.TANKS_RANGE, 'tanks'], [healers, AUTO_ASSIGN.HEALERS_RANGE, 'healers']]
    .forEach(([names, rangeName, label]) => {
      if (names === null) problems.push(`${rangeName} does not exist`);
      else if (!names.size) problems.push(`no ${label} are selected (${rangeName} is empty)`);
    });
  return problems;
}

/** "label||sheet", the key rows use to find each other's results. */
function rowKey_(label, sheetName) {
  return `${String(label).trim().toLowerCase()}||${String(sheetName).trim().toLowerCase()}`;
}

/** A toast for the few moments worth interrupting someone mid-run. */
function announce_(ss, message, seconds) {
  Logger.log(message);
  ss.toast(message, AUTO_ASSIGN.TITLE, seconds || 30);
}

/**
 * Adds a problem to the result. Setup faults (`internal`) show as one generic
 * line, with the real text in the log for whoever maintains AutoAssignData.
 */
function pushProblem_(problems, where, item) {
  if (!item.internal) {
    problems.push({ where: where, text: item.text, fix: item.fix });
    return;
  }
  Logger.log(`SETUP PROBLEM — ${where}: ${item.text}` + (item.fix ? ` | FIX: ${item.fix}` : ''));
  problems.push({
    where: where,
    text: 'could not be filled in. The sheet needs a setup fix.',
    fix: 'Nothing to do from here. Whoever looks after this sheet will find the details in the script log.',
  });
}

/** Merges identical entries (a wildcard row repeats them per sheet) into one, listing where. */
function collapseIssues_(items) {
  const groups = [];
  const indexByKey = {};

  (items || []).forEach(item => {
    const key = `${item.text}||${item.fix || ''}`;
    if (indexByKey[key] === undefined) {
      indexByKey[key] = groups.length;
      groups.push({ text: item.text, fix: item.fix, places: [] });
    }
    if (item.where) groups[indexByKey[key]].places.push(item.where);
  });

  return groups.map(group => {
    if (group.places.length <= 1) return { where: group.places[0], text: group.text, fix: group.fix };
    const extra = group.places.length > 6 ? ` +${group.places.length - 6} more` : '';
    return {
      where: `${group.places.length}×`,
      text: group.text,
      places: group.places.slice(0, 6).join(', ') + extra,
      fix: group.fix,
    };
  });
}

/**
 * Shows the run result in a dialog, or as a toast when there is no UI (editor
 * runs, triggers).
 *
 * result = { ok, headline, issues[], notices[], notes[], action? }
 */
function showResult_(ss, result) {
  Logger.log(result.headline);
  [['issue', result.issues], ['fyi', result.notices]].forEach(([label, items]) => {
    (items || []).forEach(item => Logger.log(`  ${label}: ${item.where ? item.where + ' — ' : ''}${item.text}` +
      (item.places ? ` [${item.places}]` : '') + (item.fix ? ` (${item.fix})` : '')));
  });
  (result.notes || []).forEach(note => Logger.log(`  note: ${note}`));

  try {
    const template = HtmlService.createTemplateFromFile('AutoAssignResult');
    template.result = result;
    // Ends up inside an onclick attribute, so only word characters survive.
    template.actionFn = result.action ? String(result.action.fn).replace(/[^A-Za-z0-9_]/g, '') : '';
    SpreadsheetApp.getUi().showModalDialog(
      template.evaluate().setWidth(560).setHeight(460), AUTO_ASSIGN.TITLE);
  } catch (err) {
    const flat = [result.headline]
      .concat((result.issues || []).map(i => `${i.where || ''} ${i.text}`.trim()))
      .concat(result.notes || [])
      .join(' | ');
    ss.toast(flat, AUTO_ASSIGN.TITLE, 30);
  }
}
