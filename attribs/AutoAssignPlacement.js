/**
 * Matching players against AutoRoles-style lists, and building groups.
 *
 * Groups are filled in a fixed order, each step only using people nobody has
 * placed yet: Preferred names, InheritFrom, EnsureEach, TotemParties, then
 * everyone left over.
 */

/**
 * Index of the first entry this player matches (0 = most preferred), or -1.
 * Entries (lowercased) name a Role, Class or "Spec Class". A bare Class means a
 * DPS of that class; "any <class>" includes its tanks and healers; "flex <x>"
 * matches the off-spec instead. No entries means everyone matches equally.
 */
function eligibilityRank_(player, entries) {
  if (!entries.length) return 0;

  const role = player.role.toLowerCase();
  const isSupport = role === 'tank' || role === 'healer';

  for (let i = 0; i < entries.length; i++) {
    const entry = entries[i];

    if (entry.indexOf('flex') === 0) {
      const wanted = entry.slice(4).trim();
      if (looseEquals_(wanted, player.flexrole) || looseEquals_(wanted, player.flexclassspec) ||
          looseEquals_(wanted, player.flexSpec)) return i;
      continue;
    }

    const anySpec = entry.indexOf('any ') === 0;
    const wanted = anySpec ? entry.slice(4).trim() : entry;
    if (looseEquals_(wanted, role) || looseEquals_(wanted, player.classspec)) return i;
    if (looseEquals_(wanted, player.class) && (anySpec || !isSupport)) return i;
  }
  return -1;
}

/** Case-insensitive, and tolerant of a plural "s" ("Healers", "Shamans"). */
function looseEquals_(a, b) {
  const x = String(a).trim().toLowerCase();
  const y = String(b).trim().toLowerCase();
  return x === y || singular_(x) === singular_(y);
}

function singular_(text) {
  return text.length > 1 && text.slice(-1) === 's' ? text.slice(0, -1) : text;
}

/** Whether an AutoRoles entry names a real class, spec or role (ignoring "any"/"flex"). */
function isKnownTerm_(terms, entry) {
  let bare = entry;
  if (bare.indexOf('flex') === 0) bare = bare.slice(4).trim();
  else if (bare.indexOf('any ') === 0) bare = bare.slice(4).trim();
  return !!(terms[bare] || terms[singular_(bare)]);
}

/** A marker's soft preference. A blank preference suits nobody. */
function markerSuits_(player, markerRole) {
  if (!markerRole) return false;
  return eligibilityRank_(player, [String(markerRole).trim().toLowerCase()]) === 0;
}

/**
 * Places the Preferred names in order, group by group. Being named is the
 * qualification, so AutoRoles is not checked. Returns the names not raiding.
 */
function seedPreferred_(groups, encounter, playersByName, assigned) {
  const missing = [];
  let groupIndex = 0;

  encounter.preferred.forEach(name => {
    const player = playersByName.get(name.trim().toLowerCase());
    if (!player) {
      missing.push(name);
      return;
    }
    if (assigned.has(player.player) && !encounter.allowDuplicates) return;

    while (groupIndex < groups.length && groups[groupIndex].length >= encounter.groupSizes[groupIndex]) {
      groupIndex++;
    }
    if (groupIndex >= groups.length) return;

    groups[groupIndex].push(player);
    assigned.add(player.player);
  });
  return missing;
}

/**
 * Keeps players on the marker they hold in the InheritFrom row, matched by
 * marker name when both rows have MarkerNames, otherwise by position.
 * Returns how many were kept.
 */
function seedFromInheritance_(groups, encounter, inherited, pool, assigned) {
  if (!inherited) return 0;

  const byName = new Map(pool.map(p => [p.player.toLowerCase(), p]));
  const sourceNames = inherited.markerNames.map(n => n.trim().toLowerCase());
  const byMarkerName = encounter.markerNames.length && inherited.markerNames.length;
  let kept = 0;

  groups.forEach((group, i) => {
    const source = byMarkerName
      ? sourceNames.indexOf((encounter.markerNames[i] || '').trim().toLowerCase())
      : i;
    if (source === -1 || !inherited.groups[source]) return;

    inherited.groups[source].forEach(name => {
      const player = byName.get(name.toLowerCase());
      if (!player || assigned.has(player.player) || group.length >= encounter.groupSizes[i]) return;
      group.push(player);
      assigned.add(player.player);
      kept++;
    });
  });
  return kept;
}

/**
 * Gives every group one EnsureEach match before anything optional is placed,
 * neediest (then smallest) group first, taking candidates in pool order.
 */
function seedEnsureEach_(groups, encounter, pool, assigned) {
  const covers = player => eligibilityRank_(player, encounter.ensureEach) !== -1;

  for (let guard = 0; guard <= groups.length * pool.length; guard++) {
    const needy = groups
      .map((group, i) => i)
      .filter(i => !groups[i].some(covers) && groups[i].length < encounter.groupSizes[i])
      .sort((a, b) => groups[a].length - groups[b].length)[0];
    if (needy === undefined) return;

    const candidate = pool.find(p => !assigned.has(p.player) && covers(p));
    if (!candidate) return;
    groups[needy].push(candidate);
    assigned.add(candidate.player);
  }
}

/**
 * TBC totems only buff the shaman's own party, so each marker is bound to a
 * raid party and filled from it:
 *   - a marker with a Shaman in its healer row is already covered;
 *   - uncovered markers pick first, among parties that have a shaman in the
 *     pool, and get that shaman seeded;
 *   - the remaining markers take the largest parties still unclaimed;
 *   - each bound marker is filled from its party, and whoever does not fit
 *     spills to the nearest marker with room.
 */
function seedTotemParties_(run, encounter, groups, pool, assigned, healersAbove, warnings) {
  const partyOf = player => run.parties.get(player.player.toLowerCase()) || null;

  const covered = healersAbove.map(names => names
    .map(name => run.playersByName.get(name.toLowerCase()))
    .some(player => player && player.class.toLowerCase() === 'shaman'));

  // Party sizes are counted once, before any of these markers take people.
  const membersByParty = new Map();
  const shamanByParty = new Map();
  pool.forEach(player => {
    const party = partyOf(player);
    if (assigned.has(player.player) || !party) return;
    if (!membersByParty.has(party)) membersByParty.set(party, []);
    membersByParty.get(party).push(player);
    if (player.class.toLowerCase() === 'shaman' && !shamanByParty.has(party)) shamanByParty.set(party, player);
  });

  const claimed = new Set();
  const claimLargestParty = needsShaman => {
    let best = null;
    let bestCount = -1;
    membersByParty.forEach((members, party) => {
      if (claimed.has(party) || (needsShaman && !shamanByParty.has(party))) return;
      if (members.length > bestCount) {
        best = party;
        bestCount = members.length;
      }
    });
    if (best !== null) claimed.add(best);
    return best;
  };

  const plan = groups.map(() => null);
  covered.forEach((isCovered, i) => {
    if (isCovered) return;
    const party = claimLargestParty(true);
    if (party === null) return;
    plan[i] = party;
    const shaman = shamanByParty.get(party);
    if (!assigned.has(shaman.player)) {
      groups[i].push(shaman);
      assigned.add(shaman.player);
    }
  });
  plan.forEach((party, i) => {
    if (!party) plan[i] = claimLargestParty(false);
  });

  // BalanceGroups caps each marker at its share of the pool, so one big party
  // cannot fill a marker while the others stay nearly empty.
  const totalSeats = encounter.groupSizes.reduce((n, size) => n + size, 0);
  const capOf = i => encounter.balanceGroups
    ? Math.max(1, Math.min(encounter.groupSizes[i], Math.ceil(pool.length * (encounter.groupSizes[i] / totalSeats))))
    : encounter.groupSizes[i];

  const spill = [];
  plan.forEach((party, i) => {
    if (!party) return;
    pool.forEach(player => {
      if (assigned.has(player.player) || partyOf(player) !== party) return;
      if (groups[i].length >= capOf(i)) {
        spill.push({ player: player, from: i });
        return;
      }
      groups[i].push(player);
      assigned.add(player.player);
    });
  });

  // Nearest marker first: markers are listed in physical order, so a neighbour
  // is the most likely to still be in range of the party's totems.
  spill.forEach(entry => {
    const target = groups
      .map((group, i) => i)
      .filter(i => groups[i].length < capOf(i))
      .sort((a, b) => Math.abs(a - entry.from) - Math.abs(b - entry.from) || groups[a].length - groups[b].length)[0];
    if (target === undefined) return; // fillRemaining_ seats them
    groups[target].push(entry.player);
    assigned.add(entry.player.player);
  });

  if (!run.parties.size) {
    warnings.push({
      internal: true,
      text: 'no raid party data was found, so groups were not built around shamans',
      fix: `Check the "${AUTO_ASSIGN.PARTIES_SHEET}" sheet has columns headed ActiveRosterG1 to ActiveRosterG5.`,
    });
  }
}

/**
 * Seats everyone still unplaced, in pool order, into the emptiest group
 * (relative to its size) that suits their MarkerRoles preference. When there
 * are more players than seats, the least preferred ones sit out.
 */
function fillRemaining_(groups, encounter, pool, assigned) {
  const leftover = pool.filter(p => !assigned.has(p.player));
  if (!leftover.length) return;

  const sizes = encounter.groupSizes;
  const totalSeats = sizes.reduce((n, size) => n + size, 0);
  const total = groups.reduce((n, group) => n + group.length, 0) + leftover.length;
  const targetRatio = Math.min(1, total / totalSeats);
  const fillRatio = i => groups[i].length / sizes[i];
  const unseated = [];

  leftover.forEach(player => {
    const roomy = groups.map((group, i) => i).filter(i => groups[i].length < sizes[i]);
    if (!roomy.length) {
      unseated.push(player.player);
      return;
    }
    const belowTarget = roomy.filter(i => fillRatio(i) < targetRatio);
    const options = belowTarget.length ? belowTarget : roomy;
    const suited = options.filter(i => markerSuits_(player, encounter.markerRoles[i]));
    const target = (suited.length ? suited : options).sort((a, b) => fillRatio(a) - fillRatio(b))[0];

    groups[target].push(player);
    assigned.add(player.player);
  });

  if (unseated.length && !encounter.selection) {
    Logger.log(`${encounter.encounter} sat out: ${unseated.join(', ')} (${totalSeats} slots, ${total} eligible)`);
  }
}
