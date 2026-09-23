/**
 * Raid Assignments — one button fills every assignment on every boss sheet:
 * marker groups, tanks and healers, bloodlust, power infusion, innervates,
 * misdirects, interrupt rotations and anything else configured as a row.
 *
 * It calls assignHealersAndTanks() but never modifies it. That function lives in
 * AutoAssign.gs and no longer appears in the onOpen menu, so it looks unused —
 * it is not, and deleting it would break every run.
 *
 * Run autoAssignGroups() from anywhere — a drawing button, the menu, or the
 * script editor. It does not care which sheet is active.
 *
 * Config sheets are read by COLUMN HEADER NAME, not position, and the header
 * row is located automatically — so they can be Tables, sit below a title, or
 * have columns added and reordered freely.
 *
 * WHERE EACH FACT COMES FROM (one owner per fact):
 *   Raider Data      - the whole roster: who is raiding and what they play
 *                      (A = name, B = Class, C = Spec, D = flexSpec). It is
 *                      already filtered by the Is Rostered checkbox on the
 *                      Roster sheet, so everyone listed is in. Deliberately the
 *                      only player list this script reads — a second one could
 *                      only ever drift out of step with it.
 *                      EXCEPT a Roster Swap addition: a separate script ticks
 *                      Is Rostered for them too, purely so they show up in a
 *                      manual dropdown for one fight — not to make them a real
 *                      roster member for the night. loadRoster_ excludes anyone
 *                      whose name is currently sitting in a *_RosterSwaps named
 *                      range (see loadRosterSwapNames_), so they never enter an
 *                      automatic assignment pool. They remain pickable by hand.
 *   Class/Spec Data  - which Role a Class+Spec maps to (A/B = Class/Spec,
 *                      F = Role).
 *   Dynamic Lists    - which raid party each player is in, from columns headed
 *                      ActiveRosterG1 .. ActiveRosterG5.
 *
 * WHAT A RUN DOES, in order:
 *   1. Stops if nobody is rostered — with an empty roster every later step
 *      fails describing its own symptom rather than the cause.
 *   2. Fills TANKS_GLOBAL / HEALERS_GLOBAL if they are COMPLETELY empty, using
 *      TANK_PRIORITY / HEALER_PRIORITY, and says loudly that it guessed. One
 *      name already in them means a human decided, and that is never touched.
 *   3. Calls assignHealersAndTanks() to push those out to every boss sheet.
 *   4. Assigns every row in AutoAssignData, reading the healer rows step 3
 *      just wrote.
 * Step 3 rewrites the per-boss tank and healer tables from the globals on every
 * run, so a hand-edit to one of those tables does not survive — same as pressing
 * that button yourself, just harder to forget.
 *
 * NAMED RANGES:
 *   TANKS_GLOBAL / HEALERS_GLOBAL  already exist for assignHealersAndTanks.
 *     They decide this week's real Tank and Healer roles, overriding whatever a
 *     player's spec implies. Read on every run, written only when empty.
 *   TANK_PRIORITY / HEALER_PRIORITY  optional, you create. Ordered specs
 *     deciding who gets picked first when the above are empty. Without them it
 *     falls back to plain Tank and Healer roles.
 *   Anything a Preferred or WriteBackTo column names  optional, you create,
 *     one per assignment you want a human to be able to set.
 *   *_RosterSwaps (e.g. Split1_RosterSwaps)  optional, you create one per split.
 *     Names sitting in any of these are excluded from every automatic
 *     assignment pool (loadRosterSwapNames_) — they are one-fight substitutes,
 *     not real roster members, even though a separate script ticks Is Rostered
 *     for them so they show up in a manual dropdown. Add a new split's range
 *     and nothing else needs to change.
 *
 * Nothing else is addressed by name: every assignment target comes from
 * CellAnchor, which is why moving a table only means editing one cell.
 *
 * SETUP
 *
 * 1. Sheet "AutoAssignData" — one row per boss. Required headers (any order):
 *      Encounter   - the row's name. Must match the boss sheet's TAB NAME,
 *                    unless a Sheet column says otherwise. GroupRules match on
 *                    this, so give a boss's ranged and melee rows different
 *                    names (with Sheet naming the shared tab) when they need
 *                    different rules.
 *      CellAnchor  - first DPS cell of each group. Either one anchor per group,
 *                    comma separated ("BC10, BH10, BM10"), or a single anchor
 *                    plus ColumnStep below. A sheet prefix is fine and ignored.
 *      GroupAmount - how many groups, e.g. 3 for star/triangle/square.
 *                    0 (or a blank CellAnchor) means "skip this boss".
 *      GroupSize   - how many player rows each group holds. One number applies
 *                    to every group; a list gives each its own size, e.g.
 *                    "4, 5, 4, 4, 3" when the healer rows above them differ.
 *    Optional headers:
 *      Sheet         - the tab to write to, when it differs from Encounter.
 *                      Lets "Najentus Ranged" and "Najentus Melee" both target
 *                      the "Najentus" tab while carrying their own GroupRules.
 *                      May list several tabs — "Winterchill, Anetheron, …" —
 *                      so one row covers an assignment that sits in the same
 *                      place on every boss sheet (Bloodlust, PI, innervates).
 *                      "*" means every tab the other rows name, and "-Illidan"
 *                      removes one, so "*, -Illidan" is every boss but Illidan
 *                      and picks up new bosses automatically.
 *                      Those report as one line, "Bloodlust x12".
 *                      BLANK means the tab is whatever Encounter says — which
 *                      is why a row called "Bloodlust" with an empty Sheet goes
 *                      looking for a tab of that name and fails.
 *      ColumnStep    - how many columns apart the groups sit when CellAnchor is
 *                      a single cell. 1 (default) = side by side; 5 = BC, BH, BM.
 *                      Ignored when CellAnchor lists an anchor per group.
 *      BackupRows    - how many rows at the BOTTOM of each group are reserves
 *                      rather than active slots (e.g. 1 for Bloodboil's backup
 *                      row). Purely for reporting: if a later AutoRoles entry
 *                      ends up above the backup rows, the run says so, because
 *                      that means there were not enough first-choice players to
 *                      fill the slots that actually do the mechanic.
 *      AutoRoles     - (aka EligibleRoles) an ORDERED, comma-separated list of
 *                      who may be used, most preferred first. Entries can name
 *                      a Role, a Class, or a ClassSpec:
 *                      "Ranged"                  pure ranged groups
 *                      "Ranged, Healer"          ranged, then spare healers
 *                      "Ranged, Warrior, Melee"  ranged first, then warriors,
 *                                                then any other melee
 *                      (blank)                   everyone, no preference
 *                      Order matters when there are more candidates than seats:
 *                      earlier entries are placed first, later entries fill the
 *                      bottom rows and are the first to miss out entirely.
 *                      A bare Class ("Warrior") means a DPS of that class and
 *                      never picks up the tank or healer wearing it; to include
 *                      those, name the Role ("Tank") or the full ClassSpec
 *                      ("Protection Warrior"). "Any Druid" takes every spec of
 *                      a class regardless of role — for abilities the whole
 *                      class brings, like innervate.
 *                      "Flex Tank" asks what someone's OFF-spec would make them,
 *                      so a Feral druid who can go bear qualifies without being
 *                      a tank tonight. Works with a role ("Flex Tank"), a spec
 *                      ("Flex Guardian") or both ("Flex Guardian Druid"), and
 *                      reads the flexSpec column of Raider Data.
 *                      "Tank, Flex Tank" therefore means real tanks first, then
 *                      anyone who could step in.
 *      Raid          - which instance this row belongs to: "MH" or "BT". Only
 *                      read by the per-instance buttons (autoAssignMountHyjal,
 *                      autoAssignBlackTemple), which run about half the rows
 *                      each — worth having when a full run sits near the six
 *                      minute ceiling, and closer to how a night actually goes
 *                      when the two instances are on different evenings.
 *                      BLANK means both, and so do "*" and "all", matching
 *                      what "*" means in the Sheet column beside it. A "*"
 *                      Sheet row needs no Raid value:
 *                      scoping happens before "*" is expanded, so it resolves
 *                      to that instance's tabs on its own.
 *                      autoAssignGroups ignores this and still does everything.
 *      AllowDuplicates - TRUE to let the same name appear more than once in
 *                      Preferred. Most guilds double up Power Infusion and
 *                      innervates on their best target, so "Snitzz, Fakey,
 *                      Fakey" is a real assignment rather than a typo. Off by
 *                      default: on a marker group the same repeat would place
 *                      somebody in two positions at once.
 *      Critical      - TRUE when the raid genuinely cannot do this fight without
 *                      someone eligible — Council needs a Mage for Spellsteal.
 *                      Normally, AutoRoles matching nobody (or every group
 *                      short of its EnsureEach guarantee) just means the
 *                      roster leans a certain way this week, which is ordinary
 *                      variance and gets no mention, or a routine one at most.
 *                      Critical turns that same situation into a red warning,
 *                      because this specific gap is a wipe risk rather than a
 *                      shrug. Leave it off a row that already has its own
 *                      backstop for the gap — Archimonde's main marker row
 *                      does not need it, because assignArchimondeDecurse_
 *                      separately raises its own red warning only when there
 *                      is truly nobody left to decurse anywhere.
 *      IgnoreAbove   - Rarely needed; read this before reaching for it.
 *                      A row will not pick somebody whose name sits in the few
 *                      cells ABOVE its anchor, so a marker group does not
 *                      re-pick the healer sitting on top of it. That only
 *                      applies to names that arrived by FORMULA, which is how
 *                      a marker's healer row is filled ("=AV15"). A name that
 *                      was typed or chosen from a dropdown above this block
 *                      belongs to some other table and is ignored already —
 *                      so Fear Ward sitting under the Healer Assignments box
 *                      needs nothing set. A formula that READS FROM the block
 *                      is ignored too (a pre-pull cell mirroring the rotation
 *                      under it), since that describes the block rather than
 *                      claiming anyone. Use IgnoreAbove only for the odd case
 *                      neither rule catches; it skips the scan for this row.
 *      TotemParties  - TRUE to use shaman-party clustering (see below).
 *      EnsureEach    - every group must contain at least one of these, e.g.
 *                      "Mage, Druid" so no ranged group is left unable to
 *                      decurse. Uses the AutoRoles vocabulary, so a role, a
 *                      class or a ClassSpec all work, and several entries mean
 *                      "any of them" rather than one of each.
 *                      This is the common case of a GroupRules MinPerGroup rule
 *                      with the ceremony removed. Reach for GroupRules when a
 *                      boss needs a CAP (never two hunters in one group) or
 *                      several requirements at once — a row can only carry one
 *                      EnsureEach, and both are applied if you use both.
 *      ExcludeFrom   - Encounter name(s) whose people this row must not reuse.
 *                      The mage kiting Zerevor is rostered, on the sheet and a
 *                      fine interrupter — they are simply busy, and no role or
 *                      class filter can say that. Reads who that row actually
 *                      placed, so it follows the roster rather than a fixed
 *                      name. The row it names must sit ABOVE this one.
 *      LimitTo       - never fill more slots than another row managed to fill.
 *                      "PI Target" with LimitTo "PI Caster" leaves its second
 *                      slot empty when only one Discipline Priest is raiding,
 *                      instead of telling someone to expect a buff nobody can
 *                      cast. The row it names must sit ABOVE it in this sheet,
 *                      since the limit reads what that row actually assigned.
 *      Selection     - TRUE when this row picks a few people out of many, rather
 *                      than seating everyone who qualifies. One Bloodlust from
 *                      five shamans is a selection; three ranged marker groups
 *                      are not. It only silences the "sat out" notice, because
 *                      on a selection row the people not picked are the point,
 *                      not a shortfall — while on a group row they are somebody
 *                      standing in the raid with nowhere to go.
 *      WriteBackTo   - a named range on your setup page. When it is completely
 *                      empty, everyone eligible for this row is written into
 *                      it — in AutoRoles order, as many as the range holds —
 *                      so next run they arrive back through Preferred as a
 *                      visible choice you can change. A four-cell range gets
 *                      four candidates even if the boss table only seats one.
 *                      Never overwrites: one name in there means somebody has
 *                      decided. Pairs with Preferred — one reads the setup
 *                      page, the other seeds it.
 *      Preferred     - player names, in order, to place before anything is
 *                      worked out: "Snickels" for Bloodlust rather than letting
 *                      the script pick among five shamans. Being named is the
 *                      qualification, so these bypass AutoRoles. Anyone not
 *                      raiding that night is skipped and their slot fills
 *                      automatically — the manual setup degrades to the
 *                      automatic one instead of breaking.
 *      SkipWhen      - skip this row entirely when a cell says something, e.g.
 *                      "AE2 = Good DPS" for a phase that only happens on a slow
 *                      kill. "!=" works too. The cell is read on this row's own
 *                      sheet, and nothing is read or written when it matches —
 *                      whatever is already in those cells is left alone.
 *      MarkerRoles   - a preferred Role, Class or ClassSpec per marker, in
 *                      anchor order, e.g. "Ranged, Melee, Ranged, Melee, Ranged".
 *                      A soft preference, not a filter: people go to a marker
 *                      that suits them while there is room, and overflow into
 *                      the others rather than being left out. Use it to keep a
 *                      role in the markers that also exist in another phase, so
 *                      InheritFrom has more to carry over. Leave an entry blank
 *                      for a marker with no preference.
 *      MarkerNames   - what each group is called, in anchor order, e.g.
 *                      "Star, Circle, Square, Diamond, Triangle". Used to line
 *                      groups up with InheritFrom, and to name them in reports.
 *      InheritFrom   - the Encounter name of another row whose groups this one
 *                      should keep. Players already on a marker there stay on
 *                      the same marker here, so nobody has to learn two
 *                      positions for one fight; markers with no counterpart
 *                      fill normally. Matched by MarkerNames when both rows
 *                      have them, otherwise by position. The source row must
 *                      sit ABOVE this one in Encounters, since it reads the
 *                      values that row has just written.
 *      BalanceGroups - TRUE to keep the markers even in size while clustering.
 *                      Without it, a party holding most of the casters fills one
 *                      marker completely and leaves the others nearly empty.
 *                      Overflow goes to the NEAREST marker, so list CellAnchor
 *                      in physical order (star, triangle, square with triangle
 *                      in the middle) and a caster pushed out of one zone lands
 *                      next door, still inside their own shaman's totems.
 *                      Use it where even numbers matter for the mechanic; leave
 *                      it off where one fully-buffed group is worth more.
 *
 *    IMPORTANT: CellAnchor must point at the first assignable row, NOT a healer
 *    or fixed row above it. Those are formulas pulling from the healer table and
 *    must never be overwritten. Any row whose target cells contain a formula is
 *    skipped with a message rather than damaged.
 *
 * 2. Sheet "GroupRules" — may be empty apart from its headers:
 *      Encounter    - boss tab name, or "*" to apply to every encounter.
 *      MatchType    - "Class", "Spec", "Role", "ClassSpec", or "Player".
 *                     Use Player to pull one named raider out of a fight, e.g.
 *                     a Warlock who tanks Demon Illidan and so must not be
 *                     handed a normal DPS slot there.
 *      MatchValue   - one value, or several meaning "any of these":
 *                     "Mage, Druid" with MinPerGroup 1 puts a decurser in every
 *                     group, whichever class is spare. Two separate rules would
 *                     each chase their own minimum and double one of them up.
 *                     Otherwise e.g. "Shaman" (Class), "Melee" (Role),
 *                     "Restoration" (Spec — hits BOTH Restoration Shaman and
 *                     Restoration Druid), or "Restoration Shaman" (ClassSpec).
 *      MinPerGroup  - guarantee at least this many per group. Blank = none.
 *      MaxPerGroup  - cap per group. Blank = no cap. 0 = exclude entirely.
 *      Priority     - lower solves first. Blank = solves last.
 *
 * 3. TOTEM PARTIES (TotemParties = TRUE)
 *    In TBC totems only buff the shaman's own party, so markers are built
 *    around raid parties rather than by spreading classes:
 *      a. Read the healer already above each marker. If that healer is a
 *         Shaman, bind the marker to that shaman's raid party.
 *      b. For a marker whose healer is not a Shaman (e.g. a Disc Priest), find
 *         an unused Restoration or Elemental Shaman, drop them into that
 *         marker, and bind it to *their* party.
 *      c. Fill each marker with the eligible players from its bound party.
 *      d. Anyone left over (typically Hunters parked in the melee parties) goes
 *         to whichever marker currently has the fewest players.
 *    GroupRules still apply as caps during step (d).
 *
 * 4. Paste into Extensions > Apps Script. Then either:
 *      - add to your onOpen() menu:
 *          .addItem('Assign Everything', 'autoAssignGroups')
 *      - and/or Insert > Drawing a button anywhere, then Assign script:
 *          autoAssignGroups
 */

const AUTO_ASSIGN_ROLE_MAP_SHEET = 'Class/Spec Data';
// Tab holding the per-boss config — the data this script reads, as opposed to
// "Quick Assigns" where people actually edit things.
const AUTO_ASSIGN_ENCOUNTERS_SHEET = 'AutoAssignData';
const AUTO_ASSIGN_GROUP_RULES_SHEET = 'GroupRules';
const AUTO_ASSIGN_RAIDER_DATA_SHEET = 'Raider Data';
const AUTO_ASSIGN_DYNAMIC_LISTS_SHEET = 'Dynamic Lists';
// Shown as the dialog title and on every toast. It started life assigning DPS
// groups; it now does tanks, healers, markers and every utility besides.
const AUTO_ASSIGN_TOAST_TITLE = 'Auto Assignments Overview';

// Named ranges owned by assignHealersAndTanks. Normally read-only — the one
// exception is filling them when they are completely empty, see autoFillPicks_.
const AUTO_ASSIGN_TANKS_RANGE = 'TANKS_GLOBAL';
const AUTO_ASSIGN_HEALERS_RANGE = 'HEALERS_GLOBAL';

// Optional. Ordered lists deciding who gets picked first when the above are
// empty — one cell each, e.g. "Protection Warrior, Protection Paladin".
// Which tank you want as MT is a preference, not a fact, so it lives in the
// sheet rather than in here where nobody copying this could change it.
const AUTO_ASSIGN_TANK_PRIORITY_RANGE = 'TANK_PRIORITY';
const AUTO_ASSIGN_HEALER_PRIORITY_RANGE = 'HEALER_PRIORITY';

// Archimonde's Decurse Assignments table — a one-off, hand-built exception to
// the usual AutoAssignData row model, because nothing else needs "read a real
// placement, then cycle leftovers evenly across cells that repeat names."
// See assignArchimondeDecurse_. Update these if the layout ever moves.
const AUTO_ASSIGN_ARCHIMONDE_SHEET = 'Archimonde';
// Star, Triangle, Square member cells — must match that row's CellAnchor
// (BC10) / ColumnStep (5) / GroupSize (5) in AutoAssignData.
const AUTO_ASSIGN_ARCHIMONDE_SOURCE_RANGES = ['BC10:BC14', 'BH10:BH14', 'BM10:BM14'];
// Star, Triangle, Square, Tank Group, Melee Group, in that order — the order
// IS the priority order when coverage has to be split unevenly.
const AUTO_ASSIGN_ARCHIMONDE_TARGET_CELLS = ['BH17', 'BH18', 'BH19', 'BH20', 'BH21'];

// Only used when no Archimonde row has an EnsureEach to read. Written in the
// AutoRoles vocabulary, so "Any" ignores role: in TBC every druid and mage
// spec can Remove Curse, tanks and healers included.
const AUTO_ASSIGN_ARCHIMONDE_FALLBACK_DECURSERS = ['any mage', 'any druid'];

// Used only when those ranges are absent. Deliberately plain: every tank spec,
// every healer spec, so it fills something sensible without pretending to know
// which of them a given guild wants leading.
const AUTO_ASSIGN_DEFAULT_TANK_PRIORITY = ['Tank'];
const AUTO_ASSIGN_DEFAULT_HEALER_PRIORITY = ['Healer'];

// The Roster tab, by the gid in its URL. Used to jump people there from the
// "no roster" dialog — an id rather than a name, so renaming the tab is safe.
const AUTO_ASSIGN_ROSTER_SHEET_ID = 407225357;

const AUTO_ASSIGN_PARTY_COUNT = 5;
// Ceiling on a per-sheet snapshot. Blocks scattered to opposite corners of a
// sheet make the bounding box enormous, at which point reading each range
// separately is cheaper than one huge read.
const AUTO_ASSIGN_MAX_SNAPSHOT_CELLS = 40000;
// How far down a config sheet to look for its header row (tables, title rows).
const AUTO_ASSIGN_HEADER_SCAN_ROWS = 10;
// How far ABOVE a group's anchor to look for the names already fixed to it —
// healers, and any other named role like a Warlock tanking Demon Illidan.
//
// Deliberately small: a boss sheet can stack several assignment tables in the
// same columns (Illidan's Phase 2&3 groups sit above its Demon Phase groups,
// Najentus's melee groups below its ranged ones), and scanning further would
// treat a neighbouring table as part of this group. Four covers a marker with
// three fixed names above it while still clearing the gap to the next table.
const AUTO_ASSIGN_HEALER_SCAN_ROWS = 4;

/**
 * Where the run spends its time. Apps Script stops a script at 6 minutes, so
 * knowing which phase dominates matters before optimising anything.
 */
const AUTO_ASSIGN_TIMING = {
  reads: 0, clearing: 0, flushing: 0, validating: 0, writing: 0,
  guarding: 0, scanning: 0,
};

function timed_(phase, work) {
  const started = new Date();
  const result = work();
  AUTO_ASSIGN_TIMING[phase] += new Date() - started;
  return result;
}

/** Entry point. Assigns every boss configured in AutoAssignData. */
function autoAssignGroups() {
  autoAssignScoped_(null);
}

/**
 * Per-instance buttons. Same run, limited to rows whose Raid column matches —
 * about half the work each, which matters when a full run sits close to the
 * six-minute ceiling. Rows with a blank Raid are included in both, so anything
 * spanning the whole night needs no marking.
 *
 * Assign these to their own drawings or menu items alongside autoAssignGroups.
 */
function autoAssignMountHyjal() {
  autoAssignScoped_('mh');
}

function autoAssignBlackTemple() {
  autoAssignScoped_('bt');
}

function autoAssignScoped_(scope) {
  const ss = SpreadsheetApp.getActive();
  try {
    runAutoAssignGroups_(ss, scope);
  } catch (err) {
    const message = String((err && err.message) || err);

    // Google's own errors, which say nothing about this sheet and mean "try
    // again" rather than "fix something".
    const transient = /service unavailable|internal error|try again|timed out|too many/i
      .test(message);

    showResult_(ss, {
      ok: false,
      // Deliberately not "stopped before writing": this catch covers the whole
      // run, so a failure here may well have left some sheets already assigned.
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

/**
 * Called by the "Open Instructions" button on the no-roster dialog.
 *
 * Switches to the Roster tab and then opens the tutorial, so people land on the
 * page the instructions talk about rather than having to find it.
 */
function openRosterInstructions() {
  const sheet = SpreadsheetApp.getActive().getSheets()
    .filter(s => s.getSheetId() === AUTO_ASSIGN_ROSTER_SHEET_ID)[0];

  if (sheet) sheet.activate();
  else Logger.log(`No sheet with id ${AUTO_ASSIGN_ROSTER_SHEET_ID} — check the gid in its URL.`);

  // Replaces this dialog when it opens one of its own.
  if (typeof showRosterTutorialPopup === 'function') showRosterTutorialPopup();
}

/**
 * A step of the run, for the Execution log only.
 *
 * These used to toast as well, which meant a popup per boss row — 125 of them,
 * including "Assigning Caster PI" a dozen times over as the wildcard row
 * walked the sheets. Most appeared for a second or two, and none of it was
 * anything a person waiting on the run would act on. The log still carries the
 * full sequence for diagnosing a bad run afterwards; see announce_ for the
 * handful of things actually worth interrupting somebody with.
 */
function progress_(ss, message) {
  Logger.log(message);
}

/**
 * Adds a problem to the dialog, keeping setup detail away from people who
 * cannot act on it.
 *
 * Whoever reads this dialog on raid night can fix a roster or widen a
 * dropdown. They cannot fix a CellAnchor, and sending them into AutoAssignData
 * to try is how the config gets broken. So anything marked `internal` shows as
 * a single short line and puts its real text and fix in the Execution log,
 * where the person who maintains the sheet will actually look.
 *
 * Everything else — someone sat out, a dropdown refusing a name, a fight
 * missing the class it needs — is left exactly as written, because that is
 * the raid's problem to solve rather than the spreadsheet's.
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

/**
 * The rare toast worth showing a person mid-run.
 *
 * Reserved for "something is happening, leave the sheet alone" — the only
 * thing anybody watching a several-minute run needs to know. A modal cannot
 * do this job: Apps Script is busy inside this very call, so a dialog opened
 * here could not be updated until the work it is describing had finished.
 */
function announce_(ss, message, seconds) {
  Logger.log(message);
  ss.toast(message, AUTO_ASSIGN_TOAST_TITLE, seconds || 30);
}

/**
 * Shows the run result as a dialog. Falls back to a toast when there is no UI
 * to attach to (editor runs, triggers), and always logs, so no outcome is ever
 * silently lost.
 *
 * result = { ok, headline, assigned[], skipped[], notes[], issues[] }
 */
function showResult_(ss, result) {
  Logger.log(result.headline);
  (result.assigned || []).forEach(line => Logger.log(`  assigned: ${line}`));
  (result.skipped || []).forEach(line => Logger.log(`  skipped: ${line}`));
  (result.notes || []).forEach(line => Logger.log(`  note: ${line}`));
  [['issue', result.issues], ['fyi', result.notices]].forEach(([label, items]) => {
    (items || []).forEach(item =>
      Logger.log(`  ${label}: ${item.where ? item.where + ' — ' : ''}${item.text}` +
        (item.places ? ` [${item.places}]` : '') +
        (item.fix ? ` (${item.fix})` : '')));
  });

  try {
    const html = HtmlService.createHtmlOutput(renderResultHtml_(result))
      .setWidth(560)
      .setHeight(460);
    SpreadsheetApp.getUi().showModalDialog(html, AUTO_ASSIGN_TOAST_TITLE);
  } catch (err) {
    const flat = [result.headline]
      .concat((result.issues || []).map(i => `${i.where || ''} ${i.text}`.trim()))
      .concat(result.notes || [])
      .join(' | ');
    ss.toast(flat, AUTO_ASSIGN_TOAST_TITLE, 30);
  }
}

/**
 * Merges entries that say the same thing into one, listing where they happened.
 *
 * A wildcard row runs on every boss sheet, so one mistake becomes a dozen
 * identical lines — which buries the other problems and makes the count
 * meaningless. Same message and same fix is one finding in several places.
 */
function collapseIssues_(items) {
  const groups = [];
  const seen = {};

  (items || []).forEach(item => {
    const key = `${item.text}||${item.fix || ''}`;
    if (seen[key] === undefined) {
      seen[key] = groups.length;
      groups.push({ text: item.text, fix: item.fix, places: [] });
    }
    if (item.where) groups[seen[key]].places.push(item.where);
  });

  return groups.map(group => {
    if (group.places.length <= 1) {
      return { where: group.places[0], text: group.text, fix: group.fix };
    }
    const shown = group.places.slice(0, 6).join(', ');
    const extra = group.places.length > 6 ? ` +${group.places.length - 6} more` : '';
    return {
      where: `${group.places.length}×`,
      text: group.text,
      places: shown + extra,
      fix: group.fix,
    };
  });
}

function escapeHtml_(text) {
  return String(text)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

function renderResultHtml_(result) {
  // Issues are the reason this dialog exists, so they get the room. Everything
  // that went fine is compressed to a line — a per-boss list of successes just
  // pushes the actionable items off screen.
  const entryList = items => (items || []).map(item => `
      <li>
        <div class="what">${item.where ? `<b>${escapeHtml_(item.where)}:</b> ` : ''}${escapeHtml_(item.text)}</div>
        ${item.places ? `<div class="places">${escapeHtml_(item.places)}</div>` : ''}
        ${item.fix ? `<div class="fix">${escapeHtml_(item.fix)}</div>` : ''}
      </li>`).join('');

  const issues = entryList(result.issues);
  const notices = entryList(result.notices);
  const notes = (result.notes || []).map(note => `<li>${escapeHtml_(note)}</li>`).join('');

  // An optional button that calls another function in this project. The name is
  // stripped to word characters before going into the page — it ends up inside
  // a script tag, so anything else has no business being there.
  const safeFn = result.action
    ? String(result.action.fn).replace(/[^A-Za-z0-9_]/g, '') : '';
  const action = safeFn ? `
      <button class="action" onclick="
        this.disabled = true; this.textContent = 'Opening…';
        google.script.run.${safeFn}();
      ">${escapeHtml_(result.action.label)}</button>` : '';

  // Nothing here lists what went right. On a good run that was fifty boss
  // names nobody reads, pushing the handful of things that need a person off
  // the screen.
  //
  // Skipped rows are deliberately absent too: a SkipWhen firing means the
  // sheet was set up to skip that phase and did exactly that. Announcing it
  // every week only teaches people to skim past this panel. result.skipped is
  // still written to the Execution log by showResult_.

  return `
    <style>
      :root { color-scheme: light; }
      body {
        margin: 0;
        font: 13px/1.55 -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
        color: #202124;
        background: #fff;
      }
      .wrap { padding: 14px 18px 24px; }
      .headline {
        font-size: 15px;
        font-weight: 600;
        padding: 11px 14px;
        border-radius: 6px;
        border-left: 4px solid;
        margin-bottom: 14px;
      }
      .ok   { background: #e7f5ec; border-color: #1e8e3e; color: #12652b; }
      .fail { background: #fce8e6; border-color: #d93025; color: #a50e0e; }
      h2 {
        font-size: 11px;
        text-transform: uppercase;
        letter-spacing: .07em;
        color: #5f6368;
        margin: 0 0 8px;
        font-weight: 600;
      }
      .count {
        display: inline-block;
        background: #f1f3f4;
        border-radius: 9px;
        padding: 0 7px;
        margin-left: 3px;
        font-size: 11px;
        letter-spacing: 0;
      }
      section { margin-bottom: 18px; }
      ul { margin: 0; padding: 0; list-style: none; }
      .issues li {
        border-left: 3px solid #f28b82;
        padding: 2px 0 2px 10px;
        margin-bottom: 10px;
      }
      .issues .what { color: #a50e0e; }
      .issues .fix, .notices .fix { color: #5f6368; font-size: 12px; margin-top: 2px; }
      .places { color: #3c4043; font-size: 12px; margin-top: 2px; }
      .guessed {
        background: #fef7e0;
        border: 1px solid #f9ab00;
        border-radius: 6px;
        padding: 12px 14px;
        margin-bottom: 16px;
        color: #7a4f01;
      }
      .guessed .picked {
        font-family: ui-monospace, Menlo, Consolas, monospace;
        font-size: 12px;
        margin: 6px 0 8px;
        color: #3c4043;
      }
      .notices li {
        border-left: 3px solid #c6dafc;
        padding: 2px 0 2px 10px;
        margin-bottom: 10px;
      }
      .notices .what { color: #3c4043; }
      .notes li {
        color: #8a5a00;
        border-left: 3px solid #fdd663;
        padding: 2px 0 2px 10px;
        margin-bottom: 8px;
      }
      .quiet {
        color: #5f6368;
        font-size: 12px;
        padding-top: 8px;
        border-top: 1px solid #e8eaed;
        margin-top: 4px;
      }
      .quiet b { color: #3c4043; }
      .empty { color: #5f6368; font-style: italic; }
      .action {
        font: inherit;
        font-weight: 600;
        color: #fff;
        background: #1a73e8;
        border: 0;
        border-radius: 4px;
        padding: 9px 18px;
        cursor: pointer;
        margin-bottom: 14px;
      }
      .action:hover { background: #1765cc; }
      .action:disabled { background: #9aa0a6; cursor: default; }
    </style>
    <div class="wrap">
      <div class="headline ${result.ok ? 'ok' : 'fail'}">${escapeHtml_(result.headline)}</div>
      ${(result.autoPicked || []).length ? `<div class="guessed">
        <b>Nobody had picked tanks or healers, so the script chose by class priority:</b>
        <div class="picked">${escapeHtml_(result.autoPicked.join(' · '))}</div>
        Set them yourself in Quick Assigns and run this again. Everything below was
        built on these picks, so a wrong main tank means wrong groups too.
      </div>` : ''}
      ${issues ? `<section class="issues">
        <h2>Needs attention <span class="count">${result.issues.length}</span></h2>
        <ul>${issues}</ul></section>` : ''}
      ${notes ? `<section class="notes">
        <h2>Setup notes <span class="count">${result.notes.length}</span></h2>
        <ul>${notes}</ul></section>` : ''}
      ${notices ? `<section class="notices">
        <h2>For your information <span class="count">${result.notices.length}</span></h2>
        <ul>${notices}</ul></section>` : ''}
      ${action}
      ${issues || notices ? '' : '<p class="empty">Nothing needed your attention.</p>'}
    </div>`;
}

function runAutoAssignGroups_(ss, scope) {
  const started = new Date();

  announce_(ss, 'Starting — reading your roster…', 30);
  progress_(ss, 'Checking tank and healer picks…');

  const setupWarnings = [];
  let tankPicks = loadNamesFromNamedRange_(ss, AUTO_ASSIGN_TANKS_RANGE);
  let healerPicks = loadNamesFromNamedRange_(ss, AUTO_ASSIGN_HEALERS_RANGE);

  // With an empty roster everything downstream fails, but each failure describes
  // its own symptom rather than the cause — the tank check would send you to
  // Quick Assigns when the real problem is that nobody has been ticked in yet.
  const rosterSheet = ss.getSheetByName(AUTO_ASSIGN_RAIDER_DATA_SHEET);
  if (!rosterSheet) throw new Error(`Sheet "${AUTO_ASSIGN_RAIDER_DATA_SHEET}" not found`);

  // Counting rows is not enough: Raider Data is a table of formulas, so every
  // row still exists when nobody is ticked in — they just return blank.
  const rosterLastRow = rosterSheet.getLastRow();
  const rosteredNames = rosterLastRow < 2 ? [] :
    rosterSheet.getRange(2, 1, rosterLastRow - 1, 1)
      .getDisplayValues()
      .flat()
      .filter(value => String(value).trim() !== '');

  if (!rosteredNames.length) {
    showResult_(ss, {
      ok: false,
      headline: 'No Roster Setup!',
      issues: [{ text: 'Run "Instructions" in Roster.' }],
      // Always offered: even without the tutorial function, jumping to the
      // right tab is worth a click.
      action: { label: 'Open Instructions', fn: 'openRosterInstructions' },
    });
    return;
  }

  const roleMap = loadRoleMap_(ss);

  // Loaded early so it can also be excluded from the tank/healer emergency
  // auto-fill below, not just the main roster: a Roster Swap addition gets
  // "Is Rostered" ticked by a different script (to populate a manual
  // dropdown), so without this it could quietly become eligible for automatic
  // assignment on every boss all night, or even get auto-picked as tonight's
  // Main Tank. See loadRoster_ for where the exclusion actually happens.
  const swapNames = loadRosterSwapNames_(ss);

  // Who is actually raiding tonight, before any of this week's picks colour
  // anyone's role. Needed here rather than later because last week's picks
  // have to be checked against it before they are broadcast anywhere.
  const rosterNow = loadRoster_(ss, roleMap, new Set(), new Set(), swapNames);
  const rosteredKeys = new Set(rosterNow.map(p => p.player.toLowerCase()));

  Logger.log(`Roster as the script reads it: ${rosterNow.map(p =>
    `${p.player}=${p.classspec || '?'}`).join(', ')}`);

  // Last week's picks are still sitting in the sheet. Anyone not raiding is
  // taken out and their slot refilled, before the broadcast can put a name on
  // a boss sheet for somebody who is not in the building. See reconcilePicks_.
  const staleNotes = [];

  [
    { range: AUTO_ASSIGN_TANKS_RANGE, label: 'tank', protectFirst: true,
      priority: loadPriority_(ss, AUTO_ASSIGN_TANK_PRIORITY_RANGE, AUTO_ASSIGN_DEFAULT_TANK_PRIORITY) },
    { range: AUTO_ASSIGN_HEALERS_RANGE, label: 'healer', protectFirst: false,
      priority: loadPriority_(ss, AUTO_ASSIGN_HEALER_PRIORITY_RANGE, AUTO_ASSIGN_DEFAULT_HEALER_PRIORITY) },
  ].forEach(job => {
    const outcome = reconcilePicks_(ss, job.range, rosteredKeys, rosterNow,
      job.priority, job.protectFirst);

    if (outcome.removed.length) {
      staleNotes.push({
        info: true,
        text: `${outcome.removed.join(', ')} ${outcome.removed.length === 1 ? 'is' : 'are'} ` +
          `not raiding, so they were taken out of your ${job.label} picks` +
          (outcome.added.length
            ? `. ${outcome.added.join(', ')} filled the ` +
              `${outcome.added.length === 1 ? 'gap' : 'gaps'}`
            : ''),
      });
    } else if (outcome.added.length) {
      // Blank slots filled without anybody having been dropped, which is what
      // happens when somebody clears a cell to make the script choose again.
      staleNotes.push({
        info: true,
        text: `Empty ${job.label} slot${outcome.added.length === 1 ? '' : 's'} filled with ` +
          `${outcome.added.join(', ')}.`,
      });
    }

    // Named on its own line rather than buried in the list above. Everything
    // else about a swapped tank is routine; who leads the pull is not, and
    // this is the prompt that gets somebody to check it.
    if (outcome.firstFilled) {
      staleNotes.push({
        info: true,
        text: `${outcome.firstFilled} was picked as your main tank. ` +
          `Change it in Quick Assigns if that is not who you want.`,
      });
    }

  });

  if (staleNotes.length) {
    SpreadsheetApp.flush();
    tankPicks = loadNamesFromNamedRange_(ss, AUTO_ASSIGN_TANKS_RANGE);
    healerPicks = loadNamesFromNamedRange_(ss, AUTO_ASSIGN_HEALERS_RANGE);
  }

  // Nobody has picked, so pick for them rather than refusing to run. Only ever
  // fires on a completely empty range: one name means a human has decided, and
  // that is never overwritten.
  const autoPicked = [];
  const needsTanks = tankPicks !== null && !tankPicks.size;
  const needsHealers = healerPicks !== null && !healerPicks.size;

  Logger.log(
    `${AUTO_ASSIGN_TANKS_RANGE}: ${tankPicks === null ? 'missing' : tankPicks.size + ' name(s)'}, ` +
    `${AUTO_ASSIGN_HEALERS_RANGE}: ${healerPicks === null ? 'missing' : healerPicks.size + ' name(s)'} ` +
    `— auto-fill ${needsTanks || needsHealers ? 'will run' : 'not needed'}`);

  if (needsTanks || needsHealers) {
    progress_(ss, 'No tanks or healers picked. Choosing some from the roster...');
    const provisional = rosterNow;

    if (needsTanks) {
      const names = autoFillPicks_(ss, AUTO_ASSIGN_TANKS_RANGE, provisional,
        loadPriority_(ss, AUTO_ASSIGN_TANK_PRIORITY_RANGE, AUTO_ASSIGN_DEFAULT_TANK_PRIORITY));
      if (names.length) autoPicked.push(`tanks (${names.join(', ')})`);
    }
    if (needsHealers) {
      const names = autoFillPicks_(ss, AUTO_ASSIGN_HEALERS_RANGE, provisional,
        loadPriority_(ss, AUTO_ASSIGN_HEALER_PRIORITY_RANGE, AUTO_ASSIGN_DEFAULT_HEALER_PRIORITY));
      if (names.length) autoPicked.push(`healers (${names.join(', ')})`);
    }

    if (autoPicked.length) {
      SpreadsheetApp.flush();
      tankPicks = loadNamesFromNamedRange_(ss, AUTO_ASSIGN_TANKS_RANGE);
      healerPicks = loadNamesFromNamedRange_(ss, AUTO_ASSIGN_HEALERS_RANGE);
    }
  }

  // Global preflight — an empty tank/healer table invalidates every boss.
  const missingPicks = describeMissingPicks_(tankPicks, healerPicks);
  if (missingPicks.length) {
    showResult_(ss, {
      ok: false,
      headline: 'Stopped before writing: tanks and healers are not set',
      issues: missingPicks.map(text => ({
        text: text,
        fix: 'Pick your tanks and healers on the Roster page, then run Assign Healers/Tanks.',
      })),
      // Without this, a range emptied a moment ago because everybody in it is
      // out sick reads as "you never set these up", which sends somebody
      // hunting for a problem that does not exist.
      notes: staleNotes.map(note => note.text).concat([
        'The healer row above each group is filled by formula from those picks, so without them every group would be built on blank input.',
      ]),
    });
    return;
  }

  progress_(ss, `Found ${tankPicks.size} tank(s) and ${healerPicks.size} healer(s). Reading the Encounters list…`);

  // Push the picks out to the per-boss tables before assigning anything. The
  // group blocks read the healer rows sitting above them, and those are driven
  // by this — so running the groups without it builds on yesterday's healers.
  //
  // Guarded by typeof so this file still works on its own if someone copies it
  // without the original AutoAssign.gs.
  if (typeof assignHealersAndTanks === 'function') {
    progress_(ss, 'Sending tanks and healers out to the boss sheets…');

    logBroadcastCoverage_(ss, tankPicks.size, healerPicks.size);

    try {
      assignHealersAndTanks();
      SpreadsheetApp.flush(); // healer rows must be current before we snapshot them
    } catch (err) {
      setupWarnings.push(`Could not run assignHealersAndTanks: ${err.message}`);
    }
  } else {
    setupWarnings.push(
      'assignHealersAndTanks was not found, so the per-boss tank and healer tables ' +
      'were not refreshed. Run that separately if they look out of date.');
  }

  const encounters = loadEncounterConfigs_(ss, setupWarnings, scope);
  if (!encounters.length) {
    showResult_(ss, {
      ok: false,
      headline: 'Nothing to do',
      issues: [{
        text: `no boss rows were found in the "${AUTO_ASSIGN_ENCOUNTERS_SHEET}" sheet`,
        fix: 'Add a row per boss with Encounter, CellAnchor, GroupAmount and GroupSize filled in.',
      }],
    });
    return;
  }

  progress_(ss, `${encounters.length} boss row(s) found. Gathering rostered players…`);

  const partyByPlayer = loadPartyMembership_(ss);
  const excludedSwaps = [];
  const allPlayers = loadRoster_(ss, roleMap, tankPicks, healerPicks, swapNames, excludedSwaps);

  // Roles, not just specs. This is the list every AutoRoles match is made
  // against, and the role here already reflects tonight's tank and healer
  // picks overriding whatever the spec implies. Without it, "why was this
  // person not eligible" needs guessing: a Feral Druid who is tanking reads
  // as Tank and is invisible to an AutoRoles of "Melee", which looks like a
  // bug right up until you can see the role.
  Logger.log(`Roles used for matching: ${allPlayers.map(p =>
    `${p.player}=${p.classspec || '?'}/${p.role || 'no role'}` +
    (p.flexclassspec ? ` (flex ${p.flexclassspec})` : '')).join(', ')}`);
  const playersByName = new Map(allPlayers.map(p => [p.player.toLowerCase(), p]));
  const allRules = loadGroupRules_(ss);
  const prepared = prepareEncounters_(ss, encounters);
  const ownedCells = prepared.owned;
  const snapshots = prepared.snapshots;
  // "label||sheet" -> how many that row placed, so LimitTo rows can cap themselves.
  const placementCounts = {};
  // WriteBackTo targets already handled this run.
  const writtenBack = {};
  // "label||sheet" -> who that row placed, so ExcludeFrom rows can skip them.
  const placedNames = {};
  // "label||sheet" -> true when a SkipWhen sent that row home. Lets an
  // InheritFrom row tell "the phase I copy from is not happening tonight",
  // which is expected and silent, from "it ran and produced nothing", which
  // usually means something is actually wrong.
  const skippedRows = {};
  // Named ranges seeded because Quick Assigns was empty — reported as one
  // combined notice at the end rather than one line per range, since to a
  // reader "Bloodlust was empty, so X went in" repeated ten times says nothing
  // that one sentence doesn't already say.
  const writeBackFilled = [];

  progress_(ss,
    `${allPlayers.length} rostered, ${partyByPlayer.size} placed in raid parties, ` +
    `${allRules.length} group rule(s). Starting assignments…`);

  // The one that matters: from here the run spends minutes writing, and the
  // only useful thing to tell somebody is not to type into the sheet while it
  // does. Long timeout so it is still on screen when they look back.
  announce_(ss,
    `Assigning ${encounters.length} rows — this takes a few minutes. ` +
    `Please don't edit the sheet until it finishes.`, 300);

  // Ten updates over the whole run, no matter how many rows there are. Enough
  // to see it moving; far short of the 125 popups this replaced. Rows that are
  // switched off never reach the loop body, so they are left out of the total
  // rather than making the bar stall on nothing.
  const barTotal = encounters.filter(e => e.groupCount && e.anchor).length;
  const barEvery = Math.max(1, Math.ceil(barTotal / 10));
  let barDone = 0;

  const assignedBosses = [];
  const skipped = [];
  const problems = [];  // things the user must act on
  const notices = [];   // expected outcomes worth seeing, but not faults

  // Gathered before the boss loop, back when the picks were reconciled against
  // tonight's roster. Added here so they sit with everything else rather than
  // needing their own section.
  staleNotes.forEach(item => notices.push(item));

  // Named so a leftover-write error caught elsewhere can point at the row that
  // actually wrote it, rather than whichever row happened to read next.
  let previousLabel = null;

  encounters.forEach((encounter, index) => {
    // A row with no split configured is off on purpose and needs no announcing.
    // Only a SkipWhen condition firing is worth a mention, since that one is a
    // decision the sheet made this run rather than a permanent setting.
    if (!encounter.groupCount || !encounter.anchor) return;

    // A wildcard row runs on every boss sheet, so the label alone does not say
    // where something went wrong — name the tab whenever the row is repeated.
    const where = encounter.repeated
      ? `${encounter.encounter} on ${encounter.sheetName}`
      : encounter.encounter;

    progress_(ss, `Assigning ${encounter.encounter} (${index + 1} of ${encounters.length})…`);

    barDone++;
    if (barDone % barEvery === 0 || barDone === barTotal) {
      const filled = Math.min(10, Math.round(barDone / barTotal * 10));
      // No second line here: a toast collapses newlines to a space, and the
      // "don't edit" plea was already made once when the run started.
      announce_(ss,
        `${'▓'.repeat(filled)}${'░'.repeat(10 - filled)}  ` +
        `${Math.round(barDone / barTotal * 100)}%`, 300);
    }

    try {
      const outcome = assignEncounter_({
        ss: ss,
        encounter: encounter,
        allEncounters: encounters,
        knownTerms: roleMap.terms,
        ownedCells: ownedCells[encounter.sheetName] || {},
        snapshot: snapshots[encounter.sheetName] || null,
        placementCounts: placementCounts,
        placedNames: placedNames,
        skippedRows: skippedRows,
        writtenBack: writtenBack,
        writeBackFilled: writeBackFilled,
        allPlayers: allPlayers,
        playersByName: playersByName,
        partyByPlayer: partyByPlayer,
        allRules: allRules,
        previousLabel: previousLabel,
      });
      previousLabel = where;

      if (outcome.skipped) {
        skipped.push(`${where} (${outcome.skipped})`);
        skippedRows[`${encounter.encounter.toLowerCase()}||${encounter.sheetName.toLowerCase()}`] = true;
      } else if (outcome.problem) {
        pushProblem_(problems, where, outcome.problem);
      } else {
        const key = `${encounter.encounter.toLowerCase()}||${encounter.sheetName.toLowerCase()}`;
        placementCounts[key] = outcome.placed;
        placedNames[key] = outcome.names;

        assignedBosses.push(encounter.repeated
          ? { label: encounter.encounter, repeated: true }
          : { label: `${encounter.encounter} (${outcome.placed})` });
        outcome.warnings.forEach(w => {
          if (w.info) {
            notices.push({ where: encounter.encounter, text: w.text, fix: w.fix });
            return;
          }
          pushProblem_(problems, encounter.encounter, w);
        });
      }
    } catch (err) {
      // An exception escaping a row is always a setup fault rather than
      // anything the raid can act on, so it is never shown raw.
      pushProblem_(problems, where, {
        internal: true,
        text: err.message,
        fix: /range not found/i.test(err.message)
          ? `That cell does not exist on this sheet, which may be narrower or shorter than the others. Exclude it with "-${encounter.sheetName}" in Sheet if it has no such table.`
          : undefined,
      });
    }
  });

  // A one-off, hand-built exception to the AutoAssignData row model — see
  // assignArchimondeDecurse_. Runs after the loop above so Star/Triangle/
  // Square are already written and there is something real to read.
  // Only when this run actually covered that tab. Asking "did any row target
  // Archimonde?" rather than "is this the Mount Hyjal button?" keeps it right
  // whatever the Raid column says, and stops a Black Temple run reporting on a
  // fight it never touched.
  const touchedArchimonde = encounters.some(e =>
    e.groupCount && e.anchor && e.sheetName === AUTO_ASSIGN_ARCHIMONDE_SHEET);

  try {
    if (touchedArchimonde) assignArchimondeDecurse_(ss, allPlayers, problems, notices, encounters);
  } catch (err) {
    pushProblem_(problems, 'Archimonde Decurse Assignments', { internal: true, text: err.message });
  }

  // A row spanning twelve sheets would otherwise fill the dialog with twelve
  // near-identical entries, so those collapse to "Bloodlust x12".
  const tally = [];
  const repeatCounts = {};
  assignedBosses.forEach(entry => {
    if (!entry.repeated) {
      tally.push(entry.label);
      return;
    }
    if (repeatCounts[entry.label] === undefined) {
      repeatCounts[entry.label] = tally.push(entry.label) - 1;
    }
  });
  Object.keys(repeatCounts).forEach(label => {
    const count = assignedBosses.filter(e => e.repeated && e.label === label).length;
    tally[repeatCounts[label]] = `${label} ×${count}`;
  });

  if (writeBackFilled.length) {
    Logger.log(`Quick Assigns seeded this run: ${writeBackFilled.join(', ')}`);
    notices.push({
      info: true,
      text: 'We auto filled it for you, adjust if needed in Quick Assigns.',
    });
  }

  // Last on purpose. Nobody reads a panel expecting the least urgent thing
  // first, and this one is only ever "here is who the script left alone".
  if (excludedSwaps.length) {
    notices.push({
      info: true,
      text: `Sitting out of automatic assignment as Roster Swap pick(s): ${excludedSwaps.join(', ')}. ` +
        `Place them by hand if you need them for a specific fight.`,
    });
  }

  const seconds = ((new Date() - started) / 1000).toFixed(1);
  const spent = Object.keys(AUTO_ASSIGN_TIMING)
    .map(phase => `${phase} ${(AUTO_ASSIGN_TIMING[phase] / 1000).toFixed(1)}s`)
    .join(', ');
  Logger.log(`Time spent — ${spent} (total ${seconds}s of the 6 minute limit)`);
  // Notices deliberately do not make the run "not clean" — only real faults do.
  const clean = !problems.length && !setupWarnings.length;

  // Counted after merging, so the headline reflects distinct problems rather
  // than how many sheets each one touched.
  const issues = collapseIssues_(problems);

  // Opened BEFORE the result dialog on purpose: the dialog is modal, so it
  // lands on top and the sidebar is already sitting there once it is closed.
  // The other way round, the sidebar would appear behind something the user
  // has to dismiss first and would look like it failed to open.
  //
  // Only on a full run or a Black Temple one. Every string the hub exports
  // (Bloodboil, Souls, Council) belongs to Black Temple, so popping it after a
  // Mount Hyjal run would hand somebody assignments for bosses they are not
  // doing tonight. Listed the other way round — "open for these" rather than
  // "skip for MH" — so adding a third instance later stays quiet by default
  // instead of opening a panel nobody asked for. If a Mount Hyjal export is
  // ever added to T6_EXPORT_SOURCES, this condition needs revisiting.
  const exportsApplyHere = !scope || scope === 'bt';

  if (exportsApplyHere && typeof openT6ExportHubAfterRun_ === 'function') {
    openT6ExportHubAfterRun_(ss);
  }

  // Named in the headline so a half-run is never mistaken for a full one —
  // "all done" reading the same either way is how somebody walks into Black
  // Temple with nothing assigned.
  const scopeLabel = scope === 'mh' ? ' (Mount Hyjal only)'
    : scope === 'bt' ? ' (Black Temple only)' : '';

  showResult_(ss, {
    ok: clean && !autoPicked.length,
    autoPicked: autoPicked,
    headline: autoPicked.length
      ? `We do not know your roster, so tanks and healers were chosen for you`
      : (clean
        ? `All done: ${assignedBosses.length} assignments written in ${seconds}s${scopeLabel}`
        : `${assignedBosses.length} assignments written in ${seconds}s${scopeLabel}` +
          (issues.length ? `, ${issues.length} to look at` : '')),
    assigned: tally,
    skipped: skipped,
    issues: issues,
    notices: collapseIssues_(notices),
    notes: setupWarnings,
  });
}

/**
 * Archimonde's Decurse Assignments (BH17:BH21): Star, Triangle, Square, Tank
 * Group, Melee Group, in that order.
 *
 * Star/Triangle/Square are read directly from the marker groups EnsureEach
 * already wrote (AUTO_ASSIGN_ARCHIMONDE_SOURCE_RANGES) — never recomputed, so
 * whoever is actually standing in Star is always Star's decurse name. A
 * decurse ability is a short-GCD, long-duration thing (per Joey: 1.5s cast,
 * 5-minute curse), so one person can realistically cover more than one group
 * by moving between them — the only real limit is getting in range. That is
 * why, unlike every other row in this file, the same name is deliberately
 * allowed to repeat across multiple cells here.
 *
 * Any slot with nobody real, including Tank Group and Melee Group which have
 * no physical position of their own, goes to whichever decurser is covering
 * the fewest groups so far. That produces the evenest possible split (2
 * decursers is 3-2, 3 is 2-2-1, 4 is 2-1-1-1, 5+ is 1 each) while making sure
 * the second group anybody picks up goes to the spec you trust most, since
 * ties are broken by EnsureEach order.
 *
 * Who counts, and in what order, is read from the Archimonde row's EnsureEach
 * rather than hardcoded here. See the note in the body.
 *
 * A hand-set Preferred pick on the main Archimonde row can put a decurser
 * somewhere other than where this cycle would expect them — real placements
 * always win regardless, so accuracy for Star/Triangle/Square never slips,
 * even if that costs the leftover slots a slightly less even split that week.
 *
 * Silently does nothing if the Archimonde tab does not exist, so a spreadsheet
 * copy without this feature set up is unaffected.
 */
function assignArchimondeDecurse_(ss, allPlayers, problems, notices, encounters) {
  const sheet = ss.getSheetByName(AUTO_ASSIGN_ARCHIMONDE_SHEET);
  if (!sheet) return;

  // Who counts as a decurser comes from the Archimonde marker row's EnsureEach,
  // not from constants in here.
  //
  // Spec vocabulary differs between copies of this sheet. One guild's caster
  // druids read "Balance Druid", another's read "Dreamstate Druid", and a
  // hardcoded list silently misses the second: their only caster druid was
  // never counted and the mages quietly picked up the extra groups. Reading
  // the row also means this table and the per-group guarantee can never
  // disagree about who can decurse, since they now consult the same cell.
  //
  // Order matters and comes from the row too: the first entry is the most
  // preferred, exactly like AutoRoles everywhere else.
  const source = (encounters || []).filter(e =>
    e.sheetName === AUTO_ASSIGN_ARCHIMONDE_SHEET && e.ensureEach.length)[0];
  const entries = source ? source.ensureEach : AUTO_ASSIGN_ARCHIMONDE_FALLBACK_DECURSERS;

  if (!source) {
    Logger.log(`Archimonde Decurse: no EnsureEach on any Archimonde row, ` +
      `falling back to [${entries.join(', ')}]`);
  }

  const rankOf = player => eligibilityRank_(player, entries);
  const isDecurser = player => rankOf(player) !== -1;

  const playersByName = new Map(allPlayers.map(p => [p.player.toLowerCase(), p]));

  const real = AUTO_ASSIGN_ARCHIMONDE_SOURCE_RANGES.map(range => {
    const names = sheet.getRange(range).getDisplayValues().flat()
      .map(v => String(v).trim()).filter(Boolean);
    const decurser = names.map(n => playersByName.get(n.toLowerCase())).find(p => p && isDecurser(p));
    return decurser ? decurser.player : null;
  });

  const pool = allPlayers.filter(isDecurser).sort((a, b) => rankOf(a) - rankOf(b));

  Logger.log(`Archimonde Decurse: matching on [${entries.join(', ')}] found ` +
    `${pool.length} decurser(s): ${pool.map(p => `${p.player} (${p.classspec})`).join(', ') || 'nobody'}`);

  if (!pool.length) {
    // Not "no PI available" territory — a curse nobody removes on this fight
    // is usually a wipe, and it snowballs. This one stays red on purpose.
    problems.push({
      where: 'Archimonde Decurse Assignments',
      text: 'no Mage, Balance Druid or Restoration Druid is rostered, so nothing can decurse',
      fix: 'This fight cannot be safely done without a decurser. Get one rostered or reassign specs before the pull.',
    });
    return;
  }

  const cells = AUTO_ASSIGN_ARCHIMONDE_TARGET_CELLS;

  // Least loaded first, ties going to whoever ranks higher in EnsureEach.
  //
  // A plain round robin sent the extra groups to whoever happened to sit at
  // the next index, which in practice meant the two mages standing in markers
  // picked up both doubles while the one druid covered a single group. Mages
  // are famously unreliable at pressing the button, so when somebody has to
  // cover two groups it should be the spec you trust, and that is exactly the
  // order EnsureEach already states.
  //
  // Real placements are counted before any borrowing starts, so somebody
  // already standing in a marker is not handed a second group while an unused
  // decurser sits idle.
  const load = new Map(pool.map(p => [p.player, 0]));
  real.forEach(name => {
    if (name && load.has(name)) load.set(name, load.get(name) + 1);
  });

  const nextDecurser = () => {
    // pool is already in EnsureEach order, and strict "less than" keeps the
    // first of any tie, so priority breaks ties without a second sort.
    let pick = pool[0];
    pool.forEach(p => { if (load.get(p.player) < load.get(pick.player)) pick = p; });
    load.set(pick.player, load.get(pick.player) + 1);
    return pick;
  };

  const values = cells.map((cell, i) =>
    (i < real.length && real[i]) ? real[i] : nextDecurser().player);

  // Not a fault — the fight is winnable with one decurser covering several
  // groups, which is exactly what the cycling above arranges. But whoever is
  // calling the raid should know it before the pull rather than work it out
  // from the table, since it means real running about and range to manage.
  if (pool.length < cells.length) {
    const covers = {};
    values.forEach(name => { covers[name] = (covers[name] || 0) + 1; });
    const stretched = Object.keys(covers)
      .filter(name => covers[name] > 1)
      .map(name => `${name} has ${covers[name]}`);

    notices.push({
      info: true,
      text: `Only ${pool.length} decurser${pool.length === 1 ? '' : 's'} for ${cells.length} groups on Archimonde` +
        (stretched.length ? `, so ${stretched.join(' and ')}` : ''),
    });
  }

  // A formula or anything unexpected here is left alone rather than clobbered
  // — matches the formula guard every other row in this file gets.
  const blocked = [];
  cells.forEach((cell, i) => {
    const range = sheet.getRange(cell);
    if (range.getFormula()) {
      blocked.push(cell);
      return;
    }
    range.setValue(values[i]);
  });

  if (blocked.length) {
    pushProblem_(problems, 'Archimonde Decurse Assignments', {
      internal: true,
      text: `${blocked.join(', ')} still ${blocked.length === 1 ? 'has' : 'have'} a formula in it, so nothing was written there`,
      fix: 'Clear the formula in that cell and re-run to have it auto-filled too.',
    });
  }

  Logger.log(`Archimonde Decurse Assignments: ${cells.map((cell, i) => `${cell}=${values[i]}`).join(', ')}`);
}

/**
 * Assigns one boss. Returns { placed, warnings } on success, or { problem }
 * when this boss must be skipped — one bad boss never stops the rest of the run.
 */
function assignEncounter_(ctx) {
  const encounter = ctx.encounter;
  const sheet = encounter.sheet || ctx.ss.getSheetByName(encounter.sheetName);
  if (!sheet) {
    return { problem: {
      internal: true,
      text: `there is no tab called "${encounter.sheetName}"`,
      fix: 'Make Encounter match the tab name exactly, or add a Sheet column naming the tab. Set GroupAmount to 0 to skip it.',
    } };
  }

  // Checked before anything is read or written, so a phase that will not happen
  // costs nothing and leaves whatever is in its cells alone.
  if (encounter.skipWhen) {
    // Apps Script can queue a write from an earlier row and only actually run
    // it once something later forces a sync — which can be this read, on a
    // completely different sheet. Flushing here means a rejected write throws
    // right now, with a message that says so, instead of being read as if
    // THIS cell were the problem.
    try {
      SpreadsheetApp.flush();
    } catch (err) {
      return { problem: {
        internal: true,
        text: `a write from an earlier assignment in this run failed to save, and only ` +
          `surfaced here when this SkipWhen cell was read: ${err.message}`,
        fix: `This is not about ${encounter.skipWhen.cell}. Check ` +
          `${ctx.previousLabel ? `"${ctx.previousLabel}", the assignment run just before this one` : 'the assignment run just before this one'}. ` +
          `A name it tried to write was refused by a dropdown.`,
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
    if (check.matched) {
      return { skipped: `${encounter.skipWhen.cell} is "${check.actual}"` };
    }
  }

  let targetRanges = encounter.ranges;
  try {
    if (!targetRanges) targetRanges = resolveTargetRanges_(sheet, encounter);
  } catch (err) {
    return { problem: {
      internal: true,
      text: `the CellAnchor "${encounter.anchor}" could not be read`,
      fix: 'Use a single cell like BC10, or one per group: BC10, BH10, BM10.',
    } };
  }

  const formulaCell = findFormulaCell_(targetRanges, ctx.snapshot);
  if (formulaCell) {
    return { problem: {
      internal: true,
      text: `${formulaCell.cell} contains a formula, so none of this block was written`,
      fix: formulaCell.onAnchorRow
        ? 'CellAnchor is probably one row too high. Point it at the first player slot, below the healer.'
        : `That cell is inside the block but is not a plain dropdown. Clear the formula in ${formulaCell.cell}, or shorten GroupSize so the block stops above it.`,
    } };
  }

  const warnings = [];

  // AutoRoles both filters and ranks: keep the eligible, most preferred first.
  // Everything downstream consumes the pool in order, so this ordering decides
  // who gets a slot when there are more candidates than seats, and who lands in
  // the bottom (backup) rows.
  let pool = ctx.allPlayers
    .map(p => ({ player: p, rank: eligibilityRank_(p, encounter.eligibleRoles) }))
    .filter(entry => entry.rank !== -1)
    .sort((a, b) => a.rank - b.rank)
    .map(entry => entry.player);

  // An AutoRoles entry that matches nobody is only worth reporting when it is
  // not a real term either — "Restoration Druid" matching nobody just means
  // none is raiding, which happens most weeks and needs no attention.
  const unmatched = encounter.eligibleRoles.filter(entry =>
    !ctx.allPlayers.some(p => eligibilityRank_(p, [entry]) === 0) &&
    !isKnownTerm_(ctx.knownTerms, entry));
  if (unmatched.length) {
    warnings.push({
      internal: true,
      text: `AutoRoles lists ${unmatched.map(e => `"${e}"`).join(', ')}, which is not a class, spec or role`,
      fix: 'Check the spelling against Class/Spec Data: columns A, B and F list every valid term.',
    });
  }

  // Tracked so an empty pool can say which stage emptied it — "nobody matches"
  // and "they all matched but were already placed" need opposite fixes.
  const matchedAutoRoles = pool.map(p => p.player);

  // Someone another row already spent. The mage kiting Zerevor is on the sheet
  // and rostered and a perfectly good interrupter — they are simply busy, which
  // nothing else in the config can express.
  const busyElsewhere = [];
  encounter.excludeFrom.forEach(label => {
    const key = `${label.toLowerCase()}||${encounter.sheetName.toLowerCase()}`;
    const names = ctx.placedNames[key];
    if (names === undefined) {
      warnings.push({
        internal: true,
        text: `ExcludeFrom "${label}" did not run before this row, so nobody was excluded`,
        fix: 'Move that row above this one in AutoAssignData, since the exclusion reads what it assigned.',
      });
      return;
    }
    names.forEach(name => busyElsewhere.push(name.toLowerCase()));
  });

  // Tracked separately from busyElsewhere itself: that list can include names
  // ExcludeFrom claimed that were never in this row's pool to begin with (a
  // healer excluded from a DPS-only row, say), which would make an empty pool
  // look like it happened for the wrong reason.
  const removedAsBusy = [];
  if (busyElsewhere.length) {
    pool = pool.filter(p => {
      if (busyElsewhere.indexOf(p.player.toLowerCase()) === -1) return true;
      removedAsBusy.push(p.player);
      return false;
    });
  }

  // Anyone already placed above the anchor (the healer row) is spoken for —
  // unless this row sits under something unrelated and says so, in which case
  // the scan is skipped entirely rather than filtered afterwards, since there
  // is nothing to learn from reading a neighbouring table.
  const removedAsPlaced = [];
  if (!encounter.ignoreAbove) {
    const prePlaced = loadPrePlacedNames_(sheet, targetRanges, ctx.ownedCells, ctx.snapshot);
    pool = pool.filter(p => {
      const key = p.player.toLowerCase();
      if (!prePlaced.has(key)) return true;
      removedAsPlaced.push(`${p.player} in ${prePlaced.get(key)}`);
      return false;
    });
  }

  const rules = ctx.allRules.filter(r =>
    r.encounter === '*' || r.encounter === encounter.encounter.trim().toLowerCase());

  // MaxPerGroup = 0 means "keep these players out of this encounter entirely".
  const hardExclusionRules = rules.filter(r => r.max === 0);
  const removedByHardExclusion = [];
  pool = pool.filter(p => {
    if (!hardExclusionRules.some(rule => matchesRule_(rule, p))) return true;
    removedByHardExclusion.push(p.player);
    return false;
  });

  Logger.log(`${encounter.encounter} on ${encounter.sheetName}: AutoRoles matched ` +
    `[${matchedAutoRoles.join(', ') || 'nobody'}], busyElsewhere removed [${busyElsewhere.join(', ') || 'nobody'}], ` +
    `prePlaced removed [${removedAsPlaced.join(', ') || 'nobody'}], hardExclusion removed ` +
    `[${removedByHardExclusion.join(', ') || 'nobody'}], ` +
    `final pool going into seeding [${pool.map(p => p.player).join(', ') || 'EMPTY'}], ` +
    `totemParties=${encounter.totemParties}, groupCount=${encounter.groupCount}, ` +
    `groupSizes=[${encounter.groupSizes.join(', ')}]`);

  if (!pool.length) {
    const wanted = encounter.eligibleRoles.length
      ? ` "${encounter.eligibleRoles.join(', ')}"` : '';
    // How to say the same thing to somebody who has never opened the config.
    const needs = encounter.eligibleRoles.length
      ? encounter.eligibleRoles.map(e => e.replace(/^(any|flex)\s+/i, '')).join(' or ')
      : 'anyone';

    if (!matchedAutoRoles.length) {
      // A spelling mistake in AutoRoles is still caught above (the "unmatched"
      // check) regardless of Critical — this is specifically the "the term is
      // fine, nobody just has that class/spec this week" case, which is normal
      // roster variance unless the row says otherwise.
      if (!encounter.critical) {
        return { skipped: `nobody on the roster matches AutoRoles${wanted}` };
      }
      // Critical means the raid cannot do the fight, which is the raid's
      // problem and not the sheet's — so this one stays in plain words.
      return { problem: {
        text: `no ${needs} is rostered, and this fight needs one`,
        fix: 'Get someone who can do it into the raid, or go in knowing this is uncovered.',
      } };
    }
    if (removedAsBusy.length) {
      const claimedBy = encounter.excludeFrom.join(', ');
      // A Selection row picking from a pool ExcludeFrom already emptied out is
      // the normal "no spares this week" case, not a misconfiguration — e.g.
      // every decurser already claimed by the ranged groups, most weeks.
      if (encounter.selection || !encounter.critical) {
        return { skipped:
          `${removedAsBusy.join(', ')} match AutoRoles${wanted}, but ${claimedBy} already used them — no spares this week` };
      }
      return { problem: {
        text: `${removedAsBusy.join(', ')} could cover this, but ${claimedBy} already used them. ` +
          `Nobody is free, and this fight needs it`,
        fix: 'Roster another one, or take somebody off that other job.',
      } };
    }
    if (removedAsPlaced.length) {
      return { problem: {
        internal: true,
        text: `${removedAsPlaced.join(', ')} matched AutoRoles${wanted}, but that cell ` +
          `sits just above this block, so they count as already placed`,
        fix: 'Clear that cell, move the block down, or point CellAnchor somewhere the name is not directly above.',
      } };
    }
    return { problem: {
      internal: true,
      text: `${matchedAutoRoles.join(', ')} match AutoRoles${wanted}, but a GroupRules ` +
        `row with MaxPerGroup 0 excluded them all`,
      fix: 'Check GroupRules for a rule excluding these players from this encounter.',
    } };
  }

  const activeRules = rules.filter(r => r.max !== 0).sort((a, b) => a.priority - b.priority);

  // EnsureEach is a MinPerGroup rule with the ceremony removed. Placed first so
  // "every group needs a decurser" is satisfied before anything optional.
  if (encounter.ensureEach.length) {
    activeRules.unshift({
      matchType: 'ensure',
      entries: encounter.ensureEach,
      matchValue: encounter.ensureEach.join(', '),
      min: 1,
      max: Infinity,
      priority: -1,
    });
  }
  const assigned = new Set();

  // Keeping people on the marker they already know beats any optimisation, so
  // inherited placements are seeded before clustering or rules get a say.
  let inheritedGroups = null;
  if (encounter.inheritFrom) {
    const inherited = readInheritedGroups_(ctx.ss, ctx.allEncounters, encounter.inheritFrom);
    if (inherited.error) {
      warnings.push({
        internal: true,
        text: `nothing was inherited: ${inherited.error}`,
        fix: 'Rows sharing a tab need a Sheet column naming it, with Encounter as their label.',
      });
    } else {
      inheritedGroups = inherited;
    }
  }

  const seedCtx = {
    label: encounter.encounter,
    allowDuplicates: encounter.allowDuplicates,
    pool: pool,
    allPlayers: ctx.allPlayers,
    assigned: assigned,
    groupSizes: encounter.groupSizes,
    markerNames: encounter.markerNames,
  };
  const preSeeded = Array.from({ length: encounter.groupCount }, () => []);

  // Your explicit picks go in before anything is worked out.
  const manual = seedPreferred_(seedCtx, preSeeded, encounter.preferred, warnings);
  if (manual.missing.length) {
    // Says the same thing without naming the Preferred column. Somebody
    // reading this before a pull knows what "picked for this job" means; they
    // have no idea what a Preferred column is, and should not have to.
    const one = manual.missing.length === 1;
    warnings.push({
      info: true,
      text: `${manual.missing.join(', ')} ${one ? 'is' : 'are'} picked for this job but ` +
        `${one ? 'is' : 'are'} not raiding tonight, so it was filled automatically instead`,
      fix: 'Nothing to do unless you expected them to be in.',
    });
  }

  const keptCount = seedFromInheritance_(seedCtx, preSeeded, inheritedGroups);

  // Minimums are promises, party clustering is an optimisation, so the promises
  // are kept first. Run the other way round, clustering hands the whole caster
  // party to one marker and the decurse rule finds nobody left to spread.
  const guaranteed = seedGroupsByRules_(encounter.groupCount, pool, activeRules, assigned,
    preSeeded, encounter.groupSizes);

  const groups = encounter.totemParties
    ? seedGroupsByTotemParties_({
        preSeeded: guaranteed,
        sheet: sheet,
        label: encounter.encounter,
        targetRanges: targetRanges,
        groupCount: encounter.groupCount,
        pool: pool,
        ownedCells: ctx.ownedCells,
        snapshot: ctx.snapshot,
        playersByName: ctx.playersByName,
        partyByPlayer: ctx.partyByPlayer,
        groupSizes: encounter.groupSizes,
        balanceGroups: encounter.balanceGroups,
        assigned: assigned,
        warnings: warnings,
      })
    : guaranteed;

  fillRemaining_(groups, pool, activeRules, assigned, warnings,
    encounter.groupSizes, encounter.markerRoles, encounter.selection, encounter.encounter);

  Logger.log(`${encounter.encounter} on ${encounter.sheetName}: groups after fillRemaining_ = ` +
    groups.map((g, i) => `[${i}: ${g.map(p => p.player).join(', ') || 'empty'}]`).join(' '));

  // Checked after everything is placed, since the general fill may well have
  // covered a group the rule could not. Running short here is quiet otherwise:
  // the rule simply stops when it runs out of people, and a group ends up
  // unable to do the thing the row exists to guarantee.
  if (encounter.ensureEach.length) {
    const rule = { matchType: 'ensure', entries: encounter.ensureEach };
    const uncovered = groups
      .map((group, i) => ({ i: i, ok: group.some(p => matchesRule_(rule, p)) }))
      .filter(entry => !entry.ok)
      .map(entry => markerLabel_(encounter, entry.i));

    if (uncovered.length) {
      // "no any mage or any druid" is how the raw entries read. The prefixes
      // change how matching works, not what the thing is called.
      const wanted = encounter.ensureEach
        .map(entry => entry.replace(/^(any|flex)\s+/i, ''))
        .join(' or ');

      // Same Critical switch as the empty-pool checks: a group missing its
      // guaranteed member is only a red "we cannot do this" when nothing else
      // backstops it.
      //
      // On Archimonde something does: the Decurse Assignments table names a
      // person for every marker whether or not one is standing there. Saying
      // only "no mage or druid in Square" reads like nobody is covering it,
      // which is the opposite of what the sheet now says — so name the
      // backstop instead of raising an alarm about a hole that was filled.
      const backstopped = encounter.sheetName === AUTO_ASSIGN_ARCHIMONDE_SHEET;

      if (backstopped) {
        // Log only. assignArchimondeDecurse_ already reports the shortage and
        // who is covering how many, which says everything this would — two
        // lines about one shortage is one too many.
        Logger.log(`${encounter.encounter}: no ${wanted} standing in ${uncovered.join(', ')}, ` +
          `covered from elsewhere by the Decurse Assignments table`);
      } else {
        warnings.push({
          info: !encounter.critical,
          text: `no ${wanted} in ${uncovered.join(', ')}`,
          fix: encounter.critical
            ? 'Not enough of them raiding to cover every group, and this row is marked Critical.'
            : 'Not enough of them raiding to cover every group.',
        });
      }
    }
  }

  noteMarkerRoleMisfits_(groups, encounter, warnings);

  if (keptCount) {
    // Log only: this is inheritance working exactly as configured, and a panel
    // that announces its own successes buries the things that need a person.
    Logger.log(`${encounter.encounter}: kept ${keptCount} player(s) on the same marker ` +
      `as "${encounter.inheritFrom}"`);
  } else if (inheritedGroups) {
    // Two very different situations reach here. If the source row was skipped
    // by its own SkipWhen, that phase is not happening tonight — there is
    // nothing to inherit and nothing to think about, so saying so every week
    // is the same noise as announcing the skip itself. If it actually ran and
    // still came back empty, something is wrong and silence would read as
    // "inheritance is broken" months later.
    const sourceKey =
      `${encounter.inheritFrom.trim().toLowerCase()}||${encounter.sheetName.toLowerCase()}`;

    if (!(ctx.skippedRows && ctx.skippedRows[sourceKey])) {
      // "Inherit" is the column's word, not a raider's. What actually happened
      // is that people did not keep the spots they had in the other phase, and
      // that is the bit worth knowing before a pull.
      warnings.push({
        info: true,
        text: `people did not keep their positions from "${encounter.inheritFrom}". ` +
          `Nothing is filled in there, so fresh spots were picked here`,
        fix: 'Normal if that part has not been assigned yet.',
      });
    }
  }

  // A target block should never promise more than there are casters to deliver.
  // Two Power Infusion targets with one Discipline Priest raiding means the
  // second person is told they are getting something that will not arrive.
  if (encounter.limitTo) {
    const key = `${encounter.limitTo.toLowerCase()}||${encounter.sheetName.toLowerCase()}`;
    const cap = ctx.placementCounts[key];

    if (cap === undefined) {
      // Two ways to get here, and only one is a fault.
      //
      // The row we cap against was skipped, which happens the moment nobody
      // can do that job: take the Discipline Priest out of the roster and the
      // Power Infusion row finds nobody, so every PI Target row hanging off it
      // lands here. That is a roster fact, not a broken sheet, and reporting
      // it as one produced fourteen red lines telling somebody to go fix a
      // config that was perfectly fine.
      //
      // Anything else means the row genuinely has not run yet, which is a
      // real ordering mistake in AutoAssignData and stays reportable.
      if (ctx.skippedRows && ctx.skippedRows[key]) {
        // Cap it at nothing rather than leaving it uncapped: promising Power
        // Infusion when no priest is raiding is exactly what LimitTo exists
        // to prevent, and a skipped source means the count really is zero.
        groups.forEach(group => { group.length = 0; });
        Logger.log(`${encounter.encounter}: "${encounter.limitTo}" was skipped, so this was left empty`);
      } else {
        warnings.push({
          internal: true,
          text: `LimitTo "${encounter.limitTo}" did not run before this row, so no limit was applied`,
          fix: 'Move that row above this one in AutoAssignData, since the limit reads what it assigned.',
        });
      }
    } else {
      const before = groups.reduce((n, g) => n + g.length, 0);
      let remaining = cap;
      groups.forEach(group => {
        const keep = Math.max(0, Math.min(group.length, remaining));
        group.length = keep;
        remaining -= keep;
      });

      // On a selection row an empty slot is the expected shape of the roster —
      // one Disc priest means one PI, every week — so saying so is just noise.
      if (before > cap && !encounter.selection) {
        warnings.push({
          info: true,
          text: `${before - cap} slot(s) left empty: only ${cap} "${encounter.limitTo}" to go round`,
          fix: 'Deliberate: better an empty cell than promising something nobody can cast.',
        });
      }
    }
  }

  noteFallbacksInMainSlots_(groups, encounter, warnings);

  // writeGroups_ writes first and only investigates if the sheet refuses it.
  const written = writeGroups_(targetRanges, groups, ctx.snapshot, encounter.encounter);

  if (written.rejected.length) {
    const names = written.rejected.slice(0, 6).join(', ') +
      (written.rejected.length > 6 ? ` and ${written.rejected.length - 6} more` : '');
    const one = written.rejected.length === 1;
    return { problem: {
      text: `a dropdown here would not accept ${names}, so this block was left as it was`,
      fix: `${one ? 'That name is' : 'Those names are'} missing from the list of options the cell allows. ` +
        `Click the cell, check Data > Data validation, and add ${one ? 'the name' : 'them'} to its list.`,
    } };
  }

  // Once per run per target. A row spanning twelve sheets would otherwise check
  // the same range twelve times to discover eleven times that it is already
  // full — and they would all be writing identical names anyway.
  if (encounter.writeBackTo) {
    const key = encounter.writeBackTo.toLowerCase();
    if (!ctx.writtenBack[key]) {
      ctx.writtenBack[key] = true;
      writeBackPicks_(ctx.ss, encounter.writeBackTo, pool, warnings, ctx.writeBackFilled);
    }
  }

  // Deliberately NOT internal: who is missing from a group matters on the
  // pull, even though the remedy is a config change somebody else makes.
  written.overflow.forEach(o => warnings.push({
    text: `${o.dropped.join(', ')} had nowhere to go in ${markerLabel_(encounter, o.group - 1)}. ` +
      `It only has room for ${o.capacity}`,
    fix: 'That table needs more rows before everyone fits.',
  }));

  const placedNames = [];
  groups.forEach(group => group.forEach(player => placedNames.push(player.player)));

  return { placed: placedNames.length, names: placedNames, warnings: warnings };
}

/**
 * One read of each sheet, covering every block it holds plus the rows scanned
 * above them, instead of two reads per range.
 *
 * Safe to take once at the start because nothing this run alters what these
 * reads look for: target cells are checked for formulas and we only ever write
 * plain values, and the rows above are either cells we own (skipped anyway) or
 * formula healer rows that nothing here touches.
 */
function prepareEncounters_(ss, encounters) {
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
      return; // reported later by the row that owns it
    }

    // Resolving a range is not free, and this used to happen three times per
    // encounter — here, again for the snapshot boxes, and again when assigning.
    encounter.sheet = sheet;
    encounter.ranges = ranges;

    if (!owned[encounter.sheetName]) owned[encounter.sheetName] = {};
    const forSheet = owned[encounter.sheetName];

    ranges.forEach(range => {
      const column = range.getColumn();
      const firstRow = range.getRow();
      const numRows = range.getNumRows();

      for (let i = 0; i < numRows; i++) forSheet[`${column},${firstRow + i}`] = true;

      const top = Math.max(1, firstRow - AUTO_ASSIGN_HEALER_SCAN_ROWS);
      const bottom = firstRow + numRows - 1;
      const box = boxes[encounter.sheetName];

      if (!box) {
        boxes[encounter.sheetName] =
          { minRow: top, maxRow: bottom, minCol: column, maxCol: column };
        return;
      }
      box.minRow = Math.min(box.minRow, top);
      box.maxRow = Math.max(box.maxRow, bottom);
      box.minCol = Math.min(box.minCol, column);
      box.maxCol = Math.max(box.maxCol, column);
    });
  });

  const snapshots = {};
  Object.keys(boxes).forEach(name => {
    const box = boxes[name];
    const rows = box.maxRow - box.minRow + 1;
    const cols = box.maxCol - box.minCol + 1;

    if (rows * cols > AUTO_ASSIGN_MAX_SNAPSHOT_CELLS) {
      Logger.log(`"${name}" spans ${rows}x${cols} cells — reading its ranges individually instead`);
      return;
    }

    const range = ss.getSheetByName(name).getRange(box.minRow, box.minCol, rows, cols);
    snapshots[name] = {
      firstRow: box.minRow,
      firstCol: box.minCol,
      values: timed_('scanning', () => range.getDisplayValues()),
      formulas: timed_('guarding', () => range.getFormulas()),
    };
  });

  return { owned: owned, snapshots: snapshots };
}

/** A cell from a snapshot grid, or null when it falls outside. */
function snapshotCell_(snapshot, grid, row, column) {
  if (!snapshot) return null;

  const r = row - snapshot.firstRow;
  const c = column - snapshot.firstCol;
  const cells = snapshot[grid];

  if (r < 0 || c < 0 || r >= cells.length || c >= cells[0].length) return null;
  return cells[r][c];
}

/**
 * Every cell this script writes, keyed by sheet then "column,row".
 *
 * Lets the pre-placed check tell "a name another system put here" (a formula
 * healer row, a hand-picked tank) apart from "a name we wrote ourselves from
 * another row", which must not count as someone already being spoken for.
 */
/**
 * Reads the groups another Encounters row has already written, so this one can
 * keep the same people on the same markers.
 *
 * Matched by marker NAME, not position: a Demon Phase row with five markers
 * inherits Star/Square/Triangle from a three-marker phase row and leaves its
 * extra markers to fill normally. Falls back to matching by position when
 * MarkerNames is not set on both rows.
 *
 * Reads the sheet, so the source row must sit ABOVE this one in Encounters —
 * otherwise it has not been assigned yet this run and the values are stale.
 */
function readInheritedGroups_(ss, encounters, label) {
  const wanted = label.trim().toLowerCase();
  const source = encounters.filter(e => e.encounter.trim().toLowerCase() === wanted)[0];

  // Each failure gets its own message — "row not found" and "row found but its
  // tab is missing" need very different fixes, and reporting both as the former
  // sends you looking in the wrong place.
  if (!source) {
    return { error: `no Encounters row is named "${label}"` };
  }
  if (!source.groupCount || !source.anchors.length) {
    return { error: `"${label}" has no groups configured to copy` };
  }

  const sheet = ss.getSheetByName(source.sheetName);
  if (!sheet) {
    return { error: `"${label}" writes to a tab called "${source.sheetName}", which does not exist` };
  }

  let ranges;
  try {
    ranges = resolveTargetRanges_(sheet, source);
  } catch (err) {
    return { error: `"${label}" has a CellAnchor that could not be read` };
  }

  return {
    markerNames: source.markerNames,
    groups: ranges.map(range => range.getDisplayValues()
      .flat()
      .map(value => String(value).trim())
      .filter(value => value !== '')),
  };
}

/**
 * Places the people named in Preferred, in order, before anything is computed.
 *
 * A human choice beats an algorithm picking arbitrarily among five shamans, so
 * these bypass AutoRoles entirely — being named IS the qualification. Anyone
 * not raiding tonight is skipped and their slot falls through to the normal
 * assignment, which is the whole point: the manual setup degrades to automatic
 * rather than breaking.
 *
 * Fills group by group, which is what single-group rows (Bloodlust, PI) want.
 */
function seedPreferred_(ctx, groups, preferred, warnings) {
  if (!preferred.length) return { placed: 0, missing: [] };

  const byName = new Map(ctx.allPlayers.map(p => [p.player.toLowerCase(), p]));
  const missing = [];
  // Every name gets an outcome, including the two that used to vanish without
  // trace. "I changed my pick and nothing happened" is unanswerable otherwise,
  // and the usual cause is the quietest one: more names in the range than the
  // block has rows, so anything past the end is read and then dropped.
  const outcomes = [];
  let placed = 0;
  let groupIndex = 0;

  preferred.forEach(name => {
    const player = byName.get(name.trim().toLowerCase());
    if (!player) {
      missing.push(name);
      outcomes.push(`${name} is not rostered`);
      return;
    }
    // A repeat is normally a mistake, but on a duty row it is the point: the
    // same person really can be given two innervates or two Power Infusions.
    if (ctx.assigned.has(player.player) && !ctx.allowDuplicates) {
      outcomes.push(`${player.player} was already placed by an earlier pick`);
      return;
    }

    while (groupIndex < groups.length && groups[groupIndex].length >= ctx.groupSizes[groupIndex]) {
      groupIndex++;
    }
    if (groupIndex >= groups.length) {
      outcomes.push(`${player.player} had no room left, every slot was full`);
      return;
    }

    groups[groupIndex].push(player);
    ctx.assigned.add(player.player);
    placed++;
    outcomes.push(`${player.player} placed in group ${groupIndex + 1}`);
  });

  const capacity = ctx.groupSizes.reduce((total, size) => total + size, 0);
  Logger.log(`${ctx.label || 'this row'} Preferred: asked for ` +
    `${preferred.length} name(s) [${preferred.join(', ')}] into ${capacity} slot(s). ` +
    outcomes.join('. '));

  return { placed: placed, missing: missing };
}

/** Seeds each group with the players the inherited assignment had on that marker. */
function seedFromInheritance_(ctx, groups, inherited) {
  if (!inherited) return 0;

  const byName = new Map(ctx.pool.map(p => [p.player.toLowerCase(), p]));
  const useNames = ctx.markerNames.length && inherited.markerNames.length;
  let kept = 0;

  groups.forEach((group, i) => {
    let sourceIdx = i;
    if (useNames) {
      const marker = (ctx.markerNames[i] || '').trim().toLowerCase();
      sourceIdx = inherited.markerNames
        .map(n => n.trim().toLowerCase())
        .indexOf(marker);
    }
    if (sourceIdx === -1 || !inherited.groups[sourceIdx]) return;

    inherited.groups[sourceIdx].forEach(name => {
      const player = byName.get(name.toLowerCase());
      if (!player || ctx.assigned.has(player.player)) return;
      if (group.length >= ctx.groupSizes[i]) return;
      group.push(player);
      ctx.assigned.add(player.player);
      kept++;
    });
  });

  return kept;
}

/**
 * Reports anyone standing in a marker their role does not suit — the roster was
 * lopsided enough that the preference could not be honoured. Not a fault, but
 * it explains why someone is somewhere unexpected.
 */
function noteMarkerRoleMisfits_(groups, encounter, warnings) {
  const roles = encounter.markerRoles;
  if (!roles.length) return;

  const misfits = [];
  groups.forEach((group, i) => {
    if (!roles[i]) return;
    group.forEach(player => {
      if (markerSuits_(player, roles[i])) return;

      // Someone no marker was ever meant for — a Tank where the preferences
      // only name Ranged and Melee — is filling a gap, not being displaced.
      // Reporting them would put the same names in the dialog every week.
      const hasSuitableMarker = roles.some(role => markerSuits_(player, role));
      if (!hasSuitableMarker) return;

      misfits.push(`${player.player} in ${markerLabel_(encounter, i)}`);
    });
  });
  if (!misfits.length) return;

  // Log only. MarkerRoles is a soft preference by design — people overflow
  // into the "wrong" marker rather than being left out — so this fires on any
  // lopsided roster, which is most of them. It was also the longest block on
  // the panel, a dozen names deep, pushing everything that mattered off the
  // screen to say "this went the way it is supposed to".
  Logger.log(`${encounter.encounter} did not fit the marker roles: ${misfits.join(', ')}`);
}

/** "Star" if MarkerNames is set, otherwise "group 3". */
function markerLabel_(encounter, index) {
  return encounter.markerNames[index] || `group ${index + 1}`;
}

/**
 * Flags when a fallback pick ended up in a main slot rather than the backup row.
 *
 * On a fight like Bloodboil the bottom row is a deliberate reserve — melee down
 * there is fine. Melee in the rows above it means there were not enough of the
 * first-choice group to fill the slots that actually perform the mechanic, which
 * is something the raid leader wants to know before the pull.
 *
 * "First choice" is the first AutoRoles entry; anything matched by a later entry
 * counts as a fallback. Does nothing unless BackupRows is set.
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

/* ------------------------------------------------------------------ *
 * Group building
 * ------------------------------------------------------------------ */

/**
 * Totem clustering. Each marker is bound to a raid party so that party's
 * shaman keeps buffing them, then filled from that party's eligible players.
 */
function seedGroupsByTotemParties_(ctx) {
  const groups = ctx.preSeeded || Array.from({ length: ctx.groupCount }, () => []);
  const markerHealers =
    findMarkerHealers_(ctx.sheet, ctx.targetRanges, ctx.ownedCells, ctx.snapshot);

  const boundParties = new Set();
  const partyOf = player => ctx.partyByPlayer.get(player.player.toLowerCase()) || null;
  const plan = new Array(ctx.groupCount).fill(null);

  // (a) Which markers already have totems standing on them.
  //
  // A Shaman in the healer row above a marker means that marker is covered,
  // and covered markers are then FREE to take whichever party keeps the most
  // people together. That is the important difference from how this used to
  // work: it bound such a marker to its healer's party, which locked the two
  // markers with resto shaman healers onto those parties and left whatever
  // remained to be scattered. A marker that already has its totems does not
  // need a party for totem reasons at all.
  const covered = markerHealers.map((names, i) => {
    const label = `${ctx.label} marker ${i + 1}`;

    const healers = names
      .map(name => ctx.playersByName.get(name.toLowerCase()))
      .filter(player => player);

    const shamanHealer = healers.filter(p => p.class.toLowerCase() === 'shaman')[0];
    if (shamanHealer) {
      Logger.log(`${label}: totems already covered by healer ${shamanHealer.player}, ` +
        `so it is free to take any party`);
      return true;
    }

    Logger.log(`${label}: no Shaman in the healer row` +
      (healers.length ? ` (${healers.map(p => `${p.player} (${p.classspec})`).join(', ')})` : '') +
      `, so it needs a shaman from the pool`);
    return false;
  });

  // Who is still available, grouped by their raid party, and which parties can
  // supply a totem shaman for this row. A shaman not in the pool cannot anchor
  // anything here, which is what keeps an Enhancement shaman out of a ranged
  // marker and a caster shaman out of a melee one.
  const membersByParty = new Map();
  const shamanByParty = new Map();

  ctx.pool.forEach(player => {
    if (ctx.assigned.has(player.player)) return;
    const party = partyOf(player);
    if (!party) return;

    if (!membersByParty.has(party)) membersByParty.set(party, []);
    membersByParty.get(party).push(player);

    if (player.class.toLowerCase() === 'shaman' && !shamanByParty.has(party)) {
      shamanByParty.set(party, player);
    }
  });

  // The party with the most people still available, so cohesion is maximised
  // rather than left to whoever happened to be first in pool order.
  const claimBestParty = requireShaman => {
    let best = null;
    let bestCount = -1;

    membersByParty.forEach((members, party) => {
      if (boundParties.has(party)) return;
      if (requireShaman && !shamanByParty.has(party)) return;
      if (members.length > bestCount) {
        best = party;
        bestCount = members.length;
      }
    });

    if (best !== null) boundParties.add(best);
    return best;
  };

  // (b) Uncovered markers choose first. Their options are the narrower ones,
  // since only a party carrying a usable shaman is any good to them, and a
  // covered marker can happily take whatever is left.
  markerHealers.forEach((names, i) => {
    if (covered[i]) return;
    const label = `${ctx.label} marker ${i + 1}`;

    const party = claimBestParty(true);
    if (party === null) {
      Logger.log(`${label}: no unclaimed party has a shaman available in this pool`);
      return;
    }

    plan[i] = party;

    // Seed the shaman explicitly rather than trusting the fill below to reach
    // them. They may sit late in pool order, and a marker that misses its own
    // totem carrier defeats the point of binding it at all.
    const shaman = shamanByParty.get(party);
    if (shaman && !ctx.assigned.has(shaman.player)) {
      groups[i].push(shaman);
      ctx.assigned.add(shaman.player);
    }

    Logger.log(`${label}: bound to party ${party} (${membersByParty.get(party).length} available) ` +
      `by seeding ${shaman ? shaman.player : 'nobody'}`);
  });

  // (c) Covered markers, and any uncovered one that found no shaman, take the
  // largest parties still going. No shaman requirement here: keeping a party
  // together is worth doing even when nobody in it drops totems.
  markerHealers.forEach((names, i) => {
    if (plan[i]) return;
    const label = `${ctx.label} marker ${i + 1}`;

    const party = claimBestParty(false);
    if (party === null) {
      Logger.log(`${label}: no unclaimed party left to bind to`);
      return;
    }

    plan[i] = party;
    Logger.log(`${label}: bound to party ${party} ` +
      `(${membersByParty.get(party).length} available) for cohesion`);
  });

  // (d) Fill each bound marker from its party.
  //
  // With BalanceGroups on, a party can only take an even share of the slots.
  // That matters when one party holds most of the casters: without a cap it
  // swallows a whole marker and leaves the others nearly empty.
  // Share of the total seats this group is entitled to, scaled by its own size
  // so a 3-slot marker is not asked to hold as many as a 5-slot one.
  const totalSeats = ctx.groupSizes.reduce((n, c) => n + c, 0);
  const capOf = i => ctx.balanceGroups
    ? Math.max(1, Math.min(ctx.groupSizes[i],
        Math.ceil(ctx.pool.length * (ctx.groupSizes[i] / totalSeats))))
    : ctx.groupSizes[i];

  const spill = [];
  plan.forEach((party, i) => {
    if (!party) return;
    ctx.pool.forEach(p => {
      if (ctx.assigned.has(p.player)) return;
      if (partyOf(p) !== party) return;
      if (groups[i].length >= capOf(i)) {
        spill.push({ player: p, from: i });
        return;
      }
      groups[i].push(p);
      ctx.assigned.add(p.player);
    });
  });

  // Party-mates who did not fit go to the NEAREST marker with room, not the
  // emptiest. Markers are ordered by position, so the neighbour is the group
  // most likely to still be inside their own shaman's totem range.
  spill.forEach(entry => {
    const candidates = groups.map((g, i) => i).filter(i => groups[i].length < capOf(i));
    if (!candidates.length) return; // fillRemaining_ will find them a seat

    const target = candidates.sort((a, b) =>
      Math.abs(a - entry.from) - Math.abs(b - entry.from) ||
      groups[a].length - groups[b].length)[0];

    groups[target].push(entry.player);
    ctx.assigned.add(entry.player.player);
  });

  if (!ctx.partyByPlayer.size) {
    ctx.warnings.push({
      internal: true,
      text: `no raid party data was found, so groups were not built around shamans`,
      fix: `Check the "${AUTO_ASSIGN_DYNAMIC_LISTS_SHEET}" sheet has columns headed ActiveRosterG1 to ActiveRosterG5.`,
    });
  } else {
    const unbound = plan.reduce((n, party) => n + (party ? 0 : 1), 0);
    if (unbound) {
      // Log only: more markers than parties with anyone left in them is the
      // ordinary case, and those markers still get filled by fillRemaining_.
      // Nothing here for a person to decide.
      Logger.log(`${ctx.label}: ${unbound} of ${ctx.groupCount} marker(s) had no party left ` +
        `to bind to and were filled with whoever remained`);
    }
  }

  return groups;
}

/** Rules mode: satisfy each rule's MinPerGroup, highest priority first. */
function seedGroupsByRules_(groupCount, pool, activeRules, assigned, preSeeded, capacities) {
  const groups = preSeeded || Array.from({ length: groupCount }, () => []);
  const roomIn = i => !capacities || groups[i].length < capacities[i];

  activeRules.forEach((rule, ruleIndex) => {
    if (!rule.min) return;
    const laterRules = activeRules.slice(ruleIndex + 1);

    let safety = 0;
    while (safety++ < groupCount * pool.length + 1) {
      const needy = groups
        .map((g, i) => ({ i, count: countMatching_(g, rule) }))
        .filter(({ i, count }) =>
          count < rule.min && roomIn(i) && !isGroupAtMax_(groups[i], rule))
        .sort((a, b) => a.count - b.count || groups[a.i].length - groups[b.i].length)[0];
      if (!needy) break;

      const candidates = pool.filter(p => !assigned.has(p.player) && matchesRule_(rule, p));
      if (!candidates.length) break;

      // Spend the least flexible players first, saving multi-rule players for later rules.
      candidates.sort((a, b) => countOtherMatches_(a, laterRules) - countOtherMatches_(b, laterRules));
      groups[needy.i].push(candidates[0]);
      assigned.add(candidates[0].player);
    }
  });

  return groups;
}

/**
 * Everyone still unplaced goes to the smallest group that won't breach a Max.
 *
 * Groups are never filled past capacity, so when there are more eligible
 * players than seats the ones left out are the last in pool order — i.e. the
 * least preferred by AutoRoles, not whoever happened to sort last.
 */
function fillRemaining_(groups, pool, activeRules, assigned, warnings, capacities, markerRoles,
    isSelection, label) {
  const leftover = pool.filter(p => !assigned.has(p.player));
  if (!leftover.length) return;

  const roles = markerRoles || [];

  const totalSeats = capacities.reduce((n, c) => n + c, 0);
  const total = groups.reduce((n, g) => n + g.length, 0) + leftover.length;
  const allIdx = groups.map((g, i) => i);
  const smallestOf = idx => idx.sort((a, b) => groups[a].length - groups[b].length)[0];

  // Balance by how full each group is relative to its own size, so a 3-slot
  // marker is not judged against a 5-slot one.
  const fillRatio = i => groups[i].length / capacities[i];
  const targetRatio = Math.min(1, total / totalSeats);

  const unseated = [];

  leftover.forEach(player => {
    const roomy = allIdx.filter(i => groups[i].length < capacities[i]);
    if (!roomy.length) {
      unseated.push(player.player);
      return;
    }

    // Prefer a marker whose MarkerRoles suits this player, and only fall back to
    // the emptiest when none of them has room. Soft on purpose: a lopsided
    // roster overflows into the "wrong" markers rather than leaving people out.
    const emptiestOf = idx => {
      const suited = idx.filter(i => markerSuits_(player, roles[i]));
      return (suited.length ? suited : idx).sort((a, b) => fillRatio(a) - fillRatio(b))[0];
    };
    const balanced = roomy.filter(i =>
      !violatesAnyMax_(groups[i], activeRules, player) && fillRatio(i) < targetRatio);
    const anyValid = roomy.filter(i => !violatesAnyMax_(groups[i], activeRules, player));

    let targetIdx;
    if (balanced.length) {
      targetIdx = emptiestOf(balanced);
    } else if (anyValid.length) {
      targetIdx = emptiestOf(anyValid);
    } else {
      targetIdx = roomy
        .map(i => ({ i, violations: countViolations_(groups[i], activeRules, player) }))
        .sort((a, b) => a.violations - b.violations || fillRatio(a.i) - fillRatio(b.i))[0].i;
      warnings.push({
        internal: true,
        text: `${player.player} had to break a MaxPerGroup limit to get a slot`,
        fix: 'Your GroupRules are tighter than the roster allows, so relax a MaxPerGroup for this boss.',
      });
    }

    groups[targetIdx].push(player);
    assigned.add(player.player);
  });

  // Log only. More eligible people than seats is the ordinary shape of a
  // roster — seventeen ranged for fifteen slots means two sit out, every week,
  // and nobody reading the panel on raid night is going to act on it. The
  // shortfall that DOES matter (fallbacks pushed into main slots rather than
  // the backup row) is reported separately by noteFallbacksInMainSlots_.
  //
  // On a selection row it is not even a shortfall: one Bloodlust from five
  // shamans means four were simply not the pick.
  if (unseated.length && !isSelection) {
    Logger.log(`${label || 'this row'} sat out: ${unseated.join(', ')} ` +
      `(${totalSeats} slots, ${total} eligible)`);
  }
}

/* ------------------------------------------------------------------ *
 * Header-driven config reading
 * ------------------------------------------------------------------ */

function normalizeHeader_(header) {
  return String(header).trim().toLowerCase().replace(/[^a-z0-9]/g, '');
}

/**
 * Compares config values against roster values, ignoring case and a trailing
 * plural "s" — so "Healers" in AutoRoles still matches the "Healer" role, and
 * "Shamans" matches the Shaman class. No WoW class, spec or role name ends in
 * "s", so this cannot collapse two distinct values into one.
 */
function looseEquals_(a, b) {
  const x = String(a).trim().toLowerCase();
  const y = String(b).trim().toLowerCase();
  if (x === y) return true;

  const singular = text => (text.length > 1 && text.slice(-1) === 's') ? text.slice(0, -1) : text;
  return singular(x) === singular(y);
}

function buildHeaderIndexForRow_(sheet, row) {
  const headers = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
  const index = {};
  headers.forEach((header, i) => {
    const key = normalizeHeader_(header);
    if (key && index[key] === undefined) index[key] = i;
  });
  return index;
}

/**
 * Finds the header row by looking for one that carries every required column.
 * Scans the top of the sheet rather than assuming row 1, so config sheets can
 * be Google Sheets Tables, sit below a title, or start partway down.
 */
function locateHeaders_(sheet, requiredAliasGroups, sheetLabel) {
  const scanTo = Math.min(AUTO_ASSIGN_HEADER_SCAN_ROWS, sheet.getLastRow());

  for (let row = 1; row <= scanTo; row++) {
    const index = buildHeaderIndexForRow_(sheet, row);
    const complete = requiredAliasGroups.every(aliases =>
      aliases.some(alias => index[normalizeHeader_(alias)] !== undefined));
    if (complete) return { row: row, index: index };
  }

  const wanted = requiredAliasGroups.map(aliases => aliases[0]).join(', ');
  throw new Error(
    `Could not find a header row in "${sheetLabel}" containing: ${wanted}. ` +
    `Checked the first ${scanTo} row(s).`
  );
}

function pickColumn_(headerIndex, aliases, sheetName) {
  const i = pickOptionalColumn_(headerIndex, aliases);
  if (i === -1) {
    throw new Error(`Sheet "${sheetName}" needs a column headed one of: ${aliases.join(' / ')}`);
  }
  return i;
}

function pickOptionalColumn_(headerIndex, aliases) {
  for (const alias of aliases) {
    const i = headerIndex[normalizeHeader_(alias)];
    if (i !== undefined) return i;
  }
  return -1;
}

/** "Ranged, Warrior, Melee" -> ['ranged','warrior','melee']. Blank -> [] = everyone. */
function parseRoleList_(value) {
  return String(value).split(',').map(s => s.trim().toLowerCase()).filter(s => s !== '');
}

/**
 * How preferred this player is for an encounter, as the index of the first
 * AutoRoles entry they match — 0 is most preferred, -1 means not eligible.
 *
 * Entries may name a Role ("Ranged"), a Class ("Warrior") or a ClassSpec
 * ("Restoration Shaman"), so an ordered list like "Ranged, Warrior, Melee"
 * both filters and ranks: every ranged player is placed before any warrior,
 * and warriors before the remaining melee. Because groups fill in this order,
 * the lowest rows of each group — the backup slots — end up holding the least
 * preferred players automatically.
 */
function eligibilityRank_(player, entries) {
  if (!entries.length) return 0; // no list = everyone equally eligible

  const role = player.role.toLowerCase();
  const className = player.class.toLowerCase();
  const classspec = player.classspec.toLowerCase();

  // A bare Class means "a DPS of that class". Listing "Warrior" to get mobile
  // melee must not drag in the Protection Warrior who is tanking the boss.
  //
  // "Any Druid" turns that off, for abilities every spec of a class brings —
  // innervate comes from the tank and feral druids just as well as the balance
  // one, and there the role is irrelevant.
  const isSupport = role === 'tank' || role === 'healer';

  for (let i = 0; i < entries.length; i++) {
    let entry = entries[i];

    // "Flex Tank" asks what their off-spec would make them, so the druid who
    // can go bear counts without being a tank tonight. Only role and spec make
    // sense there — an off-spec never changes anyone's class.
    if (entry.indexOf('flex') === 0) {
      const flexWanted = entry.slice(4).trim();
      if (looseEquals_(flexWanted, player.flexrole)) return i;
      if (looseEquals_(flexWanted, player.flexclassspec)) return i;
      if (looseEquals_(flexWanted, player.flexSpec)) return i;
      continue;
    }

    const anySpec = entry.indexOf('any ') === 0;
    const wanted = anySpec ? entry.slice(4).trim() : entry;

    if (looseEquals_(wanted, role)) return i;
    if (looseEquals_(wanted, classspec)) return i;
    if (looseEquals_(wanted, className) && (anySpec || !isSupport)) return i;
  }
  return -1;
}

/**
 * A Raid column value, lowercased, with every way of writing "both" reduced to
 * an empty string. Blank, "*" and "all" all mean the row belongs to every
 * instance and should survive any scoped run.
 */
function normalizeRaidScope_(value) {
  const text = String(value === null || value === undefined ? '' : value).trim().toLowerCase();
  return (text === '*' || text === 'all') ? '' : text;
}

function parseBoolean_(value) {
  const text = String(value).trim().toLowerCase();
  return text === 'true' || text === 'yes' || text === '1';
}

/**
 * Every row of AutoAssignData, in sheet order.
 *
 * `scope` limits the run to one instance, matched against the Raid column
 * (blank there means every scope). Omit it for the everything button.
 */
function loadEncounterConfigs_(ss, warnings, scope) {
  const sheet = ss.getSheetByName(AUTO_ASSIGN_ENCOUNTERS_SHEET);
  if (!sheet) throw new Error(`Sheet "${AUTO_ASSIGN_ENCOUNTERS_SHEET}" not found`);

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const found = locateHeaders_(sheet, [
    ['Encounter'],
    ['CellAnchor', 'Anchor'],
    ['GroupAmount', 'GroupCount'],
    ['GroupSize', 'RowsPerGroup'],
  ], AUTO_ASSIGN_ENCOUNTERS_SHEET);

  const headers = found.index;
  const name = AUTO_ASSIGN_ENCOUNTERS_SHEET;
  const cEncounter = pickColumn_(headers, ['Encounter'], name);
  const cAnchor = pickColumn_(headers, ['CellAnchor', 'Anchor'], name);
  const cCount = pickColumn_(headers, ['GroupAmount', 'GroupCount'], name);
  const cSize = pickColumn_(headers, ['GroupSize', 'RowsPerGroup'], name);
  const cRoles = pickOptionalColumn_(headers, ['AutoRoles', 'EligibleRoles', 'Roles', 'Pool']);
  const cTotem = pickOptionalColumn_(headers, ['TotemParties', 'PartyCluster', 'Totems']);
  const cStep = pickOptionalColumn_(headers, ['ColumnStep', 'GroupSpacing', 'Step']);
  const cBackup = pickOptionalColumn_(headers, ['BackupRows', 'Backup', 'ReserveRows']);
  const cSheet = pickOptionalColumn_(headers, ['Sheet', 'TabName', 'Tab']);
  const cBalance = pickOptionalColumn_(headers, ['BalanceGroups', 'Balance', 'EvenGroups']);
  const cMarkers = pickOptionalColumn_(headers, ['MarkerNames', 'Markers', 'GroupNames']);
  const cMarkerRoles = pickOptionalColumn_(headers, ['MarkerRoles', 'RolePerMarker', 'PreferPerMarker']);
  const cInherit = pickOptionalColumn_(headers, ['InheritFrom', 'KeepFrom', 'Inherit']);
  const cSkipWhen = pickOptionalColumn_(headers, ['SkipWhen', 'SkipIf', 'OnlyIfNot']);
  const cPreferred = pickOptionalColumn_(headers, ['Preferred', 'Manual', 'PreSet']);
  const cLimitTo = pickOptionalColumn_(headers, ['LimitTo', 'LimitBy', 'MatchCount']);
  const cWriteBack = pickOptionalColumn_(headers, ['WriteBackTo', 'FillPreferred', 'SaveTo']);
  const cSelection = pickOptionalColumn_(headers, ['Selection', 'PickFew', 'IsSelection']);
  const cExcludeFrom = pickOptionalColumn_(headers, ['ExcludeFrom', 'NotFrom', 'AlreadyBusy']);
  const cEnsureEach = pickOptionalColumn_(headers, ['EnsureEach', 'OnePerGroup', 'SpreadEvenly']);
  const cCritical = pickOptionalColumn_(headers, ['Critical', 'MustHave', 'WipeRisk']);
  const cIgnoreAbove = pickOptionalColumn_(headers, ['IgnoreAbove', 'AllowReuse', 'SkipAboveScan']);
  const cRaid = pickOptionalColumn_(headers, ['Raid', 'Instance', 'Zone']);
  const cAllowDuplicates = pickOptionalColumn_(headers,
    ['AllowDuplicates', 'Duplicates', 'RepeatNames']);

  // Log only: a missing column is a state of the config, which is not the
  // business of whoever is reading the panel before a pull. Still recorded,
  // because "every tank is eligible for every boss" looks like an assignment
  // bug when you do not know the column is simply absent.
  if (cRoles === -1) {
    Logger.log(`NOTE: "${AUTO_ASSIGN_ENCOUNTERS_SHEET}" has no AutoRoles column, so every ` +
      `rostered player is eligible for every boss (tanks and melee included).`);
  }
  if (cTotem === -1) {
    Logger.log(`NOTE: no TotemParties column, so shaman-party clustering is off everywhere.`);
  }

  if (lastRow <= found.row) return [];

  const allRows = sheet.getRange(found.row + 1, 1, lastRow - found.row, sheet.getLastColumn()).getValues()
    .filter(r => String(r[cEncounter]).trim() !== '')
    .map(r => {
      const anchors = parseAnchors_(r[cAnchor]);
      const label = String(r[cEncounter]).trim();
      const groupCount = Number(r[cCount]) || 0;

      // Sheet accepts "*" for every tab another row names, optionally with
      // "-Illidan" style removals, so a raid-wide assignment does not carry a
      // hand-maintained list that goes stale the moment a boss is added.
      const tabEntries = cSheet === -1 ? [] : parseLabelList_(r[cSheet]);
      const wildcard = tabEntries.some(t => t === '*' || t.toLowerCase() === 'all');
      const excluded = tabEntries
        .filter(t => t.charAt(0) === '-')
        .map(t => t.slice(1).trim().toLowerCase());
      const tabs = tabEntries.filter(t =>
        t !== '*' && t.toLowerCase() !== 'all' && t.charAt(0) !== '-');

      return {
        wildcardSheets: wildcard,
        excludedSheets: excluded,
        // Encounter is the label rules match on; Sheet is the tab written to.
        // They differ when one boss needs several rows (ranged vs melee groups)
        // that want their own GroupRules, and Sheet can list several tabs when
        // the same assignment repeats across bosses (Bloodlust, PI, innervates).
        encounter: label,
        sheetNames: tabs.length ? tabs : [label],
        anchor: anchors.join(', '), // for messages
        anchors: anchors,
        groupCount: groupCount,
        groupSizes: parseSizes_(r[cSize], groupCount),
        columnStep: (cStep === -1 ? 0 : Number(r[cStep])) || 1,
        backupRows: (cBackup === -1 ? 0 : Number(r[cBackup])) || 0,
        eligibleRoles: cRoles === -1 ? [] : parseRoleList_(r[cRoles]),
        totemParties: cTotem === -1 ? false : parseBoolean_(r[cTotem]),
        balanceGroups: cBalance === -1 ? false : parseBoolean_(r[cBalance]),
        markerNames: cMarkers === -1 ? [] : parseLabelList_(r[cMarkers]),
        markerRoles: cMarkerRoles === -1 ? [] : parseSlotList_(r[cMarkerRoles]),
        inheritFrom: cInherit === -1 ? '' : String(r[cInherit]).trim(),
        skipWhen: cSkipWhen === -1 ? null : parseCondition_(r[cSkipWhen]),
        preferred: cPreferred === -1 ? [] : parseLabelList_(r[cPreferred]),
        limitTo: cLimitTo === -1 ? '' : String(r[cLimitTo]).trim(),
        writeBackTo: cWriteBack === -1 ? '' : String(r[cWriteBack]).trim(),
        selection: cSelection === -1 ? false : parseBoolean_(r[cSelection]),
        excludeFrom: cExcludeFrom === -1 ? [] : parseLabelList_(r[cExcludeFrom]),
        ensureEach: cEnsureEach === -1 ? [] : parseRoleList_(r[cEnsureEach]),
        // TRUE means "the raid cannot win this fight without this" (Council's
        // Mage for Spellsteal, Archimonde's decurse) — an empty pool for a
        // Critical row stays a red problem. For everything else, nobody
        // matching a valid AutoRoles term just means nobody has that class or
        // spec this week, which is normal and gets no mention at all.
        critical: cCritical === -1 ? false : parseBoolean_(r[cCritical]),
        // TRUE stops this row treating names in the few cells ABOVE its anchor
        // as already spoken for. That check exists for marker groups, where
        // the healer sitting above genuinely is part of the group — but a
        // utility table that merely happens to sit under an unrelated block
        // (Fear Ward beneath the healer table) would lose those people for no
        // reason. See loadPrePlacedNames_.
        ignoreAbove: cIgnoreAbove === -1 ? false : parseBoolean_(r[cIgnoreAbove]),
        // Which instance this row belongs to, for the per-raid buttons.
        //
        // Blank means "both". So do "*" and "all", because the Sheet column
        // one over already uses "*" for "every tab", and nobody should have to
        // remember that the same character means something different here.
        // Not accepting it dropped rows from every scoped run while leaving
        // them working on Assign Everything, which is a horrible way to fail.
        raid: normalizeRaidScope_(cRaid === -1 ? '' : r[cRaid]),
        // TRUE lets the same name appear more than once in Preferred. Right
        // for duty rows where a person can receive something twice, so two
        // priests both PI'ing the same mage, or two druids innervating the
        // same healer. Wrong for a marker group, where a repeated name would
        // put somebody in two positions at once, which is why it is off by
        // default rather than always allowed.
        allowDuplicates: cAllowDuplicates === -1 ? false : parseBoolean_(r[cAllowDuplicates]),
      };
    });

  // Scoping happens BEFORE the wildcard pass below, which is the whole trick:
  // "*" then expands to only the tabs this raid's own rows name, so a row like
  // Bloodlust follows the scope automatically instead of needing its own Raid
  // value. A blank Raid means "both", which is right for anything that spans
  // the night.
  const rows = !scope ? allRows : allRows.filter(row => !row.raid || row.raid === scope);

  if (scope) {
    // Names what was left out and why. A row dropped by scoping produces no
    // other trace at all, so a Raid column holding something unexpected looks
    // exactly like a feature that stopped working.
    const dropped = allRows.filter(row => row.raid && row.raid !== scope)
      .map(row => `${row.encounter} (Raid="${row.raid}")`);

    Logger.log(`Scoped to "${scope}": kept ${rows.length} of ${allRows.length} row(s)` +
      (dropped.length ? `. Left out: ${dropped.join(', ')}` : ''));
  }

  // "*" means every tab the other rows name — resolved after all rows are read,
  // and ignoring other wildcard rows so they cannot feed each other.
  const namedTabs = [];
  rows.forEach(row => {
    if (row.wildcardSheets) return;
    row.sheetNames.forEach(name => {
      if (namedTabs.indexOf(name) === -1) namedTabs.push(name);
    });
  });
  rows.forEach(row => {
    if (!row.wildcardSheets) return;

    // "*" already means every tab, so a plain name beside it adds nothing and
    // was almost certainly meant to be an exclusion. Log only — a Sheet column
    // written slightly wrong is config, not something to raise before a pull.
    if (row.sheetNames.length && row.sheetNames[0] !== row.encounter) {
      Logger.log(`"${row.encounter}" lists ${row.sheetNames.join(', ')} alongside "*", which ` +
        `does nothing — "*" is already every sheet. Prefix with "-" to exclude them instead.`);
    }

    row.sheetNames = namedTabs.filter(name =>
      row.excludedSheets.indexOf(name.toLowerCase()) === -1);
  });

  // One config per tab. A row listing twelve sheets becomes twelve encounters
  // sharing a label, so everything downstream stays a simple "assign this here".
  const expanded = [];
  rows.forEach(row => {
    row.sheetNames.forEach(name => {
      const copy = {};
      Object.keys(row).forEach(key => { copy[key] = row[key]; });
      copy.sheetName = name;
      copy.repeated = row.sheetNames.length > 1;
      delete copy.wildcardSheets;
      delete copy.excludedSheets;
      expanded.push(copy);
    });
  });
  return expanded;
}

/**
 * "Najentus!BC10" -> "BC10". Each row already names its own sheet via the
 * Encounter column, so a prefix is decorative. Stripping it also avoids A1
 * quoting problems with names like Kaz'rogal.
 */
function stripSheetPrefix_(anchor) {
  const text = String(anchor).trim();
  const bang = text.lastIndexOf('!');
  return bang === -1 ? text : text.slice(bang + 1).trim();
}

/** "BC10, BH10, BM10" -> ['BC10','BH10','BM10']. Single anchors still work. */
function parseAnchors_(value) {
  return String(value)
    .split(',')
    .map(part => stripSheetPrefix_(part))
    .filter(part => part !== '');
}

/** "Star, Circle, Square" -> ['Star','Circle','Square']. */
function parseLabelList_(value) {
  return String(value).split(',').map(s => s.trim()).filter(s => s !== '');
}

/**
 * "AE2 = Good DPS" -> { cell: 'AE2', negate: false, value: 'Good DPS' }.
 * Also accepts "!=". The cell is read on the encounter's own sheet, so a sheet
 * prefix is stripped and ignored like everywhere else.
 */
function parseCondition_(value) {
  const text = String(value).trim();
  if (!text) return null;

  const parts = text.match(/^(.+?)\s*(!=|=)\s*(.*)$/);
  if (!parts) return null;

  return {
    cell: stripSheetPrefix_(parts[1]),
    negate: parts[2] === '!=',
    value: parts[3].trim(),
  };
}

/** Evaluates a SkipWhen condition against the sheet. */
function evaluateCondition_(sheet, condition) {
  let actual;
  try {
    actual = String(sheet.getRange(condition.cell).getDisplayValue()).trim();
  } catch (err) {
    // Covers both a bad reference and a read that failed for Google's own
    // reasons, which look identical without the underlying message.
    return { error: `could not read the cell "${condition.cell}" (${(err && err.message) || err})` };
  }

  const same = actual.toLowerCase() === condition.value.toLowerCase();
  const matched = condition.negate ? !same : same;

  // Logged because both outcomes are otherwise silent: a condition that does
  // not match looks identical to a SkipWhen column the script never read.
  Logger.log(
    `SkipWhen on "${sheet.getName()}": ${condition.cell} is "${actual}", ` +
    `testing ${condition.negate ? '!=' : '='} "${condition.value}" -> ${matched ? 'SKIP' : 'run'}`);

  return { matched: matched, actual: actual };
}

/**
 * Like parseLabelList_ but keeps empty entries in place, so "Ranged, , Melee"
 * still lines up with group 1 and group 3.
 */
function parseSlotList_(value) {
  const text = String(value).trim();
  return text === '' ? [] : text.split(',').map(s => s.trim());
}

/**
 * Whether a marker's preferred role suits this player. Accepts a Role, Class or
 * ClassSpec, same as AutoRoles entries. A blank preference suits nobody, which
 * is what makes it fall through to normal balancing.
 */
function markerSuits_(player, markerRole) {
  if (!markerRole) return false;
  return eligibilityRank_(player, [String(markerRole).trim().toLowerCase()]) === 0;
}

/** "4, 5, 4, 4, 3" -> [4,5,4,4,3]. A single number applies to every group. */
function parseSizes_(value, groupCount) {
  const list = String(value)
    .split(',')
    .map(part => Number(part.trim()))
    .filter(n => !isNaN(n) && n > 0);

  if (!list.length) return [];
  if (list.length === 1) return new Array(groupCount).fill(list[0]);
  return list;
}

/**
 * Where each group's list lives. Two ways to say it:
 *   - One anchor per group, comma separated: "BC10, BH10, BM10". Handles any
 *     layout, including groups that are not evenly spaced.
 *   - A single anchor plus ColumnStep: "BC10" with step 5 gives BC, BH, BM.
 *     Step defaults to 1 (groups sitting side by side).
 */
function resolveTargetRanges_(sheet, encounter) {
  const anchors = encounter.anchors;
  const sizes = encounter.groupSizes;
  const ranges = [];

  if (sizes.length !== encounter.groupCount) {
    throw new Error(
      `GroupSize gives ${sizes.length} size(s) but GroupAmount is ${encounter.groupCount}`);
  }

  if (anchors.length > 1) {
    if (anchors.length !== encounter.groupCount) {
      throw new Error(
        `CellAnchor lists ${anchors.length} anchor(s) but GroupAmount is ${encounter.groupCount}`);
    }
    anchors.forEach((a, i) => {
      ranges.push(sheet.getRange(a).offset(0, 0, sizes[i], 1));
    });
    return ranges;
  }

  const first = sheet.getRange(anchors[0]);
  for (let i = 0; i < encounter.groupCount; i++) {
    ranges.push(first.offset(0, i * encounter.columnStep, sizes[i], 1));
  }
  return ranges;
}

/** All rules, tagged with their encounter (lowercased, "*" kept as-is). */
function loadGroupRules_(ss) {
  // Optional, and silent about it. Caps and extra minimums are a thing most
  // rosters never need, and no sheet behaves identically to an empty one — so
  // there is nothing to report, and saying so every run would be pure noise.
  // This used to throw, which meant deleting a sheet that LOOKS unused killed
  // the entire run, every boss, not just the rules.
  //
  // Note this is NOT the engine. EnsureEach builds the same shape of rule in
  // assignEncounter_ and is unaffected by this sheet existing or not.
  const sheet = ss.getSheetByName(AUTO_ASSIGN_GROUP_RULES_SHEET);
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const found = locateHeaders_(sheet, [
    ['Encounter'],
    ['MatchType'],
    ['MatchValue'],
    ['MinPerGroup', 'Min'],
    ['MaxPerGroup', 'Max'],
    ['Priority'],
  ], AUTO_ASSIGN_GROUP_RULES_SHEET);

  const headers = found.index;
  const name = AUTO_ASSIGN_GROUP_RULES_SHEET;
  const cEncounter = pickColumn_(headers, ['Encounter'], name);
  const cType = pickColumn_(headers, ['MatchType'], name);
  const cValue = pickColumn_(headers, ['MatchValue'], name);
  const cMin = pickColumn_(headers, ['MinPerGroup', 'Min'], name);
  const cMax = pickColumn_(headers, ['MaxPerGroup', 'Max'], name);
  const cPriority = pickColumn_(headers, ['Priority'], name);

  const blank = v => v === '' || v === null || v === undefined;
  if (lastRow <= found.row) return [];

  return sheet.getRange(found.row + 1, 1, lastRow - found.row, sheet.getLastColumn()).getValues()
    .filter(r => String(r[cEncounter]).trim() !== '')
    .map(r => {
      const enc = String(r[cEncounter]).trim().toLowerCase();
      return {
        encounter: (enc === 'all') ? '*' : enc,
        matchType: normalizeHeader_(r[cType]), // class | spec | role | classspec
        matchValue: String(r[cValue]).trim(),
        min: blank(r[cMin]) ? 0 : Number(r[cMin]),
        max: blank(r[cMax]) ? Infinity : Number(r[cMax]),
        priority: blank(r[cPriority]) ? Number.MAX_SAFE_INTEGER : Number(r[cPriority]),
      };
    });
}

/* ------------------------------------------------------------------ *
 * Roster, specs and parties
 * ------------------------------------------------------------------ */

/**
 * Returns { roles, terms } from Class/Spec Data.
 *
 * "terms" is every class, spec, spec+class and role the game has, which is what
 * lets a mis-typed AutoRoles entry be told apart from a real one that simply is
 * not raiding tonight — plenty of weeks have no Restoration Druid.
 */
function loadRoleMap_(ss) {
  const sheet = ss.getSheetByName(AUTO_ASSIGN_ROLE_MAP_SHEET);
  if (!sheet) throw new Error(`Sheet "${AUTO_ASSIGN_ROLE_MAP_SHEET}" not found`);

  const lastRow = sheet.getLastRow();
  const roles = {};
  const terms = {};
  if (lastRow < 2) return { roles: roles, terms: terms };

  const remember = value => {
    const key = String(value).trim().toLowerCase();
    if (key) terms[key] = true;
  };

  sheet.getRange(2, 1, lastRow - 1, 6).getValues().forEach(([cls, spec, , , , role]) => {
    if (!cls && !spec) return;
    roles[roleMapKey_(cls, spec)] = String(role).trim();

    remember(cls);
    remember(spec);
    remember(role);
    if (cls && spec) remember(`${String(spec).trim()} ${String(cls).trim()}`);
  });
  return { roles: roles, terms: terms };
}

/**
 * Whether an AutoRoles entry names something real, ignoring the "Any " and
 * "Flex " prefixes and a trailing plural.
 *
 * Both prefixes change which field is compared rather than naming anything
 * themselves — "Flex Tank" is the ordinary role Tank, asked of an off-spec — so
 * they have to come off before checking against Class/Spec Data.
 */
function isKnownTerm_(terms, entry) {
  let bare = entry;
  if (bare.indexOf('flex') === 0) bare = bare.slice(4).trim();
  else if (bare.indexOf('any ') === 0) bare = bare.slice(4).trim();

  if (terms[bare]) return true;

  const singular = (bare.length > 1 && bare.slice(-1) === 's') ? bare.slice(0, -1) : bare;
  return !!terms[singular];
}

function roleMapKey_(cls, spec) {
  return `${String(cls).trim().toLowerCase()}|${String(spec).trim().toLowerCase()}`;
}

/** Player -> raid party number, from Dynamic Lists columns ActiveRosterG1..G5. */
function loadPartyMembership_(ss) {
  const sheet = ss.getSheetByName(AUTO_ASSIGN_DYNAMIC_LISTS_SHEET);
  const byPlayer = new Map();
  if (!sheet) return byPlayer;

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return byPlayer;

  // Party data is optional unless TotemParties is on, so a miss is not fatal:
  // seedGroupsByTotemParties_ warns if it ends up with nothing to cluster on.
  let found;
  try {
    found = locateHeaders_(sheet, [['ActiveRosterG1']], AUTO_ASSIGN_DYNAMIC_LISTS_SHEET);
  } catch (err) {
    return byPlayer;
  }

  const headers = found.index;
  if (lastRow <= found.row) return byPlayer;

  const values = sheet
    .getRange(found.row + 1, 1, lastRow - found.row, sheet.getLastColumn())
    .getDisplayValues();

  for (let party = 1; party <= AUTO_ASSIGN_PARTY_COUNT; party++) {
    const col = pickOptionalColumn_(headers, [`ActiveRosterG${party}`]);
    if (col === -1) continue;
    values.forEach(row => {
      const name = String(row[col]).trim();
      if (name) byPlayer.set(name.toLowerCase(), party);
    });
  }
  return byPlayer;
}

/**
 * Everyone currently sitting in a *_RosterSwaps named range, lowercased.
 *
 * One range per split (e.g. Split1_RosterSwaps, Split2_RosterSwaps) — the same
 * "find every named range matching a suffix" trick assignHealersAndTanks uses
 * for *_tanks/*_healers, so a new split needs a new named range on the sheet
 * and nothing here has to change. These are people a separate script ticks
 * "Is Rostered" for so they show up in a manual dropdown for one fight; they
 * are deliberately excluded from every automatic assignment pool in
 * loadRoster_, and stay pickable only by hand.
 */
function loadRosterSwapNames_(ss) {
  const names = new Set();
  ss.getNamedRanges()
    .filter(nr => /_RosterSwaps$/i.test(nr.getName()))
    .forEach(nr => {
      nr.getRange().getDisplayValues().flat().forEach(value => {
        const name = String(value).trim();
        if (name) names.add(name.toLowerCase());
      });
    });
  return names;
}

/**
 * Rostered players with an effective role.
 *
 * Spec precedence: Raider Data (what they actually picked) over AutoAssign.
 * Role precedence: this week's TANKS_GLOBAL / HEALERS_GLOBAL picks, then the
 * spec-based role from Class/Spec Data.
 */
function loadRoster_(ss, roleMap, assignedTanks, assignedHealers, swapNames, excludedOut) {
  const sheet = ss.getSheetByName(AUTO_ASSIGN_RAIDER_DATA_SHEET);
  if (!sheet) throw new Error(`Sheet "${AUTO_ASSIGN_RAIDER_DATA_SHEET}" not found`);

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  // Raider Data is already filtered by the Is Rostered checkbox on the Roster
  // sheet, so everyone listed here is raiding. There is deliberately no second
  // roster to consult — two player lists can only ever drift apart.
  let rows = sheet.getRange(2, 1, lastRow - 1, 4).getDisplayValues()
    .filter(([player]) => String(player).trim() !== '');

  // A Roster Swap addition gets "Is Rostered" ticked too (by a different
  // script, purely to populate a manual dropdown for one fight), which would
  // otherwise make them a full night-long roster member here — eligible for
  // automatic assignment on every boss. swapNames comes straight from the
  // *_RosterSwaps named ranges (see loadRosterSwapNames_), so this excludes
  // exactly who is currently sitting in one of those lists, no matter how
  // Group Builder (25) or Static Lists compute anything else downstream. They
  // stay pickable by hand through whatever dropdown Roster Swaps feeds — this
  // only keeps the auto-assign logic itself from reaching for them.
  if (swapNames && swapNames.size) {
    const excluded = [];
    rows = rows.filter(([player]) => {
      const isSwap = swapNames.has(String(player).trim().toLowerCase());
      if (isSwap) excluded.push(String(player).trim());
      return !isSwap;
    });
    if (excluded.length) {
      Logger.log(`loadRoster_: excluded as Roster Swap addition(s): ${excluded.join(', ')}`);
      if (excludedOut) excluded.forEach(name => excludedOut.push(name));
    }
  }

  return rows.map(([player, cls, spec, flexSpec]) => {
      const name = String(player).trim();
      const key = name.toLowerCase();
      const className = String(cls).trim();
      const specName = String(spec).trim();

      let role = roleMap.roles[roleMapKey_(className, specName)] || '';
      if (assignedTanks.has(key)) role = 'Tank';
      else if (assignedHealers.has(key)) role = 'Healer';

      // What their off-spec would make them — a Feral druid whose flex is
      // Guardian counts as a flex tank, without being a tank tonight.
      const flexName = String(flexSpec).trim();
      const flexRole = flexName ? (roleMap.roles[roleMapKey_(className, flexName)] || '') : '';

      return {
        player: name,
        class: className,
        spec: specName,
        // Matches the specsJoined format in Class/Spec Data ("Balance Druid"),
        // so rules can target Restoration Shaman without hitting Restoration Druid.
        classspec: specName && className ? `${specName} ${className}` : '',
        flexSpec: flexName,
        flexclassspec: flexName && className ? `${flexName} ${className}` : '',
        flexrole: flexRole,
        role: role,
      };
    });
}

/**
 * Names currently sitting in a named range, lowercased.
 * Returns null when the named range does not exist, so callers can tell
 * "not set up" apart from "set up but nobody picked yet".
 */
function loadNamesFromNamedRange_(ss, rangeName) {
  const range = ss.getRangeByName(rangeName);
  if (!range) return null;

  const names = new Set();
  range.getDisplayValues().flat().forEach(value => {
    const name = String(value).trim();
    if (name) names.add(name.toLowerCase());
  });
  return names;
}

/**
 * Lists every sheet the tank/healer broadcast can reach, and how many cells it
 * has there.
 *
 * The ranges are named for convenience rather than by boss, so the sheet is the
 * only useful key — and a sheet missing from this list is one the broadcast
 * cannot write to at all, which is silent otherwise. A count smaller than the
 * number of picks means names fall off the bottom, equally silently.
 */
function logBroadcastCoverage_(ss, tankCount, healerCount) {
  const bySheet = {};

  const record = (nr, kind) => {
    const range = nr.getRange();
    const sheetName = range.getSheet().getName();
    if (!bySheet[sheetName]) bySheet[sheetName] = { tanks: 0, healers: 0 };
    bySheet[sheetName][kind] += range.getNumRows() * range.getNumColumns();
  };

  ss.getNamedRanges().forEach(nr => {
    if (/_tanks$/i.test(nr.getName())) record(nr, 'tanks');
    else if (/_healers$/i.test(nr.getName())) record(nr, 'healers');
  });

  // Only the sheets that cannot hold what is being sent. Listing all thirteen
  // every run buried the two that were actually short.
  Object.keys(bySheet).sort().forEach(sheetName => {
    const entry = bySheet[sheetName];
    const short = [];
    if (entry.tanks < tankCount) short.push(`tanks ${entry.tanks}/${tankCount}`);
    if (entry.healers < healerCount) short.push(`healers ${entry.healers}/${healerCount}`);

    if (short.length) {
      Logger.log(`${sheetName} cannot fit everyone: ${short.join(', ')}. ` +
        `Extend its named range, or ignore if that sheet needs fewer.`);
    }
  });
}

/**
 * Writes what was chosen into a named range on the setup page, so next run it
 * comes back through Preferred as a decision somebody can see and change.
 *
 * Only ever writes when that range is completely empty — the same rule as the
 * tank and healer auto-fill. Once a name is in there it is a human's choice and
 * gets left alone, which is what stops this quietly rewriting itself every week.
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

  const alreadySet = range.getDisplayValues().flat()
    .some(value => String(value).trim() !== '');
  if (alreadySet) return;

  // Blank on screen but a formula underneath — writing would destroy it. Same
  // protection the boss sheets get; a setup page is no less worth keeping.
  const formula = range.getFormulas().flat().filter(f => f !== '')[0];
  if (formula) {
    warnings.push({
      internal: true,
      text: `"${rangeName}" holds a formula, so nothing was written back to it`,
      fix: 'Point WriteBackTo at plain cells people can type in, or drop the column.',
    });
    return;
  }

  const rows = range.getNumRows();
  const cols = range.getNumColumns();

  // Everyone eligible, in AutoRoles order — up to whatever the range holds.
  // Not just those who fitted on one boss: this is the list of people who can
  // do the job, and a four-cell range is asking for four candidates.
  const names = candidates.slice(0, rows * cols).map(player => player.player);

  Logger.log(`WriteBackTo ${rangeName}: ${rows * cols} cell(s), ` +
    `${candidates.length} candidate(s) [${candidates.map(p => p.player).join(', ')}] ` +
    `-> writing ${names.join(', ') || '(none)'}`);

  if (!names.length) return;
  const values = [];
  let index = 0;
  for (let r = 0; r < rows; r++) {
    const row = [];
    for (let c = 0; c < cols; c++) row.push(names[index++] || '');
    values.push(row);
  }
  try {
    range.setValues(values);
  } catch (err) {
    // Most likely a dropdown on those cells that does not list one of them.
    // Same shape as the boss-sheet refusal: name who was refused and where the
    // list lives, rather than sending anybody into the config.
    warnings.push({
      text: `a dropdown on "${rangeName}" would not accept ${names.join(', ')}`,
      fix: `${names.length === 1 ? 'That name is' : 'Those names are'} missing from the list ` +
        `of options those cells allow. Click one, check Data > Data validation, and add ` +
        `${names.length === 1 ? 'the name' : 'them'} to its list.`,
    });
    return;
  }

  filled.push(`${rangeName}: ${names.join(', ')}`);
}

/**
 * Takes people who are not raiding out of a picks range, then refills only the
 * slots they left behind.
 *
 * Last week's picks survive in the sheet, so without this the broadcast writes
 * a tank who is not in the building onto every boss sheet, silently. Anyone
 * missing from tonight's roster is cleared out first.
 *
 * Any empty slot is then filled from the priority list, whether it was emptied
 * here or cleared by hand before the run. Clearing a cell to make the script
 * choose again is the obvious thing to try, so it had better work.
 *
 * There is deliberately no way to mark a slot "leave this empty". A cell holds
 * a name or it does not, and nobody has to learn a placeholder to make the
 * thing behave. If a slot should stay empty, shrink the named range.
 *
 * Position carries meaning here. The first cell is the Main Tank, and
 * assignHealersAndTanks copies it into Council's one cell MT range as well as
 * the top of its two cell range. It is filled like any other slot, but who
 * landed there is reported back so the caller can say so plainly. A raid
 * leader who reads "Phil was picked as your main tank" fixes it in seconds if
 * it is wrong, where a blank slot goes unnoticed until the pull.
 *
 * Returns { removed, added, firstVacated, firstFilled }.
 */
function reconcilePicks_(ss, rangeName, rosteredKeys, candidates, priority, protectFirst) {
  const result = { removed: [], added: [], firstVacated: '', firstFilled: '' };

  const range = ss.getRangeByName(rangeName);
  if (!range) return result;

  const values = range.getDisplayValues();
  const rows = values.length;
  const cols = rows ? values[0].length : 0;

  // Flat index, reading left to right then down, so "first cell" means the
  // same thing here as it does to assignHealersAndTanks.
  const fillable = [];
  const present = new Set();
  let flat = -1;

  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) {
      flat++;
      const name = String(values[r][c]).trim();

      // Only two states carry meaning: a name, or empty. No placeholder to
      // learn, nothing to type to opt out. Anything that is not a rostered
      // name gets cleared and refilled, including a stray dash left behind by
      // hand, which is the tidier outcome anyway.
      if (name && rosteredKeys.has(name.toLowerCase())) {
        present.add(name.toLowerCase());
        continue;
      }

      if (name) {
        values[r][c] = '';
        result.removed.push(name);
        if (flat === 0) result.firstVacated = name;
      }

      // Empty now, whether it was cleared by hand before the run or emptied a
      // line ago. Both mean the same thing: this slot wants somebody in it.
      // That includes the first cell, the Main Tank. It gets filled like any
      // other and then named in the result, on the grounds that a raid leader
      // reading "Phil is your main tank" will correct it in seconds if it is
      // wrong, which beats a blank slot nobody notices until the pull.
      fillable.push({ r: r, c: c, isFirst: flat === 0 });
    }
  }

  if (!fillable.length && !result.removed.length) return result;

  // Only people who could actually do the job, best first, and never somebody
  // already sitting in another cell of this same range.
  const wanted = priority.map(p => String(p).trim().toLowerCase());
  const replacements = candidates
    .map(player => ({ player: player, rank: eligibilityRank_(player, wanted) }))
    .filter(entry => entry.rank !== -1 && !present.has(entry.player.player.toLowerCase()))
    .sort((a, b) => a.rank - b.rank);

  fillable.forEach(cell => {
    const next = replacements.shift();
    // Nobody left who can do the job, so the slot honestly stays empty. Not
    // reported: an empty tank slot is obvious to anybody looking at Quick
    // Assigns, and the run has bigger things to say.
    if (!next) return;
    values[cell.r][cell.c] = next.player.player;
    result.added.push(next.player.player);
    if (cell.isFirst && protectFirst) result.firstFilled = next.player.player;
    present.add(next.player.player.toLowerCase());
  });

  if (!result.removed.length && !result.added.length) return result;

  range.setValues(values);

  Logger.log(`${rangeName}: ` +
    (result.removed.length ? `removed ${result.removed.join(', ')} (not raiding). ` : '') +
    (result.added.length ? `filled ${result.added.join(', ')}.` : 'nothing available to fill with.'));

  return result;
}

/** The ordered list from a priority named range, or the built-in default. */
function loadPriority_(ss, rangeName, fallback) {
  const range = ss.getRangeByName(rangeName);
  if (!range) return fallback;

  const entries = range.getDisplayValues()
    .flat()
    .map(value => String(value).trim())
    .filter(value => value !== '');

  // One cell holding a list, or one name per cell — accept either.
  const parsed = entries.length === 1 ? parseLabelList_(entries[0]) : entries;
  return parsed.length ? parsed : fallback;
}

/**
 * Fills an empty TANKS_GLOBAL / HEALERS_GLOBAL from the roster.
 *
 * Only ever runs when the range is completely empty — a single name means
 * somebody has decided, and that decision is never overwritten. Picks in the
 * order the priority list gives, and takes as many as the range holds.
 *
 * Returns the names written, or [] if it did nothing.
 */
function autoFillPicks_(ss, rangeName, players, priority) {
  const range = ss.getRangeByName(rangeName);
  if (!range) {
    Logger.log(`Auto-fill ${rangeName}: that named range does not exist`);
    return [];
  }

  const capacity = range.getNumRows() * range.getNumColumns();
  Logger.log(`Auto-fill ${rangeName}: ${capacity} cell(s), priority [${priority.join(' > ')}], ` +
    `${players.length} rostered player(s)`);
  if (!capacity) return [];

  const wanted = priority.map(p => String(p).trim().toLowerCase());
  const chosen = players
    .map(player => ({ player: player, rank: eligibilityRank_(player, wanted) }))
    .filter(entry => entry.rank !== -1)
    .sort((a, b) => a.rank - b.rank)
    .slice(0, capacity);

  if (!chosen.length) {
    Logger.log(`Auto-fill ${rangeName}: nobody matched. Roster reads as — ` +
      players.map(p => `${p.player}=${p.classspec || '?'}/${p.role || '?'}`).join(', '));
    return [];
  }

  Logger.log(`Auto-fill ${rangeName}: writing ${chosen.map(entry =>
    `${entry.player.player} (${entry.player.classspec || '?'}, priority ${entry.rank + 1})`).join(', ')}`);

  const names = chosen.map(entry => entry.player.player);

  const rows = range.getNumRows();
  const cols = range.getNumColumns();
  const values = [];
  let index = 0;
  for (let r = 0; r < rows; r++) {
    const row = [];
    for (let c = 0; c < cols; c++) row.push(names[index++] || '');
    values.push(row);
  }
  range.setValues(values);

  return names;
}

/**
 * Empty tank/healer picks almost always mean Assign Healers/Tanks hasn't been
 * run yet — in which case the healer rows are blank too and every group would
 * rest on stale input. Describe exactly what's wrong.
 */
function describeMissingPicks_(tanks, healers) {
  const problems = [];
  const check = (names, rangeName, label) => {
    if (names === null) problems.push(`${rangeName} does not exist`);
    else if (!names.size) problems.push(`no ${label} are selected (${rangeName} is empty)`);
  };
  check(tanks, AUTO_ASSIGN_TANKS_RANGE, 'tanks');
  check(healers, AUTO_ASSIGN_HEALERS_RANGE, 'healers');
  return problems;
}

/**
 * Everything already written above each group's anchor — in practice the
 * formula-driven healer row. Non-player text (headers etc.) is harmless: it
 * simply never matches a roster name.
 */
/**
 * The few cells directly above a group's anchor — its healer rows — nearest
 * first. Bounded by AUTO_ASSIGN_HEALER_SCAN_ROWS so a table stacked higher up
 * the same columns is not mistaken for part of this group.
 */
function readAboveAnchor_(sheet, range, ownedCells, snapshot) {
  const anchorRow = range.getRow();
  if (anchorRow <= 1) return [];

  const height = Math.min(AUTO_ASSIGN_HEALER_SCAN_ROWS, anchorRow - 1);
  const startRow = anchorRow - height;
  const column = range.getColumn();

  const direct = snapshot ? null : timed_('scanning',
    () => sheet.getRange(startRow, column, height, 1).getDisplayValues());
  const directFormulas = snapshot ? null : timed_('scanning',
    () => sheet.getRange(startRow, column, height, 1).getFormulas());

  const found = [];
  for (let i = 0; i < height; i++) {
    const row = startRow + i;
    const raw = snapshot
      ? snapshotCell_(snapshot, 'values', row, column)
      : direct[i][0];
    const formula = snapshot
      ? snapshotCell_(snapshot, 'formulas', row, column)
      : directFormulas[i][0];

    const text = String(raw === null ? '' : raw).trim();
    if (!text) continue;

    // Skip cells this script fills from another row. Stacked blocks sit close
    // together — PI targets are three rows above innervate targets — and
    // reading our own output back as "already placed" empties the pool.
    if (ownedCells && ownedCells[`${column},${row}`]) continue;

    // Whether the name arrived by formula decides whose business it is. A
    // marker's healer row is "=AV15" pulling from the healer table, so it
    // belongs to the block below it. A name that was typed or picked from a
    // dropdown belongs to whatever table it is sitting in, which may have
    // nothing to do with the block underneath. Callers that care about
    // ownership filter on this; findMarkerHealers_ deliberately does not,
    // since a Shaman anchors their marker's party either way.
    found.push({
      text: text,
      row: row,
      column: column,
      formula: String(formula === null || formula === undefined ? '' : formula).trim(),
      hasFormula: !!(formula && String(formula).trim()),
    });
  }
  return found.reverse(); // nearest the anchor first
}

/**
 * Lowercased name -> the cell it was found in, so a rejection can point at the
 * exact cell rather than a range to go searching through.
 */
function loadPrePlacedNames_(sheet, targetRanges, ownedCells, snapshot) {
  const byName = new Map();
  targetRanges.forEach(range => {
    readAboveAnchor_(sheet, range, ownedCells, snapshot).forEach(cell => {
      // Only a formula above this block means "this person is already in the
      // group I am about to fill" — that is the healer row pulling from the
      // healer table. A typed or dropdown name above belongs to a different
      // table that merely happens to sit here (Fear Ward under the Healer
      // Assignments box), and striking those people off would quietly lose
      // them for no reason. IgnoreAbove remains as a manual override.
      if (!cell.hasFormula) return;

      // ...unless that formula reads FROM this block, in which case it is a
      // summary of what we are about to write, not a claim on anybody. A
      // pre-pull Fear Ward cell of "=IF(BO19=\"\",BO18,BO19)" mirrors the
      // rotation directly beneath it, so honouring it strikes off whoever was
      // assigned last run — and the following run strikes off the other one,
      // oscillating forever and never filling more than one slot. A real
      // healer row ("=AV18") points at the healer table instead, well outside
      // the block, and is protected exactly as before.
      if (formulaReadsFromRanges_(cell.formula, targetRanges)) return;

      const key = cell.text.toLowerCase();
      if (!byName.has(key)) {
        byName.set(key, sheet.getRange(cell.row, cell.column).getA1Notation());
      }
    });
  });
  return byName;
}

/**
 * The healer names above each marker, nearest first. A marker can carry more
 * than one (Illidan's Triangle holds both tank healers) or none at all, so this
 * returns a list and lets the caller decide which one matters.
 */
function findMarkerHealers_(sheet, targetRanges, ownedCells, snapshot) {
  return targetRanges.map(range =>
    readAboveAnchor_(sheet, range, ownedCells, snapshot).map(cell => cell.text));
}

/* ------------------------------------------------------------------ *
 * Matching helpers
 * ------------------------------------------------------------------ */

const AUTO_ASSIGN_MATCH_FIELDS = {
  class: 'class',
  spec: 'spec',
  role: 'role',
  classspec: 'classspec',
  specclass: 'classspec',
  specsjoined: 'classspec',
  player: 'player',
  name: 'player',
};

function matchesRule_(rule, player) {
  // The synthetic rule behind EnsureEach. Uses the AutoRoles vocabulary, so
  // "Mage, Druid", "Restoration Shaman" and "Ranged" all work without anyone
  // having to say which kind of thing they named.
  if (rule.matchType === 'ensure') {
    return eligibilityRank_(player, rule.entries) !== -1;
  }

  const field = AUTO_ASSIGN_MATCH_FIELDS[rule.matchType];
  if (!field) return false;

  const value = player[field];
  if (value === undefined) return false;
  if (rule.matchValue === '' || rule.matchValue === '*') return true;

  // A comma-separated MatchValue means "any of these". "Mage, Druid" with
  // MinPerGroup 1 asks for a decurser in every group, which is one requirement —
  // two separate rules would each chase their own minimum and double up.
  const wanted = rule.matchValue.split(',').map(s => s.trim()).filter(s => s !== '');

  // Player names match exactly. The plural tolerance is meant for role and
  // class words, and would let a rule for "Hand" catch a raider called "Hands".
  if (field === 'player') {
    const actual = String(value).trim().toLowerCase();
    return wanted.some(name => actual === name.toLowerCase());
  }
  return wanted.some(entry => looseEquals_(value, entry));
}

function countMatching_(group, rule) {
  return group.filter(p => matchesRule_(rule, p)).length;
}

function isGroupAtMax_(group, rule) {
  return rule.max !== Infinity && countMatching_(group, rule) >= rule.max;
}

function countOtherMatches_(player, rules) {
  return rules.filter(r => matchesRule_(r, player)).length;
}

function violatesAnyMax_(group, rules, player) {
  return rules.some(rule => matchesRule_(rule, player) && isGroupAtMax_(group, rule));
}

function countViolations_(group, rules, player) {
  return rules.filter(rule => matchesRule_(rule, player) && isGroupAtMax_(group, rule)).length;
}

/* ------------------------------------------------------------------ *
 * Writing
 * ------------------------------------------------------------------ */

/**
 * Names the dropdowns in these cells would reject.
 *
 * Only called from diagnoseWriteFailure_, after a write has already been
 * refused — reading validations up front cost more than the entire rest of the
 * run. Call it only with the target cells cleared: the dropdown lists are
 * derived from who is unassigned, so a name still sitting in a group is missing
 * from its own dropdown and would look rejected when it is not.
 */
function findValidationRejects_(targetRanges, groups) {
  const rejected = [];

  targetRanges.forEach((range, i) => {
    const names = groups[i].map(p => p.player).slice(0, range.getNumRows());
    if (!names.length) return;

    const rules = range.getDataValidations();
    names.forEach((name, row) => {
      const rule = rules[row] && rules[row][0];
      if (!rule || rule.getAllowInvalid()) return; // no rule, or warn-only: cannot throw

      const allowed = allowedValuesFor_(rule);
      if (allowed && allowed.indexOf(name.trim().toLowerCase()) === -1) {
        rejected.push(name);
      }
    });
  });

  return rejected;
}

/** Lowercased accepted values for a list/range validation rule, or null if not a list. */
function allowedValuesFor_(rule) {
  const criteria = SpreadsheetApp.DataValidationCriteria;

  // A dropdown whose source range has been deleted throws when read. That is a
  // broken dropdown, not a reason to abandon the assignment — treat it as
  // unrestricted and let the write succeed or fail on its own merits.
  try {
    const type = rule.getCriteriaType();
    const values = rule.getCriteriaValues();

    if (type === criteria.VALUE_IN_LIST) {
      return values[0].map(v => String(v).trim().toLowerCase());
    }
    if (type === criteria.VALUE_IN_RANGE) {
      return values[0].getDisplayValues().flat()
        .map(v => String(v).trim().toLowerCase())
        .filter(v => v !== '');
    }
  } catch (err) {
    Logger.log(`Could not read a dropdown's allowed values: ${err.message}`);
  }
  return null;
}

/**
 * First target cell holding a formula, as { cell, onAnchorRow }, or null.
 *
 * Which row it is changes the advice completely: a formula on the anchor row
 * usually means the anchor is one row too high and is pointing at a healer;
 * further down it means a stray formula sits inside an otherwise correct block.
 */
function findFormulaCell_(targetRanges, snapshot) {
  for (const range of targetRanges) {
    const column = range.getColumn();
    const firstRow = range.getRow();
    const height = range.getNumRows();

    const direct = snapshot ? null : timed_('guarding', () => range.getFormulas());

    for (let r = 0; r < height; r++) {
      const formula = snapshot
        ? snapshotCell_(snapshot, 'formulas', firstRow + r, column)
        : direct[r][0];

      if (formula) {
        return {
          cell: range.offset(r, 0, 1, 1).getA1Notation(),
          onAnchorRow: r === 0,
        };
      }
    }
  }
  return null;
}

/**
 * Clears each group's block, then writes it. Returns { rejected, overflow }.
 *
 * Two attempts, because the fast one is wrong in a specific case. A dropdown
 * whose list is built from "players not already assigned" cannot accept
 * somebody who is still sitting in the very cell being rewritten — and
 * clearing that cell does not help unless something forces the list to
 * recalculate before the write lands. The first attempt skips that
 * recalculation because it costs a full sync per block and is almost never
 * needed; the second one pays it, and only for blocks that were actually
 * refused.
 *
 * If the second attempt is refused too, the name really is not allowed, and
 * the previous contents are put back so a rejection leaves the sheet exactly
 * as it was rather than wiped.
 */
function writeGroups_(targetRanges, groups, snapshot, label) {
  // The rollback copy comes from the snapshot when we have one. Reading it per
  // range made this the first read barrier after the pending writes, so it was
  // absorbing all of their cost — 92s of it.
  const previous = snapshot
    ? targetRanges.map(range => {
        const column = range.getColumn();
        const firstRow = range.getRow();
        const copy = [];
        for (let i = 0; i < range.getNumRows(); i++) {
          const value = snapshotCell_(snapshot, 'values', firstRow + i, column);
          copy.push([value === null ? '' : value]);
        }
        return copy;
      })
    : timed_('reads', () => targetRanges.map(range => range.getValues()));

  const overflow = [];

  const plans = targetRanges.map((range, i) => {
    const names = groups[i].map(p => p.player);
    const capacity = range.getNumRows();

    if (names.length > capacity) {
      overflow.push({
        group: i + 1,
        needed: names.length,
        capacity: capacity,
        dropped: names.slice(capacity),
      });
    }
    return names.slice(0, capacity).map(name => [name]);
  });

  Logger.log(`writeGroups_ ${label || '(unlabeled)'}: ` +
    targetRanges.map((range, i) => `${range.getA1Notation()}=[${plans[i].map(row => row[0]).join(', ') || '(empty)'}]`).join(', '));

  const clearBlocks = () => timed_('clearing', () =>
    targetRanges.forEach(range => range.clearContent()));
  const writeBlocks = () => timed_('writing', () => targetRanges.forEach((range, i) => {
    if (plans[i].length) range.offset(0, 0, plans[i].length, 1).setValues(plans[i]);
  }));

  // FAST PATH: clear then write with nothing in between. Reading the dropdowns
  // first cost more than everything else combined, and it was only ever
  // predicting what the write would do — so let the write answer that.
  //
  // flush() right after writing is deliberate and costs real time: without it,
  // Apps Script can defer a rejected write instead of throwing here, and the
  // error surfaces later on some unrelated read, blamed on the wrong row
  // entirely (that is exactly what happened with the Shahraz/Souls mixup).
  try {
    clearBlocks();
    writeBlocks();
    timed_('flushing', () => SpreadsheetApp.flush());
    return { rejected: [], overflow: overflow };
  } catch (err) {
    Logger.log(`writeGroups_ ${label || '(unlabeled)'}: write failed — ${err.message}. ` +
      `Retrying with a recalculation between clearing and writing.`);
  }

  // SLOW PATH, only after a refusal. Nothing forced a recalculation between
  // the clear and the write above, so a dropdown built from "who is not
  // already assigned" still had the people we were in the middle of clearing
  // marked as taken — and refused to accept them back into their own cells.
  // Flushing after the clear lets those lists catch up before we write, which
  // is what the original clear/flush/check ordering did before it was taken
  // out for speed. Paid only by the rare block that actually hits this.
  try {
    clearBlocks();
    timed_('flushing', () => SpreadsheetApp.flush());
    writeBlocks();
    timed_('flushing', () => SpreadsheetApp.flush());
    Logger.log(`writeGroups_ ${label || '(unlabeled)'}: retry succeeded after recalculation.`);
    return { rejected: [], overflow: overflow };
  } catch (err) {
    Logger.log(`writeGroups_ ${label || '(unlabeled)'}: retry also refused — ${err.message}`);

    // Built only here: one entry per cell we were about to fill, so the cell
    // Google names in its error can be turned back into the name we meant to
    // put there. Doing this up front would cost hundreds of lookups a run to
    // answer a question that almost never gets asked.
    const intended = {};
    targetRanges.forEach((range, i) => {
      const letter = columnLetter_(range.getColumn());
      const firstRow = range.getRow();
      plans[i].forEach((row, r) => { intended[letter + (firstRow + r)] = row[0]; });
    });

    return {
      rejected: diagnoseWriteFailure_(targetRanges, groups, previous, err.message, intended),
      overflow: [],
    };
  }
}

/**
 * Works out which names a dropdown refused, then puts the block back as it was.
 *
 * Only runs after a write has actually failed, so the expensive validation read
 * happens once in a blue moon rather than on every assignment. The cells are
 * cleared first because the dropdown lists are built from who is unassigned —
 * reading them against the restored values would blame everybody.
 */
function diagnoseWriteFailure_(targetRanges, groups, previous, originalMessage, intended) {
  let rejected = [];

  try {
    targetRanges.forEach(range => range.clearContent());
    SpreadsheetApp.flush();
    rejected = timed_('validating', () => findValidationRejects_(targetRanges, groups));
  } catch (err) {
    Logger.log(`Could not work out which names were refused: ${err.message}`);
  }

  try {
    targetRanges.forEach((range, i) => range.setValues(previous[i]));
    SpreadsheetApp.flush();
  } catch (err) {
    Logger.log(`Could not restore the previous assignment: ${err.message}`);
  }

  if (rejected.length) return rejected;

  // Reading the allowed values does not always work — a rule built on a custom
  // formula has no list to enumerate, so nothing comes back rejected even
  // though the write plainly failed. Google's own message still names the
  // cell, and we know what was going into it, so answer with that rather than
  // handing the raw error to somebody who cannot act on it. The raw text stays
  // in the log for whoever maintains the sheet.
  Logger.log(`diagnoseWriteFailure_: no dropdown list could be read. Raw error: ${originalMessage}`);

  const match = /cell ([A-Z]+[0-9]+)/i.exec(String(originalMessage));
  const cell = match ? match[1].toUpperCase() : '';

  if (cell && intended && intended[cell]) return [`${intended[cell]} (into ${cell})`];
  if (cell) return [`the name going into ${cell}`];
  return ['a name on this sheet'];
}

/**
 * True when a formula refers to any cell inside the given ranges.
 *
 * Used to tell a cell that DESCRIBES a block from one that FEEDS it. Both sit
 * above the block and both hold formulas, but only the second is a claim on a
 * player — the first is just a readout of what this run is about to write.
 *
 * References qualified by a sheet name are dropped first: "'Static Lists'!BO19"
 * is a different sheet's BO19 and says nothing about this block. Deliberately
 * simple string matching — it only has to recognise the everyday "=AV18" and
 * "=IF(BO19=..,BO18,BO19)" shapes, and erring towards "not from this block"
 * just leaves today's protective behaviour in place.
 */
function formulaReadsFromRanges_(formula, ranges) {
  if (!formula) return false;

  const local = String(formula)
    .replace(/('[^']*'|[A-Za-z_][A-Za-z0-9_.]*)!\$?[A-Z]{1,3}\$?[0-9]+(:\$?[A-Z]{1,3}\$?[0-9]+)?/g, ' ');

  const refs = local.match(/\$?[A-Z]{1,3}\$?[0-9]+/g);
  if (!refs) return false;

  return refs.some(ref => {
    const parts = /^\$?([A-Z]{1,3})\$?([0-9]+)$/.exec(ref);
    if (!parts) return false;

    const column = letterToColumn_(parts[1]);
    const row = Number(parts[2]);

    return ranges.some(range => {
      const first = range.getRow();
      return column === range.getColumn() &&
        row >= first && row <= first + range.getNumRows() - 1;
    });
  });
}

/** "A" -> 1, "AA" -> 27. The inverse of columnLetter_. */
function letterToColumn_(letters) {
  let column = 0;
  for (let i = 0; i < letters.length; i++) {
    column = column * 26 + (letters.charCodeAt(i) - 64);
  }
  return column;
}

/** 1 -> "A", 27 -> "AA". Pure arithmetic, so it costs no spreadsheet calls. */
function columnLetter_(column) {
  let letter = '';
  let n = column;
  while (n > 0) {
    const remainder = (n - 1) % 26;
    letter = String.fromCharCode(65 + remainder) + letter;
    n = Math.floor((n - 1) / 26);
  }
  return letter;
}