/**
 * Broadcasts the raid's tank and healer picks out to every boss sheet.
 *
 * Reads TANKS_GLOBAL and HEALERS_GLOBAL, then fills every named range whose
 * name ends in _tanks or _healers. Each target starts from the top of the list
 * independently, so a sheet can carry several ranges — Council's MT appears in
 * a one-cell range and again at the top of a two-cell one, giving MT, MT, OT.
 *
 * Called by autoAssignGroups() on every run, so it looks unreferenced in the
 * onOpen menu but is not. Do not delete.
 */
function assignHealersAndTanks() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const namedRanges = ss.getNamedRanges();
  const healerGlobals = ss.getRangeByName("HEALERS_GLOBAL");
  const tankGlobals = ss.getRangeByName("TANKS_GLOBAL");
  if (!healerGlobals && !tankGlobals) return;

  const healerNames = healerGlobals ? healerGlobals.getValues().flat().filter(n => n.toString().trim() !== "") : [];
  const tankNames = tankGlobals ? tankGlobals.getValues().flat().filter(n => n.toString().trim() !== "") : [];

  const healerTargets = namedRanges.filter(nr => /_healers$/i.test(nr.getName()));
  const tankTargets = namedRanges.filter(nr => /_tanks$/i.test(nr.getName()));

  const process = (targets, names) => {
    targets.forEach(nr => {
      const range = nr.getRange();
      const rows = range.getNumRows();
      const cols = range.getNumColumns();
      const values = new Array(rows);
      let index = 0;
      for (let r = 0; r < rows; r++) {
        const rowVals = new Array(cols);
        for (let c = 0; c < cols; c++) rowVals[c] = names[index++] || "";
        values[r] = rowVals;
      }
      range.setValues(values);
    });
  };

  if (healerTargets.length && healerNames.length) process(healerTargets, healerNames);
  if (tankTargets.length && tankNames.length) process(tankTargets, tankNames);

  if (healerTargets.length && healerNames.length)
    ss.toast(`Main Spec healers assigned to boss pages.`, "🧹 Auto Assigns Helper", 10);
  if (tankTargets.length && tankNames.length)
    ss.toast(`Tank Roles assigned to boss pages.`, "🧹 Auto Assigns Helper", 10);
}