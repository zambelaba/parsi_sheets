/**
 * Copy buttons for the T6 WeakAura export strings.
 *
 * Each button carries one already-built string, taken from a named range whose
 * cell holds a TEXTJOIN formula. Nothing is read when a button is clicked: the
 * values are baked into the page when the dialog opens, and the click only
 * copies to the clipboard. That means a string is as fresh as the moment you
 * pressed the button on the sheet, so reopen the dialog after changing an
 * assignment.
 *
 * Named ranges rather than cell addresses, matching showExportHub and the rest
 * of this spreadsheet: the tables can move without anybody editing code.
 *
 * Names here are deliberately distinct from showExportHub / ExportHubFojji.
 * Every .gs file in this project shares one namespace, so two functions with
 * the same name silently override each other with no warning at all.
 */

// Named range -> the key the template reads, and the button label.
// Add a row here and a button in T6ExportHub.html to extend this.
const T6_EXPORT_SOURCES = [
  { range: 'BLOODBOIL_SOAK_TEAMS', key: 'bloodboilSoak', label: 'Bloodboil Soak Teams' },
  { range: 'SOULS_INTERRUPTS', key: 'soulsInterrupts', label: 'Souls Interrupts' },
  { range: 'COUNCIL_INTERRUPTS', key: 'councilInterrupts', label: 'Council Interrupts' },
];

// Shown in the sidebar header. Kept short: the sidebar is narrow and a long
// title is simply truncated.
const T6_EXPORT_DIALOG_TITLE = 'T6 Export Strings';

/** Reads every export string. Shared by the button and the post-run open. */
function buildT6ExportData_(ss) {
  // One read of the named range list, then a lookup per source. Cheaper than
  // getRangeByName per entry, and it lets a missing range be reported rather
  // than throwing and leaving the user with a raw stack trace.
  const byName = {};
  ss.getNamedRanges().forEach(nr => { byName[nr.getName()] = nr; });

  const data = { missing: [], filled: 0 };

  T6_EXPORT_SOURCES.forEach(source => {
    const named = byName[source.range];

    if (!named) {
      data[source.key] = '';
      data.missing.push(source.range);
      Logger.log(`T6 Export Hub: named range "${source.range}" does not exist`);
      return;
    }

    // getDisplayValue, not getValue: these cells are TEXTJOIN formulas and the
    // displayed text is exactly what belongs in the WeakAura.
    const value = String(named.getRange().getDisplayValue()).trim();
    data[source.key] = value;

    if (value) {
      data.filled++;
    } else {
      data.missing.push(`${source.range} (empty)`);
      Logger.log(`T6 Export Hub: "${source.range}" is empty`);
    }
  });

  return data;
}

/**
 * Opens the sidebar at the end of an assignment run, but only when there is
 * something in it.
 *
 * Called by runAutoAssignGroups_ BEFORE it shows the result dialog, so the
 * dialog lands on top and the sidebar is already waiting once it is closed.
 * Every failure is swallowed: a run that assigned 120 rows correctly must not
 * be reported as broken because a named range for an export string is missing.
 */
function openT6ExportHubAfterRun_(ss) {
  try {
    const data = buildT6ExportData_(ss);
    if (!data.filled) {
      Logger.log('T6 Export Hub: nothing to show, so the sidebar was not opened');
      return;
    }
    // Logged on the way in as well as on failure. Without this, "no T6 lines
    // in the log" could mean the function never ran OR that it ran perfectly,
    // which are opposite problems.
    Logger.log(`T6 Export Hub: opening the sidebar after the run (${data.filled} string(s))`);
    showT6ExportSidebar_(data);
  } catch (err) {
    Logger.log(`T6 Export Hub: could not open after the run. ${err.message}`);
  }
}

/** Entry point. Assign this to a drawing or a menu item. */
function showT6ExportHub() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast('Reading the export strings.', '⏳ Please wait...', 5);
  showT6ExportSidebar_(buildT6ExportData_(ss));
}

function showT6ExportSidebar_(data) {
  // Must match the HTML file's name in the editor exactly, without extension.
  // The script file is T6ExportHub and this one is T6Export, so they differ on
  // purpose: change this string if the HTML file is ever renamed.
  const template = HtmlService.createTemplateFromFile('T6Export');
  template.data = data;

  // A sidebar, docked, so the sheet stays usable and the strings stay put
  // while you paste them into the WeakAura one at a time.
  //
  // Google hard-codes this to 300px ("All sidebars shown by scripts are 300
  // pixels wide") and ignores setWidth, so the HTML is written to that width
  // rather than fighting it. A modeless dialog would take a width but floats
  // over the sheet instead of docking, which is worse for this job.
  //
  // The title comes from the output here, not from a second argument the way
  // showModalDialog takes one.
  SpreadsheetApp.getUi().showSidebar(
    template.evaluate().setTitle(T6_EXPORT_DIALOG_TITLE));
}