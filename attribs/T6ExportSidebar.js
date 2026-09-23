/**
 * Sidebar with one copy button per Fojji T6 WeakAura export string. Each string
 * is a TEXTJOIN formula in a named range, read when the sidebar opens, so it
 * needs reopening after an assignment changes.
 */
const T6_EXPORT_SOURCES = [
  { range: 'BLOODBOIL_SOAK_TEAMS', key: 'bloodboilSoak' },
  { range: 'SOULS_INTERRUPTS', key: 'soulsInterrupts' },
  { range: 'COUNCIL_INTERRUPTS', key: 'councilInterrupts' },
];

/** FOJJI EXPORT button on Quick Assigns. */
function showT6ExportHub() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast('Reading the export strings.', '⏳ Please wait...', 5);
  showT6ExportSidebar_(buildT6ExportData_(ss));
}

/** After an auto-assign run: opens the sidebar when there is something to copy, never fails the run. */
function openT6ExportHubAfterRun_(ss) {
  try {
    const data = buildT6ExportData_(ss);
    if (data.filled) showT6ExportSidebar_(data);
  } catch (err) {
    Logger.log(`T6 export sidebar could not open after the run: ${err.message}`);
  }
}

function buildT6ExportData_(ss) {
  const byName = {};
  ss.getNamedRanges().forEach(nr => { byName[nr.getName()] = nr; });

  const data = { missing: [], filled: 0 };
  T6_EXPORT_SOURCES.forEach(source => {
    const named = byName[source.range];
    const value = named ? String(named.getRange().getDisplayValue()).trim() : '';
    data[source.key] = value;
    if (value) data.filled++;
    else data.missing.push(named ? `${source.range} (empty)` : source.range);
  });
  return data;
}

function showT6ExportSidebar_(data) {
  // Sidebars are fixed at 300px wide; T6Export.html is laid out for that.
  const template = HtmlService.createTemplateFromFile('T6Export');
  template.data = data;
  SpreadsheetApp.getUi().showSidebar(template.evaluate().setTitle('T6 Export Strings'));
}
