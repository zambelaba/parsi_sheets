const WCL_CONFIG = Object.freeze({
  TOKEN_URL: 'https://www.warcraftlogs.com/oauth/token',
  API_URL: 'https://www.warcraftlogs.com/api/v2/client',
  SHEET_NAME: 'WCL Raid JSON',
  OUTPUT_FOLDER: 'Warcraft Logs Raid JSON',
  MINIMUM_ATTENDANCE_PERCENT: 25,
  RAID_LOGS_SHEET: 'raid_logs',
  RAID_ID_COLUMN: 5,
  ATTENDEES_JSON_COLUMN: 7,
  FIRST_DATA_ROW: 2
});

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Warcraft Logs')
    .addItem('Populate raid_logs attendees', 'main')
    .addSeparator()
    .addItem('Set up JSON sheet', 'setupRaidJsonSheet')
    .addItem('Generate all JSON files', 'generateAllRaidJsonFiles')
    .addItem('Test JSON generator', 'testRaidJsonGenerator')
    .addItem('Test live example report', 'testLiveWarcraftLogsExample')
    .addToUi();
}

/**
 * Production entry point.
 *
 * For every data row in `raid_logs`:
 * - column E must contain a Warcraft Logs report ID or URL;
 * - column G must be completely empty and contain no formula;
 * - column G receives a compact JSON string array of attendee names whose
 *   attendancePercent is strictly greater than 25.
 *
 * Duplicate report IDs are fetched only once per run. If a report fails, its
 * column G cell stays empty and receives an error note so a later run can retry it.
 */
function process_presences() {
  process_attendees();
  process_raid_logs_presence();
  update_attendance_on_ratio_sheet();
}

function process_attendees() {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(1000)) {
    throw new Error('Another raid_logs attendance update is already running.');
  }

  let runLogger;
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    runLogger = createRaidLogsLogger_();
    runLogger.log('INFO', 'RUN_START', '', '', 'Starting raid attendance update.');
    runLogger.flush();

    const sheet = spreadsheet.getSheetByName(WCL_CONFIG.RAID_LOGS_SHEET);
    if (!sheet) {
      throw new Error('Sheet "' + WCL_CONFIG.RAID_LOGS_SHEET + '" was not found.');
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < WCL_CONFIG.FIRST_DATA_ROW) {
      runLogger.log('INFO', 'RUN_COMPLETE', '', '', 'No data rows found.');
      runLogger.flush();
      return {processed: 0, failed: 0, pendingReports: 0};
    }

    const rowCount = lastRow - WCL_CONFIG.FIRST_DATA_ROW + 1;
    const raidInputs = sheet.getRange(
      WCL_CONFIG.FIRST_DATA_ROW,
      WCL_CONFIG.RAID_ID_COLUMN,
      rowCount,
      1
    ).getDisplayValues();
    const outputRange = sheet.getRange(
      WCL_CONFIG.FIRST_DATA_ROW,
      WCL_CONFIG.ATTENDEES_JSON_COLUMN,
      rowCount,
      1
    );
    const existingOutputs = outputRange.getDisplayValues();
    const outputFormulas = outputRange.getFormulas();
    const rowsByReportCode = {};
    let failed = 0;

    raidInputs.forEach(function (row, index) {
      const raidInput = String(row[0] || '').trim();
      const existingOutput = String(existingOutputs[index][0] || '').trim();
      const existingFormula = String(outputFormulas[index][0] || '').trim();

      const sheetRow = WCL_CONFIG.FIRST_DATA_ROW + index;
      if (!raidInput) {
        runLogger.log('DEBUG', 'ROW_SKIPPED', sheetRow, '', 'Column E is empty.');
        return;
      }
      if (existingFormula) {
        runLogger.log(
          'INFO', 'ROW_SKIPPED', sheetRow, '',
          'Column G contains a formula and was left unchanged.'
        );
        return;
      }
      if (existingOutput) {
        runLogger.log(
          'INFO', 'ROW_SKIPPED', sheetRow, '',
          'Column G is already populated and was left unchanged.'
        );
        return;
      }

      const reportCode = extractReportCode_(raidInput);
      if (!reportCode) {
        sheet.getRange(sheetRow, WCL_CONFIG.ATTENDEES_JSON_COLUMN)
          .setNote('Invalid Warcraft Logs report ID or URL in column E.');
        runLogger.log(
          'ERROR', 'INVALID_REPORT_ID', sheetRow, raidInput,
          'Column E does not contain a valid report ID or URL.'
        );
        failed += 1;
        return;
      }

      if (!rowsByReportCode[reportCode]) {
        rowsByReportCode[reportCode] = [];
      }
      rowsByReportCode[reportCode].push(sheetRow);
    });

    const reportCodes = Object.keys(rowsByReportCode);
    let processed = 0;
    runLogger.log(
      'INFO', 'SCAN_COMPLETE', '', '',
      'Found ' + reportCodes.length + ' distinct pending report(s).'
    );
    runLogger.flush();

    reportCodes.forEach(function (reportCode) {
      try {
        const targetRows = rowsByReportCode[reportCode];
        runLogger.log(
          'INFO', 'FETCH_START', targetRows.join(','), reportCode,
          'Fetching report for ' + targetRows.length + ' row(s).'
        );
        runLogger.flush();

        const result = generateRaidJson_(reportCode);
        const compactAttendeesJson = JSON.stringify(result.qualifiedNames);
        runLogger.log(
          'INFO', 'FETCH_SUCCESS', targetRows.join(','), reportCode,
          'Received ' + result.document.attendees.length + ' attendee(s); ' +
            result.qualifiedNames.length + ' are above 25% attendance.'
        );

        targetRows.forEach(function (sheetRow) {
          const cell = sheet.getRange(sheetRow, WCL_CONFIG.ATTENDEES_JSON_COLUMN);
          cell.setNumberFormat('@');
          cell.setValue(compactAttendeesJson);
          cell.clearNote();
          processed += 1;
          runLogger.log(
            'INFO', 'ROW_UPDATED', sheetRow, reportCode,
            'Column G populated with ' + result.qualifiedNames.length + ' name(s).'
          );
        });
      } catch (error) {
        rowsByReportCode[reportCode].forEach(function (sheetRow) {
          sheet.getRange(sheetRow, WCL_CONFIG.ATTENDEES_JSON_COLUMN)
            .setNote('Warcraft Logs error: ' + error.message);
          failed += 1;
        });
        runLogger.log(
          'ERROR', 'FETCH_ERROR', rowsByReportCode[reportCode].join(','), reportCode,
          error.message
        );
        console.error('Failed report ' + reportCode + ': ' + error.stack);
      } finally {
        runLogger.flush();
      }
    });

    SpreadsheetApp.flush();
    const summary = {
      processed: processed,
      failed: failed,
      pendingReports: reportCodes.length
    };
    runLogger.log(
      failed ? 'WARN' : 'INFO', 'RUN_COMPLETE', '', '',
      'Updated ' + processed + ' row(s); ' + failed + ' failed; ' +
        reportCodes.length + ' distinct report(s) processed.'
    );
    runLogger.flush();
    spreadsheet.toast(
      'Updated ' + processed + ' row(s); ' + failed + ' failed.',
      'Warcraft Logs',
      5
    );
    return summary;
  } catch (error) {
    if (runLogger) {
      runLogger.log('ERROR', 'RUN_ERROR', '', '', error.message);
      runLogger.flush();
    }
    throw error;
  } finally {
    lock.releaseLock();
  }
}

function createRaidLogsLogger_() {
  return {
    log: function (level, event, row, reportCode, message) {
      const timestamp = new Date();
      console.log(JSON.stringify({
        timestamp: timestamp.toISOString(),
        level: level,
        event: event,
        row: row,
        reportCode: reportCode,
        message: message
      }));
    },
    flush: function () {}
  };
}

/**
 * Creates a sheet where column A contains one report ID or report URL per row.
 */
function setupRaidJsonSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(WCL_CONFIG.SHEET_NAME);

  if (!sheet) {
    sheet = spreadsheet.insertSheet(WCL_CONFIG.SHEET_NAME);
  }

  sheet.getRange('A1:D1').setValues([[
    'Raid ID or Warcraft Logs URL',
    'JSON file',
    'Status',
    'Generated at'
  ]]);
  sheet.getRange('A1:D1').setFontWeight('bold');
  sheet.setFrozenRows(1);

  if (!sheet.getRange('A2').getValue()) {
    sheet.getRange('A2').setValue('JLtYATQjzR9wc3mH');
  }

  sheet.autoResizeColumns(1, 4);
  spreadsheet.setActiveSheet(sheet);
}

/**
 * Generates one <report ID>.json Drive file for every non-empty row in column A.
 */
function generateAllRaidJsonFiles() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(WCL_CONFIG.SHEET_NAME);

  if (!sheet) {
    throw new Error('Run setupRaidJsonSheet first.');
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    throw new Error('Add at least one Warcraft Logs report ID in column A.');
  }

  const values = sheet.getRange(2, 1, lastRow - 1, 1).getDisplayValues();
  const folder = getOrCreateOutputFolder_();

  values.forEach(function (row, index) {
    const sheetRow = index + 2;
    const input = String(row[0]).trim();

    if (!input) {
      return;
    }

    const reportCode = extractReportCode_(input);
    if (!reportCode) {
      sheet.getRange(sheetRow, 2, 1, 3).setValues([[
        '',
        'Invalid report ID or URL',
        new Date()
      ]]);
      return;
    }

    sheet.getRange(sheetRow, 3).setValue('Loading...');
    SpreadsheetApp.flush();

    try {
      const result = generateRaidJson_(reportCode);
      const file = replaceJsonFile_(folder, reportCode + '.json', result.json);
      replaceJsonFile_(
        folder,
        reportCode + '_attendees.json',
        result.qualifiedNamesJson
      );

      sheet.getRange(sheetRow, 2, 1, 3).setValues([[
        file.getUrl(),
        result.qualifiedNames.length + ' names above 25%; ' +
        result.document.summary.bossPulls + ' boss pulls',
        new Date()
      ]]);
    } catch (error) {
      sheet.getRange(sheetRow, 2, 1, 3).setValues([[
        '',
        'Error: ' + error.message,
        new Date()
      ]]);
    }
  });

  sheet.getRange(2, 4, Math.max(lastRow - 1, 1), 1)
    .setNumberFormat('yyyy-mm-dd hh:mm:ss');
  sheet.autoResizeColumns(1, 4);
}

/**
 * Fetches a report and returns both the JSON string and its object representation.
 */
function generateRaidJson_(reportCode) {
  const report = fetchWarcraftLogsReport_(reportCode);
  const document = buildRaidDocument_(report, new Date());
  const qualifiedNames = buildQualifiedAttendeeNames_(
    document,
    WCL_CONFIG.MINIMUM_ATTENDANCE_PERCENT
  );

  return {
    document: document,
    json: JSON.stringify(document, null, 2),
    qualifiedNames: qualifiedNames,
    qualifiedNamesJson: JSON.stringify(qualifiedNames, null, 2)
  };
}

function buildQualifiedAttendeeNames_(document, minimumAttendancePercent) {
  return document.attendees.filter(function (attendee) {
    return attendee.attendancePercent > minimumAttendancePercent;
  }).map(function (attendee) {
    return attendee.name;
  });
}

function fetchWarcraftLogsReport_(reportCode) {
  const accessToken = getWarcraftLogsAccessToken_();
  const query = [
    'query RaidAttendance($code: String!) {',
    '  reportData {',
    '    report(code: $code, allowUnlisted: true) {',
    '      code',
    '      startTime',
    '      endTime',
    '      masterData(translate: true) {',
    '        actors(type: "Player") {',
    '          id',
    '          name',
    '          server',
    '          subType',
    '        }',
    '      }',
    '      fights {',
    '        id',
    '        name',
    '        encounterID',
    '        kill',
    '        inProgress',
    '        startTime',
    '        endTime',
    '        friendlyPlayers',
    '      }',
    '    }',
    '  }',
    '}'
  ].join('\n');

  const response = UrlFetchApp.fetch(WCL_CONFIG.API_URL, {
    method: 'post',
    contentType: 'application/json',
    headers: {Authorization: 'Bearer ' + accessToken},
    payload: JSON.stringify({
      query: query,
      variables: {code: reportCode}
    }),
    muteHttpExceptions: true
  });

  const statusCode = response.getResponseCode();
  const body = response.getContentText();
  let result;

  try {
    result = JSON.parse(body);
  } catch (error) {
    throw new Error('Warcraft Logs returned invalid JSON (HTTP ' + statusCode + ').');
  }

  if (statusCode < 200 || statusCode >= 300 || (result.errors || []).length) {
    const messages = (result.errors || []).map(function (item) {
      return item.message;
    });
    throw new Error(messages.join('; ') || result.error_description ||
      result.error || 'Warcraft Logs HTTP error ' + statusCode);
  }

  const report = result.data && result.data.reportData && result.data.reportData.report;
  if (!report) {
    throw new Error('Report not found or inaccessible. Private reports need user OAuth.');
  }

  return report;
}

/**
 * Pure transformation kept separate so it can be tested without API credentials.
 * Attendance is the union of friendly players on completed boss encounters.
 */
function buildRaidDocument_(report, generatedAt) {
  const actors = (report.masterData && report.masterData.actors) || [];
  const allFights = report.fights || [];
  let countedFights = allFights.filter(function (fight) {
    return Number(fight.encounterID) > 0 && fight.inProgress !== true;
  });

  // Some unusual reports contain no detected boss encounter. In that case,
  // use completed logged fights so the resulting attendee list is still useful.
  if (!countedFights.length) {
    countedFights = allFights.filter(function (fight) {
      return fight.inProgress !== true;
    });
  }

  const actorById = {};
  actors.forEach(function (actor) {
    actorById[String(actor.id)] = actor;
  });

  const attendeeById = {};
  countedFights.forEach(function (fight) {
    (fight.friendlyPlayers || []).forEach(function (actorId) {
      const key = String(actorId);
      const actor = actorById[key];
      if (!actor) {
        return;
      }

      if (!attendeeById[key]) {
        attendeeById[key] = {
          name: actor.name || '',
          realm: actor.server || '',
          class: actor.subType || '',
          bossPullsAttended: 0,
          killsAttended: 0,
          fights: []
        };
      }

      const attendee = attendeeById[key];
      attendee.bossPullsAttended += 1;
      attendee.killsAttended += fight.kill === true ? 1 : 0;
      attendee.fights.push({
        id: fight.id,
        encounterId: fight.encounterID,
        name: fight.name,
        kill: fight.kill === true
      });
    });
  });

  const pullCount = countedFights.length;
  const attendees = Object.keys(attendeeById).map(function (key) {
    const attendee = attendeeById[key];
    attendee.attendancePercent = pullCount
      ? Number((attendee.bossPullsAttended * 100 / pullCount).toFixed(2))
      : 0;
    return attendee;
  }).sort(function (left, right) {
    return left.name.localeCompare(right.name);
  });

  return {
    raidId: report.code,
    generatedAt: generatedAt.toISOString(),
    report: {
      startTime: new Date(report.startTime).toISOString(),
      endTime: new Date(report.endTime).toISOString()
    },
    summary: {
      attendees: attendees.length,
      bossPulls: pullCount,
      kills: countedFights.filter(function (fight) { return fight.kill === true; }).length
    },
    attendees: attendees
  };
}

function getWarcraftLogsAccessToken_() {
  const cache = CacheService.getScriptCache();
  const cachedToken = cache.get('WCL_ACCESS_TOKEN');
  if (cachedToken) {
    return cachedToken;
  }

  const properties = PropertiesService.getScriptProperties();
  const clientId = properties.getProperty('WCL_CLIENT_ID');
  const clientSecret = properties.getProperty('WCL_CLIENT_SECRET');
  if (!clientId || !clientSecret) {
    throw new Error('Set WCL_CLIENT_ID and WCL_CLIENT_SECRET in Script Properties.');
  }

  const response = UrlFetchApp.fetch(WCL_CONFIG.TOKEN_URL, {
    method: 'post',
    contentType: 'application/x-www-form-urlencoded',
    headers: {
      Authorization: 'Basic ' + Utilities.base64Encode(clientId + ':' + clientSecret)
    },
    payload: {grant_type: 'client_credentials'},
    muteHttpExceptions: true
  });

  const statusCode = response.getResponseCode();
  const result = JSON.parse(response.getContentText());
  if (statusCode < 200 || statusCode >= 300 || !result.access_token) {
    throw new Error(result.error_description || result.error ||
      'Warcraft Logs authentication failed (HTTP ' + statusCode + ').');
  }

  const cacheSeconds = Math.max(Math.min((result.expires_in || 3600) - 60, 21600), 60);
  cache.put('WCL_ACCESS_TOKEN', result.access_token, cacheSeconds);
  return result.access_token;
}

function getOrCreateOutputFolder_() {
  const folders = DriveApp.getFoldersByName(WCL_CONFIG.OUTPUT_FOLDER);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(WCL_CONFIG.OUTPUT_FOLDER);
}

function replaceJsonFile_(folder, fileName, contents) {
  const existingFiles = folder.getFilesByName(fileName);
  if (existingFiles.hasNext()) {
    const existingFile = existingFiles.next();
    existingFile.setContent(contents);
    return existingFile;
  }
  return folder.createFile(fileName, contents, 'application/json');
}

function extractReportCode_(value) {
  const text = String(value || '').trim();
  const urlMatch = text.match(/\/reports\/([A-Za-z0-9]+)/i);
  if (urlMatch) {
    return urlMatch[1];
  }
  return /^[A-Za-z0-9]+$/.test(text) ? text : '';
}

/**
 * Credential-free in-sheet test. It throws on failure and returns a success
 * message on success. The menu exposes this as "Test JSON generator".
 */
function testRaidJsonGenerator() {
  const mockReport = {
    code: 'TEST123',
    startTime: Date.UTC(2026, 7, 1, 18, 0, 0),
    endTime: Date.UTC(2026, 7, 1, 21, 0, 0),
    masterData: {actors: [
      {id: 1, name: 'Alice', server: 'RealmOne', subType: 'Priest'},
      {id: 2, name: 'Bob', server: 'RealmTwo', subType: 'Warrior'},
      {id: 3, name: 'TrashOnly', server: 'RealmOne', subType: 'Rogue'}
    ]},
    fights: [
      {id: 10, name: 'Trash', encounterID: 0, kill: false, friendlyPlayers: [1, 3]},
      {id: 11, name: 'Boss One', encounterID: 101, kill: true, friendlyPlayers: [1, 2]},
      {id: 12, name: 'Boss Two', encounterID: 102, kill: false, friendlyPlayers: [1]}
    ]
  };

  const document = buildRaidDocument_(mockReport, new Date('2026-08-02T00:00:00.000Z'));
  assertEqual_(document.raidId, 'TEST123', 'raidId');
  assertEqual_(document.summary.attendees, 2, 'attendee count');
  assertEqual_(document.summary.bossPulls, 2, 'boss pull count');
  assertEqual_(document.summary.kills, 1, 'kill count');
  assertEqual_(document.attendees[0].name, 'Alice', 'alphabetical attendee order');
  assertEqual_(document.attendees[0].attendancePercent, 100, 'Alice attendance');
  assertEqual_(document.attendees[1].attendancePercent, 50, 'Bob attendance');
  assertEqual_(extractReportCode_('https://www.warcraftlogs.com/reports/JLtYATQjzR9wc3mH'),
    'JLtYATQjzR9wc3mH', 'report URL parsing');
  const qualifiedNames = buildQualifiedAttendeeNames_({attendees: [
    {name: 'Excluded', attendancePercent: 25},
    {name: 'Included', attendancePercent: 25.01}
  ]}, 25);
  assertEqual_(qualifiedNames.length, 1, 'strict attendance threshold count');
  assertEqual_(qualifiedNames[0], 'Included', 'strict attendance threshold name');

  SpreadsheetApp.getUi().alert('All 11 Warcraft Logs JSON generator tests passed.');
  return 'All 11 Warcraft Logs JSON generator tests passed.';
}

/**
 * Authenticated integration test using the example report from the sheet.
 * Requires valid WCL_CLIENT_ID and WCL_CLIENT_SECRET Script Properties.
 */
function testLiveWarcraftLogsExample() {
  const result = generateRaidJson_('JLtYATQjzR9wc3mH');
  assertEqual_(result.document.raidId, 'JLtYATQjzR9wc3mH', 'live report ID');

  if (!result.document.attendees.length) {
    throw new Error('The live report returned no attendees.');
  }
  if (!result.document.summary.bossPulls) {
    throw new Error('The live report returned no boss pulls.');
  }

  const message = 'Live API test passed: ' + result.document.attendees.length +
    ' attendees across ' + result.document.summary.bossPulls + ' boss pulls.';
  SpreadsheetApp.getUi().alert(message);
  return message;
}

function assertEqual_(actual, expected, label) {
  if (actual !== expected) {
    throw new Error(label + ': expected ' + expected + ', got ' + actual);
  }
}

/**
 * Build presence_details from raid_logs.
 *
 * Rows are included when the attendees column contains a valid compact JSON
 * array such as ["Alice","Bob"]. presence_details contains raid description,
 * raid date, then one presence column per roster player.
 */
function process_raid_logs_presence() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const raidLogsSheet = ss.getSheetByName('raid_logs');

  if (!raidLogsSheet) {
    throw new Error('raid_logs sheet not found.');
  }

  const lastRow = raidLogsSheet.getLastRow();
  const lastColumn = raidLogsSheet.getLastColumn();
  if (lastRow < 2) {
    Logger.log('No raid rows found in raid_logs.');
    populatePresenceDetails(ss, [], []);
    return {raids: 0, players: 0};
  }

  const headers = raidLogsSheet.getRange(1, 1, 1, lastColumn).getDisplayValues()[0];
  const dateColumnIndex = findRaidLogColumnIndex_(
    headers,
    ['raid_date', 'raid date', 'date'],
    0
  );
  const attendeesColumnIndex = findRaidLogColumnIndex_(
    headers,
    ['attendees'],
    6
  );

  Logger.log(
    'Using raid_logs column ' + columnNumberToLetter_(dateColumnIndex + 1) +
    ' for raid dates and column ' + columnNumberToLetter_(attendeesColumnIndex + 1) +
    ' for attendees.'
  );

  const rosterPlayers = getRatioPresenceRoster(ss);
  if (rosterPlayers === null) {
    throw new Error('Ratio Presence/Loot roster sheet not found.');
  }
  const rerollMaps = getRerollPresenceMaps_(ss, rosterPlayers);

  const sortedPlayers = Array.from(rosterPlayers.values()).sort(function (left, right) {
    return left.localeCompare(right);
  });
  const displayRows = raidLogsSheet
    .getRange(2, 1, lastRow - 1, lastColumn)
    .getDisplayValues();
  const raids = [];

  for (let i = 0; i < displayRows.length; i++) {
    const sheetRow = i + 2;
    const attendeesJson = String(displayRows[i][attendeesColumnIndex] || '').trim();
    if (!attendeesJson) {
      continue;
    }

    let parsedAttendees;
    try {
      parsedAttendees = JSON.parse(attendeesJson);
    } catch (error) {
      Logger.log('Skipping raid_logs row ' + sheetRow + ': invalid attendees JSON.');
      continue;
    }

    if (!Array.isArray(parsedAttendees)) {
      Logger.log('Skipping raid_logs row ' + sheetRow + ': attendees must be a JSON array.');
      continue;
    }

    const attendeeKeys = new Set();
    for (let attendeeIndex = 0; attendeeIndex < parsedAttendees.length; attendeeIndex++) {
      const attendee = parsedAttendees[attendeeIndex];
      const attendeeName = typeof attendee === 'string'
        ? attendee
        : (attendee && attendee.name) || '';
      const attendeeKey = getRosterPlayerKey(attendeeName);

      if (!attendeeKey) {
        continue;
      }
      const resolvedRosterKey = resolvePresenceAttendeeKey_(
        attendeeName,
        rosterPlayers,
        rerollMaps
      );
      if (!resolvedRosterKey) {
        Logger.log(
          'raid_logs row ' + sheetRow + ': attendee "' + attendeeName +
          '" is not in the roster or Rerolls sheet and was ignored.'
        );
        continue;
      }
      if (!rosterPlayers.has(attendeeKey)) {
        Logger.log(
          'raid_logs row ' + sheetRow + ': reroll "' + attendeeName +
          '" credited to main "' + rosterPlayers.get(resolvedRosterKey) + '".'
        );
      }
      attendeeKeys.add(resolvedRosterKey);
    }

    raids.push({
      sourceRow: sheetRow,
      description: String(displayRows[i][1] || '').trim(),
      date: String(displayRows[i][dateColumnIndex] || '').trim(),
      attendeeKeys: attendeeKeys
    });
  }

  populatePresenceDetails(ss, raids, sortedPlayers);
  Logger.log(
    'Presence processing complete: ' + raids.length + ' raid(s), ' +
    sortedPlayers.length + ' roster player(s).'
  );

  return {raids: raids.length, players: sortedPlayers.length};
}

function populatePresenceDetails(ss, raids, sortedPlayers) {
  let detailsSheet = ss.getSheetByName('presence_details');
  if (!detailsSheet) {
    detailsSheet = ss.insertSheet('presence_details');
  } else {
    detailsSheet
      .getRange(1, 1, detailsSheet.getMaxRows(), detailsSheet.getMaxColumns())
      .breakApart();
    detailsSheet.clear();
  }

  if (sortedPlayers.length === 0) {
    Logger.log('No roster players available for presence_details.');
    return;
  }

  const numColumns = sortedPlayers.length + 2;
  ensurePresenceSheetSize_(detailsSheet, raids.length + 3, numColumns);

  const headerValues = ['raid_description', 'raid_date'].concat(sortedPlayers);
  detailsSheet.getRange(1, 1, 1, numColumns).setValues([headerValues]);
  detailsSheet.setColumnWidth(1, 200);
  detailsSheet.setColumnWidth(2, 85);
  for (let playerIndex = 0; playerIndex < sortedPlayers.length; playerIndex++) {
    detailsSheet.setColumnWidth(playerIndex + 3, 75);
  }

  detailsSheet.getRange(1, 1, 1, numColumns)
    .setFontWeight('bold')
    .setBackground('#073763')
    .setFontColor('#FFFFFF')
    .setHorizontalAlignment('center')
    .setFontSize(10);
  detailsSheet.setFrozenRows(1);
  detailsSheet.setFrozenColumns(2);

  const rowData = buildPresenceDetailsRows_(raids, sortedPlayers);
  const totalSinceP2 = buildDatedPresenceTotalsRow_(
    rowData,
    sortedPlayers.length,
    20260518,
    'TOTAL SINCE P2'
  );
  const totalSinceAutomation = buildDatedPresenceTotalsRow_(
    rowData,
    sortedPlayers.length,
    20260901,
    'TOTAL SINCE AUTOMATION'
  );
  const outputRows = rowData.concat([totalSinceP2, totalSinceAutomation]);
  detailsSheet.getRange(2, 1, outputRows.length, numColumns).setValues(outputRows);
  detailsSheet.getRange(2, 1, outputRows.length, 2).setNumberFormat('@');

  for (let playerIndex = 0; playerIndex < sortedPlayers.length; playerIndex++) {
    const presenceColumn = playerIndex + 3;
    detailsSheet.getRange(2, presenceColumn, outputRows.length, 1).setNumberFormat('0');
    detailsSheet.getRange(1, presenceColumn, outputRows.length + 1, 1)
      .setBorder(null, null, null, true, null, null, 'black', SpreadsheetApp.BorderStyle.SOLID);
  }

  const firstTotalSheetRow = rowData.length + 2;
  detailsSheet.getRange(firstTotalSheetRow, 1, 2, numColumns)
    .setFontWeight('bold')
    .setBackground('#D9EAF7');

  if (raids.length === 0) {
    Logger.log('No raid rows with attendees JSON were found; total rows contain zeroes.');
  }

  Logger.log('presence_details populated with ' + rowData.length + ' raid row(s).');
}

function buildPresenceDetailsRows_(raids, sortedPlayers) {
  const rowData = [];

  for (let raidIndex = 0; raidIndex < raids.length; raidIndex++) {
    const raid = raids[raidIndex];
    const rowValues = [raid.description || '', raid.date || ''];

    for (let playerIndex = 0; playerIndex < sortedPlayers.length; playerIndex++) {
      const playerKey = getRosterPlayerKey(sortedPlayers[playerIndex]);
      rowValues.push(raid.attendeeKeys.has(playerKey) ? 1 : 0);
    }
    rowData.push(rowValues);
  }

  return rowData;
}

function buildDatedPresenceTotalsRow_(rowData, playerCount, cutoffDateKey, label) {
  const totalRow = [label, ''];

  for (let playerIndex = 0; playerIndex < playerCount; playerIndex++) {
    const presenceColumnIndex = playerIndex + 2;
    let totalPresence = 0;
    for (let rowIndex = 0; rowIndex < rowData.length; rowIndex++) {
      const raidDateKey = getPresenceDateKey_(rowData[rowIndex][1]);
      if (raidDateKey !== null && raidDateKey >= cutoffDateKey) {
        totalPresence += Number(rowData[rowIndex][presenceColumnIndex]) || 0;
      }
    }
    totalRow.push(totalPresence);
  }

  return totalRow;
}

/**
 * Copy per-player attendance totals from presence_details to column F of
 * Ratio Présence/Loot (Préloot). Only raids dated 2026-09-01 or later count.
 */
function update_attendance_on_ratio_sheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const presenceSheet = ss.getSheetByName('presence_details');
  if (!presenceSheet) {
    throw new Error('presence_details sheet not found.');
  }

  const ratioSheet = getRatioPresenceSheet(ss);
  if (!ratioSheet) {
    throw new Error('Ratio Présence/Loot (Préloot) sheet not found.');
  }

  const presenceLastRow = presenceSheet.getLastRow();
  const presenceLastColumn = presenceSheet.getLastColumn();
  if (presenceLastRow < 2 || presenceLastColumn < 3) {
    throw new Error('presence_details does not contain player totals.');
  }

  const presenceHeaders = presenceSheet
    .getRange(1, 1, 1, presenceLastColumn)
    .getDisplayValues()[0];
  const presenceRows = presenceSheet
    .getRange(2, 1, presenceLastRow - 1, presenceLastColumn)
    .getValues();
  const cutoffDateKey = 20260901;
  const totalsResult = buildDatedPresenceTotalsMap_(
    presenceHeaders,
    presenceRows,
    cutoffDateKey
  );
  const totalsByPlayer = totalsResult.totalsByPlayer;
  Logger.log(
    'Ratio attendance cutoff: 01/09/2026. Included ' + totalsResult.includedRows +
    ' raid(s); skipped ' + totalsResult.skippedRows + ' earlier/invalid row(s).'
  );

  const ratioLastRow = ratioSheet.getLastRow();
  if (ratioLastRow < 3) {
    Logger.log('No player rows found in Ratio Présence/Loot (Préloot).');
    return {updated: 0, availableTotals: totalsByPlayer.size};
  }

  const ratioPlayerNames = ratioSheet
    .getRange(3, 1, ratioLastRow - 2, 1)
    .getDisplayValues();
  const existingAttendanceValues = ratioSheet
    .getRange(3, 6, ratioLastRow - 2, 1)
    .getValues();
  const updateResult = buildRatioAttendanceValues_(
    ratioPlayerNames,
    existingAttendanceValues,
    totalsByPlayer
  );

  ratioSheet
    .getRange(3, 6, updateResult.values.length, 1)
    .setValues(updateResult.values)
    .setNumberFormat('0');

  Logger.log(
    'Updated Ratio Présence/Loot column F for ' + updateResult.updated +
    ' player(s). ' + updateResult.unmatched + ' roster row(s) had no presence total.'
  );
  return {
    updated: updateResult.updated,
    unmatched: updateResult.unmatched,
    availableTotals: totalsByPlayer.size
  };
}

function buildDatedPresenceTotalsMap_(presenceHeaders, presenceRows, cutoffDateKey) {
  const totalsByPlayer = new Map();

  for (let columnIndex = 2; columnIndex < presenceHeaders.length; columnIndex++) {
    const playerName = String(presenceHeaders[columnIndex] || '').trim();
    const playerKey = getRosterPlayerKey(playerName);
    if (!playerKey) {
      continue;
    }
    totalsByPlayer.set(playerKey, 0);
  }

  let includedRows = 0;
  let skippedRows = 0;
  for (let rowIndex = 0; rowIndex < presenceRows.length; rowIndex++) {
    const row = presenceRows[rowIndex];
    if (String(row[0] || '').trim().toUpperCase().indexOf('TOTAL') === 0) {
      continue;
    }

    const raidDateKey = getPresenceDateKey_(row[1]);
    if (raidDateKey === null || raidDateKey < cutoffDateKey) {
      skippedRows++;
      continue;
    }

    includedRows++;
    for (let columnIndex = 2; columnIndex < presenceHeaders.length; columnIndex++) {
      const playerKey = getRosterPlayerKey(presenceHeaders[columnIndex]);
      if (playerKey && totalsByPlayer.has(playerKey)) {
        totalsByPlayer.set(
          playerKey,
          totalsByPlayer.get(playerKey) + (Number(row[columnIndex]) || 0)
        );
      }
    }
  }

  return {
    totalsByPlayer: totalsByPlayer,
    includedRows: includedRows,
    skippedRows: skippedRows
  };
}

function getPresenceDateKey_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return value.getFullYear() * 10000 + (value.getMonth() + 1) * 100 + value.getDate();
  }

  const text = String(value || '').trim();
  let match = text.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})$/);
  if (match) {
    return Number(match[3]) * 10000 + Number(match[2]) * 100 + Number(match[1]);
  }
  match = text.match(/^(\d{4})[\/-](\d{1,2})[\/-](\d{1,2})$/);
  if (match) {
    return Number(match[1]) * 10000 + Number(match[2]) * 100 + Number(match[3]);
  }
  return null;
}

function buildRatioAttendanceValues_(playerRows, existingValues, totalsByPlayer) {
  const outputValues = existingValues.map(function (row) {
    return [row[0]];
  });
  let updated = 0;
  let unmatched = 0;

  for (let rowIndex = 0; rowIndex < playerRows.length; rowIndex++) {
    const playerName = String(playerRows[rowIndex][0] || '').trim();
    if (!playerName) {
      continue;
    }
    const playerKey = getRosterPlayerKey(playerName);
    if (totalsByPlayer.has(playerKey)) {
      outputValues[rowIndex][0] = totalsByPlayer.get(playerKey);
      updated++;
    } else {
      unmatched++;
    }
  }

  return {values: outputValues, updated: updated, unmatched: unmatched};
}

function getRerollPresenceMaps_(ss, rosterPlayers) {
  const maps = {exact: new Map()};
  const rerollsSheet = ss.getSheetByName('Rerolls');
  if (!rerollsSheet) {
    Logger.log('Rerolls sheet not found; only direct roster names will be matched.');
    return maps;
  }

  const lastRow = rerollsSheet.getLastRow();
  const lastColumn = rerollsSheet.getLastColumn();
  if (lastRow < 2 || lastColumn < 1) {
    Logger.log('Rerolls sheet contains no reroll data.');
    return maps;
  }

  const values = rerollsSheet.getRange(1, 1, lastRow, lastColumn).getDisplayValues();
  const headers = values[0].map(function (value) {
    return String(value || '').trim().toLowerCase().replace(/[\s_-]+/g, '');
  });
  const rerollNameIndex = headers.indexOf('rerollname');
  const mainNameIndex = headers.indexOf('mainname');
  if (rerollNameIndex === -1 || mainNameIndex === -1) {
    Logger.log('Rerolls sheet must contain "Reroll name" and "Main name" headers.');
    return maps;
  }

  let mappedRerolls = 0;
  for (let rowIndex = 1; rowIndex < values.length; rowIndex++) {
    const rerollName = String(values[rowIndex][rerollNameIndex] || '').trim();
    const mainName = String(values[rowIndex][mainNameIndex] || '').trim();
    if (!rerollName || !mainName) {
      continue;
    }

    const exactMainKey = getRosterPlayerKey(mainName);
    if (!rosterPlayers.has(exactMainKey)) {
      Logger.log(
        'Rerolls row ' + (rowIndex + 1) + ': main "' + mainName +
        '" is not in the roster; reroll "' + rerollName + '" was not mapped.'
      );
      continue;
    }

    const exactRerollKey = getRosterPlayerKey(rerollName);
    if (exactRerollKey) {
      maps.exact.set(exactRerollKey, exactMainKey);
    }
    mappedRerolls++;
  }

  Logger.log('Loaded ' + mappedRerolls + ' reroll-to-main mapping(s).');
  return maps;
}

function resolvePresenceAttendeeKey_(attendeeName, rosterPlayers, rerollMaps) {
  const exactKey = getRosterPlayerKey(attendeeName);
  if (rosterPlayers.has(exactKey)) {
    return exactKey;
  }
  if (rerollMaps.exact.has(exactKey)) {
    return rerollMaps.exact.get(exactKey);
  }
  return null;
}

function findRaidLogColumnIndex_(headers, acceptedNames, fallbackIndex) {
  const normalizedAcceptedNames = acceptedNames.map(function (name) {
    return String(name).trim().toLowerCase().replace(/[\s-]+/g, '_');
  });

  for (let index = 0; index < headers.length; index++) {
    const normalizedHeader = String(headers[index] || '')
      .trim()
      .toLowerCase()
      .replace(/[\s-]+/g, '_');
    if (normalizedAcceptedNames.indexOf(normalizedHeader) !== -1) {
      return index;
    }
  }

  Logger.log(
    'Could not find header ' + acceptedNames.join('/') +
    '; falling back to column ' + columnNumberToLetter_(fallbackIndex + 1) + '.'
  );
  return fallbackIndex;
}

function columnNumberToLetter_(columnNumber) {
  let result = '';
  let value = columnNumber;
  while (value > 0) {
    const remainder = (value - 1) % 26;
    result = String.fromCharCode(65 + remainder) + result;
    value = Math.floor((value - 1) / 26);
  }
  return result;
}

function ensurePresenceSheetSize_(sheet, requiredRows, requiredColumns) {
  if (sheet.getMaxRows() < requiredRows) {
    sheet.insertRowsAfter(sheet.getMaxRows(), requiredRows - sheet.getMaxRows());
  }
  if (sheet.getMaxColumns() < requiredColumns) {
    sheet.insertColumnsAfter(
      sheet.getMaxColumns(),
      requiredColumns - sheet.getMaxColumns()
    );
  }
}
