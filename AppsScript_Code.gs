/**
 * TEAM DASHBOARD - BACKEND (multi-sheet)
 * ------------------------------------------------------
 * Reads 4 tabs from this spreadsheet and serves them as JSON:
 *   - "IB agent wise"   -> call-level audit records
 *   - "Empathy wise "   -> per-agent empathy parameter % (note: sheet name
 *                          has a trailing space in this workbook)
 *   - "ORT"             -> attendance / AHT / CFS / adherence
 *   - "allignment"      -> agent roster (VDI ID <-> Name <-> Team Leader)
 *
 * SETUP (one-time, ~2 minutes):
 * 1. Open your Google Sheet.
 * 2. Extensions > Apps Script
 * 3. Delete any starter code, paste this whole file in.
 * 4. Deploy > New deployment > gear icon > Web app.
 * 5. Execute as: Me. Who has access: Anyone.
 * 6. Deploy, authorize (Advanced > Go to project (unsafe) > Allow).
 * 7. Copy the Web app URL and paste it into the dashboard HTML's
 *    API_URL constant.
 *
 * If you rename a tab, update the SHEET_NAMES map below to match.
 * ------------------------------------------------------
 */

var SHEET_NAMES = {
  agentWise: 'IB agent wise',
  empathyWise: 'Empathy wise ', // trailing space is intentional - matches the tab name
  ort: 'ORT',
  alignment: 'allignment'
};

function doGet(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var payload = {
    generatedAt: new Date().toISOString(),
    debug: {
      availableSheetNames: ss.getSheets().map(function(s){ return s.getName(); })
    },
    agentAudits: readAgentWise(ss),
    empathyWise: readEmpathyWise(ss),
    ort: readOrt(ss),
    alignment: readAlignment(ss)
  };

  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

// Finds a sheet by name, tolerant of extra/missing whitespace and case.
// Falls back to exact getSheetByName first (fast path), then does a
// trimmed/case-insensitive scan so small naming drift doesn't silently
// return zero rows.
function findSheetLoose(ss, name) {
  var exact = ss.getSheetByName(name);
  if (exact) return exact;
  var target = name.toString().trim().toLowerCase();
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getName().toString().trim().toLowerCase() === target) {
      return sheets[i];
    }
  }
  return null;
}

function sheetToObjects(sheet, wantedHeaders) {
  if (!sheet) return [];
  var values = sheet.getDataRange().getDisplayValues();
  if (values.length < 2) return [];
  var headers = values[0].map(function(h){ return (h || '').toString().trim(); });
  var colIndex = {};
  wantedHeaders.forEach(function(name) {
    var idx = headers.indexOf(name.trim());
    if (idx !== -1) colIndex[name] = idx;
  });

  var out = [];
  for (var r = 1; r < values.length; r++) {
    var row = values[r];
    if (row.join('') === '') continue; // skip fully blank rows
    var obj = {};
    var hasKey = false;
    for (var key in colIndex) {
      var v = row[colIndex[key]];
      obj[key] = v;
      if (v !== '' && v !== null) hasKey = true;
    }
    if (!hasKey) continue;
    out.push(obj);
  }
  return out;
}

function readAgentWise(ss) {
  var sheet = findSheetLoose(ss, SHEET_NAMES.agentWise);
  var wanted = [
    'Eval Date', 'Agent Name', 'Team Leader', 'QA', 'VDI ID', 'Site', 'Skill', 'Month', 'Tenure',
    'Pass or Fail', 'Critical Error', 'RTC', 'Audit Party',
    'Total Emp Parameter Pass Count', 'Total Emp Parameter Fail Count', 'Evaluation Score'
  ];
  var rows = sheetToObjects(sheet, wanted);
  // require a real Agent Name to filter out stray blank rows
  return rows.filter(function(r) { return r['Agent Name'] && r['Agent Name'] !== ''; });
}

function readEmpathyWise(ss) {
  var sheet = findSheetLoose(ss, SHEET_NAMES.empathyWise);
  var wanted = [
    'VDI ID', 'Skill', 'TL', 'Audit Count',
    'Warm Greetings', 'Acknowledgment & Understanding',
    'Speech Clarity / Clarity of Communication', 'Proper Listening',
    'Personalization', 'Time / Hold Management',
    'Emotional Tone with Courtesy', 'Closing Courtesy'
  ];
  var rows = sheetToObjects(sheet, wanted);
  return rows.filter(function(r) { return r['VDI ID'] && r['VDI ID'] !== ''; });
}

function readOrt(ss) {
  var sheet = findSheetLoose(ss, SHEET_NAMES.ort);
  var wanted = ['ID', 'Name', 'AHT', 'Lag Time', 'Extra Break', 'Adherence%', 'CFS%'];
  var rows = sheetToObjects(sheet, wanted);
  // exclude "Grand Total" row and blanks
  return rows.filter(function(r) { return r['Name'] && r['Name'] !== '' && r['Name'].toLowerCase().indexOf('grand total') === -1; });
}

function readAlignment(ss) {
  var sheet = findSheetLoose(ss, SHEET_NAMES.alignment);
  var wanted = ['GNX ID', 'Name', 'Gender', 'Supervisor 1', 'VDI ID'];
  var rows = sheetToObjects(sheet, wanted);
  return rows.filter(function(r) { return r['Name'] && r['Name'] !== ''; });
}
