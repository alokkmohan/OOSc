// ============================================================
// Code.gs — Dropout Verification Portal
// Standalone Google Apps Script — Web App Backend
// ============================================================

var SPREADSHEET_ID         = '11c6hAGriR8VGEIItBqx0sNrCCNqg0D1IrvHPN1xiv7A';
var DATA_SHEET_NAME        = 'Dropout list';
var VERIFICATIONS_SHEET    = 'Admission';

// ─────────────────────────────────────────
// WEB APP ENTRY POINT
// Handles both: Apps Script HTML serving AND
// GitHub Pages fetch() API calls
// ─────────────────────────────────────────
function doGet(e) {
  var action = e.parameter.action;

  // No action = serve the embedded HTML (direct Apps Script access)
  if (!action) {
    return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('Dropout Verification Portal')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  // URL-based JSON API (called from GitHub Pages via fetch)
  try {
    var result;
    if      (action === 'getHierarchy') result = getHierarchy();
    else if (action === 'getDistricts') result = getDistricts();
    else if (action === 'getBlocks')    result = getBlocks(e.parameter.district);
    else if (action === 'getSchools')   result = getSchools(e.parameter.district, e.parameter.block);
    else if (action === 'getStudents')  result = getStudents(e.parameter.district, e.parameter.block, e.parameter.school);
    else if (action === 'save')         result = saveVerification(JSON.parse(e.parameter.data));
    else                                result = { error: 'Unknown action: ' + action };

    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({ error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ─────────────────────────────────────────
// API: Full hierarchy in ONE call (cached 6h)
// Returns { districts:[], blocks:{d:[...]}, schools:{d_b:[...]} }
// ─────────────────────────────────────────
function getHierarchy() {
  var cache = CacheService.getScriptCache();
  var cached = cache.get('hierarchy');
  if (cached) return JSON.parse(cached);

  var rows = _getDataRows();
  var districts = {}, blocks = {}, schools = {};

  rows.forEach(function(row) {
    var d = _c(row[0]), b = _c(row[1]), s = _c(row[4]);
    if (!d) return;
    districts[d] = true;
    if (b) {
      if (!blocks[d]) blocks[d] = {};
      blocks[d][b] = true;
    }
    if (b && s) {
      var key = d + '||' + b;
      if (!schools[key]) schools[key] = {};
      schools[key][s] = true;
    }
  });

  var result = {
    districts: Object.keys(districts).sort(),
    blocks: {},
    schools: {}
  };
  Object.keys(blocks).forEach(function(d) {
    result.blocks[d] = Object.keys(blocks[d]).sort();
  });
  Object.keys(schools).forEach(function(k) {
    result.schools[k] = Object.keys(schools[k]).sort();
  });

  cache.put('hierarchy', JSON.stringify(result), 21600); // 6 hours
  return result;
}

// ─────────────────────────────────────────
// API: Unique Districts
// ─────────────────────────────────────────
function getDistricts() {
  var rows = _getDataRows();
  var seen = {};
  rows.forEach(function(row) {
    var d = _c(row[0]);
    if (d) seen[d] = true;
  });
  return Object.keys(seen).sort();
}

// ─────────────────────────────────────────
// API: Blocks for a District
// ─────────────────────────────────────────
function getBlocks(district) {
  var rows = _getDataRows();
  var seen = {};
  rows.forEach(function(row) {
    if (_c(row[0]) === district) {
      var b = _c(row[1]);
      if (b) seen[b] = true;
    }
  });
  return Object.keys(seen).sort();
}

// ─────────────────────────────────────────
// API: Schools for District + Block
// ─────────────────────────────────────────
function getSchools(district, block) {
  var rows = _getDataRows();
  var seen = {};
  rows.forEach(function(row) {
    if (_c(row[0]) === district && _c(row[1]) === block) {
      var s = _c(row[4]);
      if (s) seen[s] = true;
    }
  });
  return Object.keys(seen).sort();
}

// ─────────────────────────────────────────
// API: Students for District + Block + School
// ─────────────────────────────────────────
function getStudents(district, block, school) {
  var rows   = _getDataRows();
  var verMap = _getVerMap();

  var students = [];
  rows.forEach(function(row) {
    if (_c(row[0]) !== district) return;
    if (_c(row[1]) !== block)    return;
    if (_c(row[4]) !== school)   return;

    var pen = _c(row[5]);
    var s = {
      district:      _c(row[0]),
      block:         _c(row[1]),
      udiseCode:     _c(row[2]),
      schoolCat:     _c(row[3]),
      lastSchool:    _c(row[4]),
      pen:           pen,
      stateCode:     _c(row[6]),
      name:          _c(row[7]),
      gender:        _c(row[8]),
      mobile:        _c(row[9]),
      motherName:    _c(row[10]),
      fatherName:    _c(row[11]),
      subStatus:     _c(row[12]),
      lastClass:     _c(row[13]),
      eligibleClass: _c(row[14]),
      academicYear:  _c(row[15]),
      status:        'Pending',
      verInfo:       null
    };

    if (pen && verMap[pen]) {
      s.status  = verMap[pen].status;
      s.verInfo = verMap[pen];
    }
    students.push(s);
  });

  return {
    students: students,
    total:    students.length,
    verified: students.filter(function(s){ return s.status === 'Verified'; }).length,
    admitted: students.filter(function(s){ return s.status === 'Admitted'; }).length,
    pending:  students.filter(function(s){ return s.status === 'Pending';  }).length
  };
}

// ─────────────────────────────────────────
// API: Save Verification — marks student as Yes
// ─────────────────────────────────────────
function saveVerification(data) {
  try {
    var sheet   = _getOrCreateVerSheet();
    var allData = sheet.getDataRange().getValues();

    var existingRow = -1;
    for (var i = 1; i < allData.length; i++) {
      if (_c(String(allData[i][1])) === _c(String(data.pen))) {
        existingRow = i + 1;
        break;
      }
    }

    var row = [
      new Date(),               // A  Timestamp
      data.pen          || '',  // B  Student PEN
      data.studentName  || '',  // C  Student Name
      data.district     || '',  // D  District
      data.block        || '',  // E  Block
      data.lastSchool   || '',  // F  Last School
      'Yes'                     // G  Verified
    ];

    if (existingRow > 0) {
      sheet.getRange(existingRow, 1, 1, row.length).setValues([row]);
    } else {
      sheet.appendRow(row);
    }

    return { success: true, message: 'Student verified!' };

  } catch(err) {
    return { success: false, message: 'Error: ' + err.toString() };
  }
}

// ─────────────────────────────────────────
// HELPERS
// ─────────────────────────────────────────

function _c(val) {
  return String(val === null || val === undefined ? '' : val).trim();
}

function _getSS() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function _getDataRows() {
  var ss    = _getSS();
  var sheet = ss.getSheetByName(DATA_SHEET_NAME) || ss.getSheets()[0];
  if (!sheet) throw new Error("Koi sheet nahi mili spreadsheet mein.");
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var cols = Math.max(sheet.getLastColumn(), 16);
  return sheet.getRange(2, 1, lastRow - 1, cols).getValues();
}

function _getOrCreateVerSheet() {
  var ss    = _getSS();
  var sheet = ss.getSheetByName(VERIFICATIONS_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(VERIFICATIONS_SHEET);
    var hdrs = ['Timestamp', 'Student PEN', 'Student Name', 'District', 'Block', 'Last School', 'Verified'];
    sheet.getRange(1, 1, 1, hdrs.length).setValues([hdrs])
         .setFontWeight('bold')
         .setBackground('#4361ee')
         .setFontColor('white');
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, hdrs.length);
  }
  return sheet;
}

function _getVerMap() {
  var ss    = _getSS();
  var sheet = ss.getSheetByName(VERIFICATIONS_SHEET);
  if (!sheet) return {};

  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return {};

  var tz  = Session.getScriptTimeZone();
  var map = {};

  for (var i = 1; i < data.length; i++) {
    var r   = data[i];
    var pen = _c(String(r[1]));
    if (!pen) continue;
    map[pen] = {
      status:    'Verified',
      verified:  _c(String(r[6])),
      timestamp: r[0] instanceof Date
                   ? Utilities.formatDate(r[0], tz, 'dd/MM/yyyy HH:mm') : ''
    };
  }
  return map;
}
