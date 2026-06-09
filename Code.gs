// ============================================================
// Code.gs — Dropout Verification Portal
// Standalone Google Apps Script — Web App Backend
// ============================================================

var SPREADSHEET_ID         = '11c6hAGriR8VGEIItBqx0sNrCCNqg0D1IrvHPN1xiv7A';
var DATA_SHEET_NAME        = 'Dropout list '; // trailing space — actual sheet name
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
    if      (action === 'ping')         result = { ok:true, sheets: SpreadsheetApp.openById(SPREADSHEET_ID).getSheets().map(function(s){return s.getName()}) };
    else if (action === 'clearCache')   { CacheService.getScriptCache().remove('hierarchy'); result = { ok:true, message:'Cache cleared' }; }
    else if (action === 'getHierarchy') result = getHierarchy();
    else if (action === 'getDistricts') result = getDistricts();
    else if (action === 'getBlocks')    result = getBlocks(e.parameter.district);
    else if (action === 'getSchools')   result = getSchools(e.parameter.district, e.parameter.block);
    else if (action === 'getStudents')  result = getStudents(e.parameter.district, e.parameter.block, e.parameter.school);
    else if (action === 'save')         result = saveVerification(JSON.parse(e.parameter.data));
    else if (action === 'admit')        result = saveAdmission(JSON.parse(e.parameter.data));
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

  try { cache.put('hierarchy', JSON.stringify(result), 21600); } catch(e) {} // ignore if too large
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
// API: Students — col Q=Verified, Admission sheet=Admitted
// ─────────────────────────────────────────
function getStudents(district, block, school) {
  var rows;
  try { rows = _getDataRows(); } catch(e) { throw new Error('_getDataRows: ' + e.message); }
  var admMap = _getAdmMap(); // PEN -> admission info

  var students = [];
  rows.forEach(function(row) {
    if (_c(row[0]) !== district) return;
    if (_c(row[1]) !== block)    return;
    if (_c(row[4]) !== school)   return;

    var pen      = _c(row[5]);
    var verified = _c(String(row[16])).toLowerCase() === 'yes';
    var adm      = pen && admMap[pen] ? admMap[pen] : null;

    students.push({
      district:      _c(row[0]),
      block:         _c(row[1]),
      lastSchool:    _c(row[4]),
      pen:           pen,
      name:          _c(row[7]),
      gender:        _c(row[8]),
      mobile:        _c(row[9]),
      motherName:    _c(row[10]),
      fatherName:    _c(row[11]),
      subStatus:     _c(row[12]),
      lastClass:     _c(row[13]),
      eligibleClass: _c(row[14]),
      status:        adm ? 'Admitted' : (verified ? 'Verified' : 'Pending'),
      verInfo:       verified ? { timestamp: _c(String(row[17])) } : null,
      admInfo:       adm
    });
  });

  return {
    students: students,
    total:    students.length,
    admitted: students.filter(function(s){ return s.status === 'Admitted';  }).length,
    verified: students.filter(function(s){ return s.status === 'Verified';  }).length,
    pending:  students.filter(function(s){ return s.status === 'Pending';   }).length
  };
}

// ─────────────────────────────────────────
// API: Save Admission to Admission sheet
// ─────────────────────────────────────────
function saveAdmission(data) {
  try {
    var sheet = _getOrCreateVerSheet(); // creates/gets Admission sheet
    var allData = sheet.getDataRange().getValues();

    var existingRow = -1;
    for (var i = 1; i < allData.length; i++) {
      if (_c(String(allData[i][1])) === _c(String(data.pen))) {
        existingRow = i + 1; break;
      }
    }

    var tz  = Session.getScriptTimeZone();
    var row = [
      new Date(),               // A Timestamp
      data.pen          || '',  // B PEN
      data.studentName  || '',  // C Name
      data.district     || '',  // D District
      data.block        || '',  // E Block
      data.lastSchool   || '',  // F School
      data.gender       || '',  // G Gender
      data.admClass     || '',  // H Class
      data.stream       || '',  // I Stream
      data.cycle        || '',  // J Cycle
      data.admDate      || '',  // K Admission Date
    ];

    if (existingRow > 0) {
      sheet.getRange(existingRow, 1, 1, row.length).setValues([row]);
    } else {
      sheet.appendRow(row);
    }
    return { success: true, message: 'Admission saved!' };
  } catch(err) {
    return { success: false, message: 'Error: ' + err.toString() };
  }
}

// ─────────────────────────────────────────
// HELPER: PEN -> admission info map
// ─────────────────────────────────────────
function _getAdmMap() {
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
      admClass:  _c(String(r[7])),
      stream:    _c(String(r[8])),
      cycle:     _c(String(r[9])),
      admDate:   _c(String(r[10])),
      timestamp: r[0] instanceof Date ? Utilities.formatDate(r[0], tz, 'dd/MM/yyyy HH:mm') : ''
    };
  }
  return map;
}

// ─────────────────────────────────────────
// API: Save Verification
// Writes "Yes" to col Q (17) and timestamp to col R (18)
// in the source "Dropout list " sheet, matched by PEN (col F = 6)
// ─────────────────────────────────────────
function saveVerification(data) {
  try {
    var ss    = _getSS();
    var sheet = ss.getSheetByName(DATA_SHEET_NAME) || ss.getSheets()[0];
    if (!sheet) throw new Error('Data sheet not found');

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) throw new Error('No data rows found');

    var penCol   = 6;   // column F = PEN (1-based)
    var verCol   = 17;  // column Q = Verified Yes/No (1-based)
    var tsCol    = 18;  // column R = Timestamp (1-based)
    var lastCol  = Math.max(sheet.getLastColumn(), tsCol);
    var pens     = sheet.getRange(2, penCol, lastRow - 1, 1).getValues();

    var targetRow = -1;
    for (var i = 0; i < pens.length; i++) {
      if (_c(String(pens[i][0])) === _c(String(data.pen))) {
        targetRow = i + 2; // +2 because data starts at row 2
        break;
      }
    }

    if (targetRow === -1) throw new Error('PEN not found: ' + data.pen);

    var tz = Session.getScriptTimeZone();
    var ts = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm');

    sheet.getRange(targetRow, verCol).setValue('Yes');
    sheet.getRange(targetRow, tsCol).setValue(ts);

    return { success: true, message: 'Student verified!', row: targetRow };

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
  if (!sheet) throw new Error("Sheet '" + DATA_SHEET_NAME + "' not found. Sheets available: " + ss.getSheets().map(function(s){return s.getName()}).join(', '));
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return [];
  var cols = Math.max(lastCol, 18); // read at least 18 cols (col Q=verified, col R=timestamp)
  return sheet.getRange(2, 1, lastRow - 1, cols).getValues();
}

function _getOrCreateVerSheet() {
  var ss    = _getSS();
  var sheet = ss.getSheetByName(VERIFICATIONS_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(VERIFICATIONS_SHEET);
    var hdrs = ['Timestamp','PEN','Student Name','District','Block','School','Gender','Class','Stream','Cycle','Admission Date'];
    sheet.getRange(1, 1, 1, hdrs.length).setValues([hdrs])
         .setFontWeight('bold').setBackground('#4361ee').setFontColor('white');
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
