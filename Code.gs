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
    else if (action === 'clearCache')   { _clearCacheLarge('hierarchy_v1'); result = { ok:true, message:'Cache cleared' }; }
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
// API: Full hierarchy in ONE call (cached 6h with chunking)
// Returns { districts:[], blocks:{d:[...]}, schools:{d_b:[...]} }
// ─────────────────────────────────────────
function getHierarchy() {
  var dists = getDistricts();
  return { districts: dists, blocks: {}, schools: {} };
}

// ─────────────────────────────────────────
// API: Unique Districts (cached 6h)
// ─────────────────────────────────────────
function getDistricts() {
  var cache = CacheService.getScriptCache();
  var cached = cache.get('districts_v5');
  if (cached) return JSON.parse(cached);

  var rows = _getDataRows();
  var seen = {};
  rows.forEach(function(row) {
    var d = _c(row[1]); // Col B = District Name
    if (d) seen[d] = true;
  });
  var result = Object.keys(seen).sort();
  try { cache.put('districts_v5', JSON.stringify(result), 21600); } catch(e){}
  return result;
}

// ─────────────────────────────────────────
// API: Blocks for a District (cached 6h)
// ─────────────────────────────────────────
function getBlocks(district) {
  if (!district) return [];
  var key = 'blk_v5_' + Utilities.base64EncodeWebSafe(_c(district));
  var cache = CacheService.getScriptCache();
  var cached = cache.get(key);
  if (cached) return JSON.parse(cached);

  var rows = _getDataRows();
  var seen = {};
  rows.forEach(function(row) {
    if (_c(row[1]) === district) { // Col B = District Name
      var b = _c(row[2]);          // Col C = Block Name
      if (b) seen[b] = true;
    }
  });
  var result = Object.keys(seen).sort();
  try { cache.put(key, JSON.stringify(result), 21600); } catch(e){}
  return result;
}

// ─────────────────────────────────────────
// API: Schools for District (+ optional Block) (cached 6h)
// ─────────────────────────────────────────
function getSchools(district, block) {
  if (!district) return [];
  var d = _c(district);
  var b = _c(block);
  var key = 'sch_v5_' + Utilities.base64EncodeWebSafe(d + (b ? '_' + b : ''));
  var cache = CacheService.getScriptCache();
  var cached = cache.get(key);
  if (cached) return JSON.parse(cached);

  var rows = _getDataRows();
  var seen = {};
  rows.forEach(function(row) {
    if (_c(row[1]) === d) { // Col B = District
      if (!b || _c(row[2]) === b) { // Col C = Block
        var s = _c(row[4]); // Col E = Last School
        if (s) seen[s] = true;
      }
    }
  });
  var result = Object.keys(seen).sort();
  try { cache.put(key, JSON.stringify(result), 21600); } catch(e){}
  return result;
}

function _stuCacheKey(district, block, school) {
  return 'stu_v5_' + Utilities.base64EncodeWebSafe(_c(district) + '_' + _c(block) + '_' + _c(school));
}

// ─────────────────────────────────────────
// API: Students (cached 2h)
// ─────────────────────────────────────────
function getStudents(district, block, school) {
  if (!district || !school) return { students: [], total: 0, admitted: 0, verified: 0, pending: 0 };
  var d = _c(district);
  var b = _c(block);
  var s = _c(school);
  var key = _stuCacheKey(d, b, s);
  var cache = CacheService.getScriptCache();
  var cached = cache.get(key);
  if (cached) return JSON.parse(cached);

  var rows;
  try { rows = _getDataRows(); } catch(e) { throw new Error('_getDataRows: ' + e.message); }
  var admMap = _getAdmMap(); // PEN -> admission info

  var students = [];
  rows.forEach(function(row) {
    if (_c(row[1]) !== d) return;        // Col B = District
    if (b && _c(row[2]) !== b) return;   // Col C = Block
    if (_c(row[4]) !== s) return;        // Col E = Last School

    var pen      = _c(row[0]);           // Col A = Student PEN
    var name     = _c(row[5]);           // Col F = Student Name
    var gender   = _c(row[6]);           // Col G = Sex
    var mobile   = _c(row[7]);           // Col H = Mobile No
    var mother   = _c(row[8]);           // Col I = Mother Name
    var father   = _c(row[9]);           // Col J = Father Name
    var subStat  = _c(row[10]);          // Col K = Sub Status
    var lastCls  = _c(row[11]);          // Col L = Last Class
    var eligCls  = _c(row[12]);          // Col M = Eligible Class
    var verified = _c(String(row[16])).toLowerCase() === 'yes'; // Col Q = Verified
    var adm      = pen && admMap[pen] ? admMap[pen] : null;

    students.push({
      district:      _c(row[1]),
      block:         _c(row[2]),
      lastSchool:    _c(row[4]),
      pen:           pen,
      name:          name,
      gender:        gender,
      mobile:        mobile,
      motherName:    mother,
      fatherName:    father,
      subStatus:     subStat,
      lastClass:     lastCls,
      eligibleClass: eligCls,
      status:        adm ? 'Admitted' : (verified ? 'Verified' : 'Pending'),
      verInfo:       verified ? { timestamp: _c(String(row[17])) } : null,
      admInfo:       adm
    });
  });

  var result = {
    students: students,
    total:    students.length,
    admitted: students.filter(function(st){ return st.status === 'Admitted';  }).length,
    verified: students.filter(function(st){ return st.status === 'Verified';  }).length,
    pending:  students.filter(function(st){ return st.status === 'Pending';   }).length
  };

  try { cache.put(key, JSON.stringify(result), 7200); } catch(e){}
  return result;
}

// ─────────────────────────────────────────
// API: Save Admission to Admission sheet (Ultra-fast textFinder)
// ─────────────────────────────────────────
function saveAdmission(data) {
  try {
    if (!data.pen) throw new Error('PEN is required');

    var sheet = _getOrCreateVerSheet();
    var finder = sheet.getRange("B:B").createTextFinder(String(data.pen).trim()).matchEntireCell(true);
    var foundCell = finder.findNext();
    var existingRow = foundCell ? foundCell.getRow() : -1;

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
    
    // Invalidate student list cache for this school
    if (data.district && (data.lastSchool || data.school)) {
      try { CacheService.getScriptCache().remove(_stuCacheKey(data.district, data.block || '', data.lastSchool || data.school)); } catch(e){}
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
// API: Save Verification (Search Col A & F for PEN + Batch Write)
// Writes "Yes" to col Q (17) and timestamp to col R (18)
// ─────────────────────────────────────────
function saveVerification(data) {
  try {
    if (!data.pen) throw new Error('PEN is required');

    var ss    = _getSS();
    var sheet = ss.getSheetByName(DATA_SHEET_NAME) || ss.getSheets()[0];
    if (!sheet) throw new Error('Data sheet not found');

    var penStr = String(data.pen).trim();

    // 1. Search Column A (Student PEN)
    var finder = sheet.getRange("A:A").createTextFinder(penStr).matchEntireCell(true);
    var foundCell = finder.findNext();

    // 2. Fallback to Column F
    if (!foundCell) {
      finder = sheet.getRange("F:F").createTextFinder(penStr).matchEntireCell(true);
      foundCell = finder.findNext();
    }

    // 3. Fallback to entire sheet search
    if (!foundCell) {
      finder = sheet.createTextFinder(penStr).matchEntireCell(true);
      foundCell = finder.findNext();
    }

    if (!foundCell) throw new Error('PEN not found in sheet: ' + data.pen);

    var targetRow = foundCell.getRow();
    var tz = Session.getScriptTimeZone();
    var ts = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm');

    // Save target admission class into Column M (Col 13 = Eligible Class to Import)
    // ONLY set if student is Not Studying AND willing to study! Clear/blank for all others.
    var isWilling = (data.currentStatus === 'Not Studying' || data.isStudying === 'NotStudying') && (data.willing === 'Yes' || data.wantsToStudy === 'Yes');
    var targetCls = isWilling ? (data.targetClass || data.admClass || data.eligibleClass || '') : '';
    sheet.getRange(targetRow, 13).setValue(targetCls ? String(targetCls).trim() : '');

    // Single batch write for Cols Q & R (Cols 17 & 18 = Verified Yes & Timestamp)
    sheet.getRange(targetRow, 17, 1, 2).setValues([['Yes', ts]]);

    if (data.district && (data.lastSchool || data.school)) {
      try { CacheService.getScriptCache().remove(_stuCacheKey(data.district, data.block || '', data.lastSchool || data.school)); } catch(e){}
    }

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
  }
  // Always ensure headers in row 1
  var hdrs = ['Timestamp','PEN','Student Name','District','Block','School','Gender','Class','Stream','Cycle','Admission Date'];
  var firstRow = sheet.getRange(1, 1, 1, hdrs.length).getValues()[0];
  if (!firstRow[0] || String(firstRow[0]).trim() === '') {
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
