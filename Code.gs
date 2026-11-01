/**
 * Factory Check-in — Apps Script backend
 * Handles: GPS + selfie + name + location logging, GPS-to-branch matching,
 * routing each check-in into its own branch sheet, and an hourly automated
 * sweep that flags anyone who missed their slot.
 *
 * ============================================================
 * SETUP:
 * ============================================================
 *
 * 1. BRANCH COORDINATES (below in BRANCHES) — already filled in with real
 *    GPI, Paras, and Zahir coordinates.
 *
 * 2. BRANCH SHEETS — this script automatically creates and writes to
 *    three separate tabs, named exactly "GPI", "Paras", and "Zahir".
 *    Every check-in is routed into ONE of these three, based strictly
 *    on the location the scanner selected in the dropdown — the GPS
 *    match is still recorded as a column for reference, but never
 *    changes which sheet the row goes into.
 *
 * 3. ROSTER TAB — this script expects a tab named exactly "Roster" with
 *    TWO columns starting at row 1 (no header needed):
 *      Column A = employee name
 *      Column B = their branch — must be exactly "GPI", "Paras", or "Zahir"
 *    Example row: A1 = "Vivek", B1 = "GPI"
 *    Names must match what employees type into the Name field on the
 *    check-in page (case-insensitive, but spelling must match).
 *
 * 4. TRIGGER — for the missed-check-in sweep to run automatically:
 *    Apps Script editor > clock icon (Triggers) on the left > "+ Add
 *    Trigger" > function "markMissedCheckins" > Event source:
 *    "Time-driven" > Type: "Hours timer" > Every hour > Save.
 *
 * After changing this code, redeploy: Deploy > Manage deployments >
 * pencil icon > Version: New version > Deploy.
 */

var BRANCHES = [
  { name: 'GPI',   lat: 27.9174447, lon: 80.7929478 },
  { name: 'Paras', lat: 27.9119047, lon: 80.7949363 },
  { name: 'Zahir', lat: 27.9098534, lon: 80.7976292 }
];
var MATCH_RADIUS_METERS = 200;

var PHOTO_FOLDER_NAME = 'Check-in Selfies';
var ROSTER_SHEET_NAME = 'Roster';
var BRANCH_NAMES = ['GPI', 'Paras', 'Zahir'];
var HEADER_ROW = ['Timestamp', 'Slot', 'Latitude', 'Longitude', 'Accuracy (m)', 'Name', 'Selected Location', 'GPS Match', 'Remark', 'Selfie Link', 'Maps Link'];

// ============================================================
// Handles each check-in submitted from the web page
// ============================================================
function doPost(e) {
  var data = JSON.parse(e.postData.contents);
  var lat = data.lat;
  var lon = data.lon;
  var accuracy = data.accuracy;
  var locationName = data.locationName || '';
  var personName = data.personName || '';
  var slotLabel = data.slotLabel || '';
  var photoDataUrl = data.photo || '';
  var mapsLink = 'https://www.google.com/maps?q=' + lat + ',' + lon;

  var gpsMatch = matchBranch(lat, lon);

  var selfieLink = '';
  if (photoDataUrl) {
    selfieLink = savePhotoToDrive(photoDataUrl, locationName, personName);
  }

  // Routing is ALWAYS by the dropdown selection, never by GPS match.
  var targetSheet = getOrCreateBranchSheet(locationName);

  targetSheet.appendRow([new Date(), slotLabel, lat, lon, accuracy, personName, locationName, gpsMatch, 'On Time', selfieLink, mapsLink]);

  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', gpsMatch: gpsMatch, routedTo: locationName }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// Branch sheet helpers
// ============================================================
function getOrCreateBranchSheet(locationName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = locationName; // dropdown restricts this to GPI/Paras/Zahir
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADER_ROW);
  }
  return sheet;
}

// ============================================================
// GPS-to-branch matching (Haversine distance) — for reference only,
// does NOT decide which sheet a row goes into
// ============================================================
function matchBranch(lat, lon) {
  for (var i = 0; i < BRANCHES.length; i++) {
    var b = BRANCHES[i];
    var dist = haversineMeters(lat, lon, b.lat, b.lon);
    if (dist <= MATCH_RADIUS_METERS) {
      return b.name;
    }
  }
  return 'Not in factory';
}

function haversineMeters(lat1, lon1, lat2, lon2) {
  var R = 6371000;
  var toRad = function (deg) { return deg * Math.PI / 180; };
  var dLat = toRad(lat2 - lat1);
  var dLon = toRad(lon2 - lon1);
  var a = Math.sin(dLat / 2) * Math.sin(dLat / 2) +
          Math.cos(toRad(lat1)) * Math.cos(toRad(lat2)) *
          Math.sin(dLon / 2) * Math.sin(dLon / 2);
  var c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
  return R * c;
}

// ============================================================
// Selfie storage in Drive
// ============================================================
function savePhotoToDrive(dataUrl, locationName, personName) {
  var parts = dataUrl.split(',');
  var meta = parts[0];
  var base64Data = parts[1];
  var contentType = meta.split(':')[1].split(';')[0];

  var bytes = Utilities.base64Decode(base64Data);
  var safeName = (personName || 'unknown').replace(/[^a-zA-Z0-9]/g, '_');
  var safeLoc = (locationName || '').replace(/[^a-zA-Z0-9]/g, '_');
  var fileName = 'checkin_' + new Date().getTime() + '_' + safeName + (safeLoc ? '_' + safeLoc : '') + '.jpg';
  var blob = Utilities.newBlob(bytes, contentType, fileName);

  var folder = getOrCreatePhotoFolder();
  var file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  return file.getUrl();
}

function getOrCreatePhotoFolder() {
  var folders = DriveApp.getFoldersByName(PHOTO_FOLDER_NAME);
  if (folders.hasNext()) {
    return folders.next();
  }
  return DriveApp.createFolder(PHOTO_FOLDER_NAME);
}

// ============================================================
// Missed check-in sweep — run this hourly via a time-driven trigger.
// Now branch-aware: checks each employee's OWN branch sheet, using
// their assigned branch from the Roster tab.
// ============================================================
function markMissedCheckins() {
  var roster = getRosterEntries(); // [{name, branch}, ...]
  if (roster.length === 0) return; // no roster filled in yet, nothing to check

  var now = new Date();
  var minute = now.getMinutes();

  // Google's hourly trigger doesn't fire at an exact minute, so only run
  // the check when we're safely past the window close (:05) and before
  // the next window opens (:55) — this avoids false flags from timing drift.
  if (minute < 15 || minute > 50) return;

  var slotLabel = String(now.getHours()).padStart(2, '0') + ':00';

  // Group roster by branch so we only read each branch sheet once
  var byBranch = {};
  for (var i = 0; i < roster.length; i++) {
    var b = roster[i].branch;
    if (BRANCH_NAMES.indexOf(b) === -1) continue; // skip invalid branch entries
    if (!byBranch[b]) byBranch[b] = [];
    byBranch[b].push(roster[i].name);
  }

  for (var branch in byBranch) {
    checkBranchForMissed(branch, byBranch[branch], slotLabel);
  }
}

function checkBranchForMissed(branchName, namesForBranch, slotLabel) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(branchName);
  if (!sheet) sheet = getOrCreateBranchSheet(branchName); // ensure it exists even with no data yet

  var data = sheet.getDataRange().getValues();
  var presentNames = [];
  var alreadyFlagged = [];
  for (var r = 1; r < data.length; r++) {
    if (String(data[r][1]) === slotLabel) {
      var nameLower = String(data[r][5]).trim().toLowerCase();
      if (String(data[r][8]) === 'Missed') {
        alreadyFlagged.push(nameLower);
      } else {
        presentNames.push(nameLower);
      }
    }
  }

  for (var i = 0; i < namesForBranch.length; i++) {
    var name = namesForBranch[i];
    var nameLower = name.trim().toLowerCase();
    if (presentNames.indexOf(nameLower) === -1 && alreadyFlagged.indexOf(nameLower) === -1) {
      appendMissedRow(sheet, name, slotLabel);
    }
  }
}

function getRosterEntries() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(ROSTER_SHEET_NAME);
  if (!sheet) return [];
  var values = sheet.getRange('A1:B' + sheet.getLastRow()).getValues();
  var entries = [];
  for (var i = 0; i < values.length; i++) {
    var name = String(values[i][0]).trim();
    var branch = String(values[i][1]).trim();
    if (name && branch) entries.push({ name: name, branch: branch });
  }
  return entries;
}

function appendMissedRow(sheet, name, slotLabel) {
  var rowIndex = sheet.getLastRow() + 1;
  sheet.appendRow([new Date(), slotLabel, '', '', '', name, '', '', 'Missed', '', '']);
  var range = sheet.getRange(rowIndex, 1, 1, HEADER_ROW.length);
  range.setBackground('#ffcccc');
}