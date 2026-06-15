// ============================================================
//  Attendance Scanner — Google Apps Script
//  Columns: Timestamp | Name | Email | Location | Handphone Number
//  Sheet: https://docs.google.com/spreadsheets/d/113KIEqWC3NyGhiVz99Dg4lREqvcrVBTYuRktw5y6KGo
//  DEPLOY: Execute as Me, Who has access: Anyone
// ============================================================

var SPREADSHEET_ID = "113KIEqWC3NyGhiVz99Dg4lREqvcrVBTYuRktw5y6KGo";
var SHEET_GID      = 246908960;
var SHEET_NAME     = "Form_Responses";

function doGet(e) {
  var output;
  try {
    var p = e.parameter || {};

    // Test ping
    if (!p.name) {
      output = { status: "ok", message: "Script is running!" };
      return respond(output, p.callback);
    }

    var sheet = getTargetSheet();

    // Ensure spreadsheet timezone is Singapore so timestamps align
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    if (ss.getSpreadsheetTimeZone() !== "Asia/Singapore") {
      ss.setSpreadsheetTimeZone("Asia/Singapore");
    }

    // Create header if sheet is empty
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(["Timestamp", "Name", "Email", "Location", "Handphone Number"]);
      sheet.getRange(1, 1, 1, 5)
        .setFontWeight("bold")
        .setBackground("#4a1a8c")
        .setFontColor("#ffffff");
      sheet.setFrozenRows(1);
      sheet.autoResizeColumns(1, 5);
    }

    // Clean phone — strip +65 prefix, keep 8 digits only
    var rawPhone = (p.phone || "").toString().replace(/\D/g, "");
    if (rawPhone.startsWith("65") && rawPhone.length === 10) {
      rawPhone = rawPhone.substring(2); // strip country code
    }

    // Use a real Date object (Singapore timezone) so it matches the
    // existing Form_Responses timestamp column type and format
    var now = new Date();

    // Write to explicit cell range (A:E) — guarantees correct column
    // alignment regardless of any extra/empty trailing columns
    var nextRow = sheet.getLastRow() + 1;
    sheet.getRange(nextRow, 1, 1, 5).setValues([[
      now,                  // A: Timestamp (real date value)
      p.name     || "",     // B: Name
      p.email    || "",     // C: Email
      p.location || "",     // D: Location
      rawPhone              // E: Handphone Number (8 digits only)
    ]]);

    // Format the timestamp cell to match existing rows: DD/MM/YYYY HH:MM:SS
    sheet.getRange(nextRow, 1).setNumberFormat("dd/MM/yyyy HH:mm:ss");

    sheet.autoResizeColumns(1, 5);
    output = { status: "ok", message: "Logged: " + p.name };

  } catch(err) {
    output = { status: "error", message: err.toString() };
  }

  return respond(output, (e.parameter||{}).callback);
}

function respond(data, callback) {
  var json = JSON.stringify(data);
  var content = callback ? callback + "(" + json + ")" : json;
  return ContentService
    .createTextOutput(content)
    .setMimeType(callback ? ContentService.MimeType.JAVASCRIPT : ContentService.MimeType.JSON);
}

function getTargetSheet() {
  var ss     = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheets = ss.getSheets();
  // Try GID first
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getSheetId() === SHEET_GID) return sheets[i];
  }
  // Try name
  var byName = ss.getSheetByName(SHEET_NAME);
  if (byName) return byName;
  // Create new
  return ss.insertSheet(SHEET_NAME);
}
