// ============================================================
//  Attendance Scanner — Google Apps Script
//  VERSION: 2026-06-15-v3
//  Columns: Timestamp | Name | Email | Location | Handphone Number
//  Sheet: https://docs.google.com/spreadsheets/d/113KIEqWC3NyGhiVz99Dg4lREqvcrVBTYuRktw5y6KGo
//
//  ⚠️ DEPLOY INSTRUCTIONS ⚠️
//  1. Paste this ENTIRE file, replacing all existing code
//  2. Save (Ctrl+S)
//  3. Click Deploy > New deployment (create a FRESH one, don't edit old)
//  4. Type: Web app | Execute as: Me | Who has access: Anyone
//  5. Click Deploy, copy the NEW URL
//  6. Test by opening the URL — must show version "2026-06-15-v3"
// ============================================================

var SPREADSHEET_ID = "113KIEqWC3NyGhiVz99Dg4lREqvcrVBTYuRktw5y6KGo";
var SHEET_GID      = 246908960;
var SHEET_NAME     = "Form_Responses";
var VERSION        = "2026-06-15-v3";

function doGet(e) {
  var output;
  try {
    var p = e.parameter || {};

    // Test ping — shows version so we can confirm deployment
    if (!p.name) {
      output = {
        status: "ok",
        version: VERSION,
        message: "Script is running! Version " + VERSION
      };
      return respond(output, p.callback);
    }

    var sheet = getTargetSheet();

    // Ensure spreadsheet timezone is Singapore
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    if (ss.getSpreadsheetTimeZone() !== "Asia/Singapore") {
      ss.setSpreadsheetTimeZone("Asia/Singapore");
    }

    // Create header if sheet is completely empty
    if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, 5).setValues([["Timestamp", "Name", "Email", "Location", "Handphone Number"]]);
      sheet.getRange(1, 1, 1, 5)
        .setFontWeight("bold")
        .setBackground("#4a1a8c")
        .setFontColor("#ffffff");
      sheet.setFrozenRows(1);
    }

    // Clean phone — strip +65 / 65 prefix, keep digits only
    var rawPhone = (p.phone || "").toString().replace(/\D/g, "");
    if (rawPhone.length === 10 && rawPhone.substring(0,2) === "65") {
      rawPhone = rawPhone.substring(2);
    }

    var name     = (p.name || "").toString();
    var email    = (p.email || "").toString();
    var location = (p.location || "").toString();
    var now      = new Date();

    // ── Find the actual next empty row by checking column B (Name) ──
    // This avoids issues where other columns have stray data
    var lastRow = sheet.getLastRow();
    var nextRow = lastRow + 1;

    // Build the row explicitly as an array of exactly 5 values
    var rowData = [now, name, email, location, rawPhone];

    // Write each cell individually to GUARANTEE correct placement
    sheet.getRange(nextRow, 1).setValue(now);                          // A: Timestamp
    sheet.getRange(nextRow, 1).setNumberFormat("dd/MM/yyyy HH:mm:ss");
    sheet.getRange(nextRow, 2).setValue(name);                          // B: Name
    sheet.getRange(nextRow, 3).setValue(email);                         // C: Email
    sheet.getRange(nextRow, 4).setValue(location);                      // D: Location
    sheet.getRange(nextRow, 5).setValue(rawPhone);                      // E: Handphone Number

    output = {
      status: "ok",
      version: VERSION,
      row: nextRow,
      written: rowData,
      message: "Logged: " + name + " at row " + nextRow
    };

  } catch(err) {
    output = { status: "error", version: VERSION, message: err.toString() };
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
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getSheetId() === SHEET_GID) return sheets[i];
  }
  var byName = ss.getSheetByName(SHEET_NAME);
  if (byName) return byName;
  return ss.insertSheet(SHEET_NAME);
}
