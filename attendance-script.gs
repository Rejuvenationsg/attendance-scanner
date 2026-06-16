// ============================================================
//  Attendance Scanner — Google Apps Script v4
//  SHEET: https://docs.google.com/spreadsheets/d/113KIEqWC3NyGhiVz99Dg4lREqvcrVBTYuRktw5y6KGo
//
//  ⚠️  DEPLOYMENT — DO THIS EXACTLY:
//  1. Paste this code, click Save
//  2. Click Deploy > New deployment
//  3. Select type: Web app
//  4. Execute as: Me
//  5. Who has access: Anyone   ← MUST BE "Anyone" not "Anyone with Google account"
//  6. Click Deploy > Authorize > Allow
//  7. Copy the /exec URL
// ============================================================

var SPREADSHEET_ID = "113KIEqWC3NyGhiVz99Dg4lREqvcrVBTYuRktw5y6KGo";
var SHEET_GID      = 246908960;
var SHEET_NAME     = "Form_Responses";

// Handle GET requests (app sends data as URL parameters)
function doGet(e) {
  return handleRequest(e.parameter || {});
}

// Handle POST requests (fallback)
function doPost(e) {
  var params = {};
  try { params = JSON.parse(e.postData.contents); } catch(x) {
    params = e.parameter || {};
  }
  return handleRequest(params);
}

function handleRequest(p) {
  // Test ping
  if (!p || !p.name) {
    return json({ status: "ok", version: "v4", message: "Running OK" });
  }

  try {
    var sheet = getSheet();

    // Set timezone
    SpreadsheetApp.openById(SPREADSHEET_ID).setSpreadsheetTimeZone("Asia/Singapore");

    // Header if empty
    if (sheet.getLastRow() === 0) {
      sheet.getRange(1,1,1,5).setValues([["Timestamp","Name","Email","Location","Handphone Number"]]);
      sheet.getRange(1,1,1,5).setFontWeight("bold").setBackground("#4a1a8c").setFontColor("#ffffff");
      sheet.setFrozenRows(1);
    }

    // Clean phone — 8 digits only
    var phone = String(p.phone || "").replace(/\D/g,"");
    if (phone.length === 10 && phone.startsWith("65")) phone = phone.slice(2);
    if (phone.length === 11 && phone.startsWith("065")) phone = phone.slice(3);

    var row = sheet.getLastRow() + 1;
    var now = new Date();

    // Write cell by cell to guarantee column placement
    sheet.getRange(row, 1).setValue(now).setNumberFormat("dd/MM/yyyy HH:mm:ss");
    sheet.getRange(row, 2).setValue(String(p.name     || ""));
    sheet.getRange(row, 3).setValue(String(p.email    || ""));
    sheet.getRange(row, 4).setValue(String(p.location || ""));
    sheet.getRange(row, 5).setValue(phone);

    return json({ status: "ok", row: row, name: p.name });

  } catch(err) {
    return json({ status: "error", message: err.toString() });
  }
}

function json(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSheet() {
  var ss     = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getSheetId() === SHEET_GID) return sheets[i];
  }
  return ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
}
