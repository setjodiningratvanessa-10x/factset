// Paste this into your Google Sheet:
// Extensions → Apps Script → replace all content → Save → Deploy

var SECRET_TOKEN = "REPLACE_WITH_YOUR_SECRET_TOKEN"; // set this before deploying

function doPost(e) {
  try {
    if (e.parameter.token !== SECRET_TOKEN) {
      return ContentService.createTextOutput("Unauthorized").setMimeType(ContentService.MimeType.TEXT);
    }

    var payload = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    var sheetName = payload.sheet_name || "Weekly Consensus";
    var ws = ss.getSheetByName(sheetName);
    if (!ws) {
      ws = ss.insertSheet(sheetName);
    }

    ws.clearContents();
    ws.clearFormats();

    var rows = payload.rows;
    if (rows.length > 0) {
      ws.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
    }

    // Apply basic formatting
    // Title row bold
    ws.getRange(1, 1, 1, rows[0].length).setFontWeight("bold").setFontSize(12);

    // Header rows bold
    ws.getRange(5, 1, 1, rows[0].length).setFontWeight("bold");
    ws.getRange(7, 1, 1, rows[0].length).setFontWeight("bold");
    ws.getRange(11, 1, 1, rows[0].length).setFontWeight("bold").setBackground("#d9e1f2");

    // Freeze top row
    ss.setFrozenRows(11);
    ss.setFrozenColumns(2);

    return ContentService.createTextOutput("OK").setMimeType(ContentService.MimeType.TEXT);
  } catch (err) {
    return ContentService.createTextOutput("Error: " + err.message).setMimeType(ContentService.MimeType.TEXT);
  }
}

function doGet(e) {
  return ContentService.createTextOutput("Weekly Consensus Updater is running.");
}
