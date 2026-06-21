const SHEET_NAME = "data";

function getDataSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.getRange("A1").setValue("{}");
  }
  return sheet;
}

// Handles all requests via GET.
// ?action=save&d=<urlencoded json>  → saves data
// no action param                   → returns current data
function doGet(e) {
  const sheet = getDataSheet_();

  if (e.parameter && e.parameter.action === "save") {
    try {
      const body = JSON.parse(decodeURIComponent(e.parameter.d));
      sheet.getRange("A1").setValue(JSON.stringify(body));
      return ContentService
        .createTextOutput(JSON.stringify({ ok: true }))
        .setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService
        .createTextOutput(JSON.stringify({ ok: false, error: err.toString() }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  // Default: load
  const raw = sheet.getRange("A1").getValue();
  const json = raw && raw.toString().trim() ? raw.toString() : "{}";
  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

// Keep doPost as a no-op fallback so old requests don't error
function doPost(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: false, error: "Use GET with action=save" }))
    .setMimeType(ContentService.MimeType.JSON);
}
