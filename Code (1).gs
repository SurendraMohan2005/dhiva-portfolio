/**
 * Code.gs — Google Apps Script backend for the portfolio contact form.
 *
 * Receives name, email and phone via GET or POST query parameters,
 * appends a row (Timestamp, Name, Email, Phone) to the "Responses" tab
 * of the bound spreadsheet, and returns JSON: {"success": true|false}.
 *
 * Deploy: Extensions > Apps Script (from the target Google Sheet)
 *         > Deploy > New deployment > Type: Web app
 *         > Execute as: Me > Who has access: Anyone > Deploy.
 */

const SHEET_NAME = 'Responses';
const HEADERS = ['Timestamp', 'Name', 'Email', 'Phone'];

/** Handles GET requests (used by the HTML form via fetch with query params). */
function doGet(e) {
  return handleRequest_(e);
}

/** Handles POST requests (same behaviour as doGet). */
function doPost(e) {
  return handleRequest_(e);
}

function handleRequest_(e) {
  try {
    const params = (e && e.parameter) ? e.parameter : {};

    const name = clean_(params.name);
    const email = clean_(params.email);
    const phone = clean_(params.phone);

    if (!name || !email) {
      return jsonResponse_({ success: false, error: 'Name and email are required.' });
    }

    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      const sheet = getResponsesSheet_();
      sheet.appendRow([new Date(), name, email, phone]);
    } finally {
      lock.releaseLock();
    }

    return jsonResponse_({ success: true });
  } catch (err) {
    return jsonResponse_({ success: false, error: String(err) });
  }
}

/**
 * Returns the "Responses" sheet. Creates it, with bold headers, if missing.
 * Also writes the headers if the sheet exists but is empty.
 */
function getResponsesSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }

  if (sheet.getLastRow() === 0) {
    const headerRange = sheet.getRange(1, 1, 1, HEADERS.length);
    headerRange.setValues([HEADERS]);
    headerRange.setFontWeight('bold');
    sheet.setFrozenRows(1);
  }

  return sheet;
}

/**
 * Trims input and neutralises spreadsheet formula injection
 * (values starting with = + - @ are prefixed with an apostrophe).
 */
function clean_(value) {
  if (value === undefined || value === null) return '';
  let s = String(value).trim().slice(0, 500);
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}

function jsonResponse_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
