/* ============================================================
   Google Apps Script backend for student_survey_google_sheets.html

   SETUP
   1. Open the target Google Sheet.
   2. Choose Extensions > Apps Script.
   3. Replace the starter code with this file.
   4. Deploy as a Web App.
   5. Paste the resulting /exec URL into GOOGLE_SCRIPT_URL
      in the HTML survey.
   ============================================================ */

/* Health-check endpoint.
   Visiting the deployed /exec URL in a browser should display:
   Survey endpoint is alive.
*/
function doGet(e) {
  return ContentService
    .createTextOutput("Survey endpoint is alive.");
}

/* Name of the worksheet tab that stores responses. */
const SHEET_NAME = "Responses";

/* Column names and output order. */
const HEADERS = [
  "response_id",
  "submitted_at",
  "participant_id",
  "course_format",
  "expectations_clear",
  "activities_support",
  "materials_navigate",
  "comfortable_help",
  "workload",
  "technology_confidence",
  "helpful_resource",
  "improvement"
];

/* Receives POST requests from the HTML survey. */
function doPost(e) {
  const lock = LockService.getScriptLock();

  try {
    /* Convert the incoming JSON text into a JavaScript object. */
    const record = JSON.parse(e.postData.contents);

    /* Use the Google Sheet to which this Apps Script project is attached. */
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

    /* Open the Responses tab, or create it the first time. */
    let sheet = spreadsheet.getSheetByName(SHEET_NAME);
    if (!sheet) {
      sheet = spreadsheet.insertSheet(SHEET_NAME);
    }

    /* Briefly lock writes so simultaneous submissions do not collide. */
    lock.waitLock(10000);

    try {
      /* Add and format the header row when the sheet is empty. */
      if (sheet.getLastRow() === 0) {
        sheet.appendRow(HEADERS);
        sheet
          .getRange(1, 1, 1, HEADERS.length)
          .setFontWeight("bold");
        sheet.setFrozenRows(1);
      }

      /* Append one survey response in the same order as HEADERS. */
      sheet.appendRow(
        HEADERS.map(header => record[header] ?? "")
      );
    } finally {
      lock.releaseLock();
    }

    /* Return a small JSON success response. */
    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    /* Release the lock if an error occurred after it was acquired. */
    try {
      lock.releaseLock();
    } catch (_) {}

    /* Return the error as JSON for diagnostics. */
    return ContentService
      .createTextOutput(JSON.stringify({
        ok: false,
        error: String(error)
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
