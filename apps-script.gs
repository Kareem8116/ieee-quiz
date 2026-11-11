// =====================================================
// IEEE KSB Quiz – Google Apps Script Web App
// ✅ Fixed: LockService + doGet for Admin Dashboard
// =====================================================

const SHEET_NAME  = "Quiz Results";
const ADMIN_KEY   = "IEEE_KSB_ADMIN_2024";

/* ── doGet: Admin Dashboard reads from Sheet ── */
function doGet(e) {
  if (!e.parameter || e.parameter.key !== ADMIN_KEY) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "unauthorized" }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet || sheet.getLastRow() <= 1) {
      return ContentService
        .createTextOutput(JSON.stringify({ status: "ok", data: [] }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const rows = sheet.getDataRange().getValues();
    const data = rows.slice(1).map(r => ({
      num:         r[0],
      name:        r[1],
      score:       r[2],
      total:       r[3],
      pct:         String(r[4]).replace('%',''),
      pass:        String(r[5]).includes('Passed'),
      submittedAt: r[6],
      timeTaken:   r[7]
    }));

    return ContentService
      .createTextOutput(JSON.stringify({ status: "ok", data }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/* ── doPost: Student submission with LockService ── */
function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
  } catch (lockErr) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "lock_timeout", message: "Server busy, please retry." }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  try {
    const data  = JSON.parse(e.postData.contents);
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    let sheet   = ss.getSheetByName(SHEET_NAME);

    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.appendRow(["#","Name","Score","Out of","Percentage (%)","Status","Submitted At","Time Taken"]);
      const header = sheet.getRange(1, 1, 1, 8);
      header.setBackground("#003865");
      header.setFontColor("#ffffff");
      header.setFontWeight("bold");
      sheet.setFrozenRows(1);
      sheet.setColumnWidth(1, 45);
      sheet.setColumnWidth(2, 220);
      sheet.setColumnWidth(3, 70);
      sheet.setColumnWidth(4, 65);
      sheet.setColumnWidth(5, 120);
      sheet.setColumnWidth(6, 100);
      sheet.setColumnWidth(7, 190);
      sheet.setColumnWidth(8, 110);
    }

    const lastRow = sheet.getLastRow();
    const rowNum  = lastRow;

    sheet.appendRow([
      rowNum,
      data.name,
      data.score,
      data.total,
      data.pct + "%",
      data.pass ? "Passed ✅" : "Failed ❌",
      data.submittedAt,
      data.timeTaken
    ]);

    const newRow = sheet.getRange(lastRow + 1, 1, 1, 8);
    newRow.setBackground(data.pass ? "#f0fdf4" : "#fff1f2");
    SpreadsheetApp.flush();

    return ContentService
      .createTextOutput(JSON.stringify({ status: "ok", row: lastRow + 1 }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

function testSetup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log("Connected to: " + ss.getName());
  Logger.log("Sheets: " + ss.getSheets().map(s => s.getName()).join(", "));
}
