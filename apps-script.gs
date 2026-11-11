// =====================================================
// IEEE KSB Quiz – Google Apps Script Web App
// ✅ Fixed: LockService prevents concurrent-write data loss
// =====================================================
// الـ Sheet:
// https://docs.google.com/spreadsheets/d/1QstBc6sR3Uzs7QjqvsH2DcdN6H5wsNFXQu1yT4uxxLs/edit
//
// خطوات تحديث الـ Deploy:
//   1. افتح الـ Spreadsheet → Extensions → Apps Script
//   2. استبدل الكود القديم بالكود ده
//   3. Deploy → Manage Deployments → Edit (✏️) → New Version → Deploy
//      ⚠️ لازم تعمل New Version مش نفس الـ deployment القديم
// =====================================================

const SHEET_NAME = "Quiz Results";

function doPost(e) {
  // ── LockService: يمنع الـ race condition لما اتنين يسلموا بالظبط في نفس الوقت
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000); // استنى لحد 30 ثانية علشان يلاقي slot
  } catch (lockErr) {
    // لو مقدرش ياخد الـ lock، بيرجع error واضح
    return ContentService
      .createTextOutput(JSON.stringify({ status: "lock_timeout", message: "Server busy, please retry." }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  try {
    const data = JSON.parse(e.postData.contents);
    const ss   = SpreadsheetApp.getActiveSpreadsheet();

    let sheet = ss.getSheetByName(SHEET_NAME);
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
    const rowNum  = lastRow; // 1-based data row count

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

    // Color-code the row
    const newRow = sheet.getRange(lastRow + 1, 1, 1, 8);
    newRow.setBackground(data.pass ? "#f0fdf4" : "#fff1f2");

    // Flush to make sure write completes before lock is released
    SpreadsheetApp.flush();

    return ContentService
      .createTextOutput(JSON.stringify({ status: "ok", row: lastRow + 1 }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    // دايماً يحرر الـ lock حتى لو في error
    lock.releaseLock();
  }
}

function testSetup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log("Connected to: " + ss.getName());
  Logger.log("Sheets: " + ss.getSheets().map(s => s.getName()).join(", "));
}
