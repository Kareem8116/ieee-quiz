// =====================================================
// IEEE KSB Quiz – Google Apps Script Web App
// =====================================================
// Instructions:
//   1. Go to: https://script.google.com
//   2. Create a new project → paste this code
//   3. Click Deploy → New Deployment → Web App
//   4. Execute as: Me | Access: Anyone
//   5. Copy the Web App URL and paste it in index.html
//      where it says: const APPS_SCRIPT_URL = "PASTE_HERE"
// =====================================================

const SHEET_NAME = "Quiz Results";

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss   = SpreadsheetApp.getActiveSpreadsheet();

    // Get or create the sheet
    let sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      // Header row
      sheet.appendRow([
        "#", "Name", "Student ID", "Score", "Out of",
        "Percentage (%)", "Status", "Submitted At", "Time Taken"
      ]);
      // Format header
      const header = sheet.getRange(1, 1, 1, 9);
      header.setBackground("#003865");
      header.setFontColor("#ffffff");
      header.setFontWeight("bold");
      sheet.setFrozenRows(1);
      // Column widths
      sheet.setColumnWidth(1, 40);
      sheet.setColumnWidth(2, 200);
      sheet.setColumnWidth(3, 130);
      sheet.setColumnWidth(4, 70);
      sheet.setColumnWidth(5, 60);
      sheet.setColumnWidth(6, 120);
      sheet.setColumnWidth(7, 90);
      sheet.setColumnWidth(8, 180);
      sheet.setColumnWidth(9, 120);
    }

    const lastRow = sheet.getLastRow();
    const rowNum  = lastRow; // row number = count of data rows

    sheet.appendRow([
      rowNum,           // #
      data.name,        // Name
      data.studentId || "—",   // Student ID
      data.score,       // Score
      data.total,       // Out of
      data.pct + "%",   // Percentage
      data.pass ? "Passed ✅" : "Failed ❌", // Status
      data.submittedAt, // Submitted At
      data.timeTaken    // Time Taken (e.g. "12:34")
    ]);

    // Color row based on pass/fail
    const newRow = sheet.getRange(lastRow + 1, 1, 1, 9);
    if (data.pass) {
      newRow.setBackground("#f0fdf4"); // light green
    } else {
      newRow.setBackground("#fff1f2"); // light red
    }

    return ContentService
      .createTextOutput(JSON.stringify({ status: "ok", row: lastRow + 1 }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Test function — run this manually to verify setup
function testSetup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log("Connected to: " + ss.getName());
  Logger.log("Sheet URL: " + ss.getUrl());
}
