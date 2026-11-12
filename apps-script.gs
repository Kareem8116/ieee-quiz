// =====================================================
// IEEE KSB – Week 1 Quiz  |  Google Apps Script
// ✅ Hardcoded Spreadsheet ID – always writes to correct sheet
// =====================================================

const SHEET_NAME     = "Quiz Results";
const ADMIN_KEY      = "IEEE_KSB_ADMIN_2024";
const SPREADSHEET_ID = "1QstBc6sR3Uzs7QjqvsH2DcdN6H5wsNFXQu1yT4uxxLs"; // ← Week 1 Sheet

/* ── Helper: always returns the correct sheet ── */
function getSheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(["#","Name","Score","Out of","Percentage (%)","Status","Submitted At","Time Taken"]);
    const hdr = sheet.getRange(1,1,1,8);
    hdr.setBackground("#003865").setFontColor("#ffffff").setFontWeight("bold");
    sheet.setFrozenRows(1);
    [45,220,70,65,120,100,190,110].forEach((w,i) => sheet.setColumnWidth(i+1,w));
  }
  return sheet;
}

/* ── doGet: Admin reads data ── */
function doGet(e) {
  if (!e.parameter || e.parameter.key !== ADMIN_KEY) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "unauthorized" }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  try {
    const sheet = getSheet();
    if (sheet.getLastRow() <= 1) {
      return ContentService
        .createTextOutput(JSON.stringify({ status: "ok", data: [] }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    const rows    = sheet.getDataRange().getValues();
    const headers = rows[0].map(h => String(h).trim().toLowerCase());
    function col(name) { const i = headers.indexOf(name.toLowerCase()); return i >= 0 ? i : -1; }

    const iName  = col('name');
    const iScore = col('score');
    const iTotal = col('out of');
    const iPct   = col('percentage (%)') >= 0 ? col('percentage (%)') : col('percentage');
    const iStat  = col('status');
    const iSubAt = col('submitted at');
    const iTime  = col('time taken');

    const data = rows.slice(1).map((r, i) => {
      const pRaw = String(r[iPct] !== undefined ? r[iPct] : '').replace('%','');
      const pNum = parseFloat(pRaw);
      const pct  = isNaN(pNum) ? 0 : (pNum < 1.5 ? Math.round(pNum*100) : Math.round(pNum));
      return {
        num: i+1, name: r[iName]||'', score: r[iScore]||0, total: r[iTotal]||30,
        pct, pass: String(r[iStat]||'').includes('Passed'),
        submittedAt: r[iSubAt]||'', timeTaken: r[iTime]||''
      };
    });

    return ContentService
      .createTextOutput(JSON.stringify({ status: "ok", data }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/* ── doPost: Student submission ── */
function doPost(e) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(30000); }
  catch(lockErr) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "lock_timeout" }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  try {
    const data    = JSON.parse(e.postData.contents);
    const sheet   = getSheet();
    const lastRow = sheet.getLastRow();
    sheet.appendRow([
      lastRow, data.name, "", data.score, data.total,
      data.pct + "%",
      data.pass ? "Passed ✅" : "Failed ❌",
      data.submittedAt, data.timeTaken
    ]);
    sheet.getRange(lastRow+1,1,1,8).setBackground(data.pass ? "#f0fdf4" : "#fff1f2");
    SpreadsheetApp.flush();
    return ContentService
      .createTextOutput(JSON.stringify({ status: "ok", row: lastRow+1 }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

function testSetup() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  Logger.log("Connected to: " + ss.getName());
  Logger.log("Sheets: " + ss.getSheets().map(s => s.getName()).join(", "));
}
