// Google Apps Script สำหรับเก็บรายชื่อใน Google Sheets
// ใช้แท็บ "data" คอลัมน์ A = timestamp, B = name, C = note (แถวที่ 1 เป็นหัวตาราง)
// วางโค้ดนี้ใน Extensions > Apps Script ของชีต แล้ว Deploy เป็น Web app

const SHEET_NAME = "data";

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(["timestamp", "name", "note"]);
  }
  return sh;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// กันข้อความที่ขึ้นต้นด้วย = + - @ ไม่ให้ชีตตีความเป็นสูตร
function safe_(s) {
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function doGet() {
  const values = getSheet_().getDataRange().getValues().slice(1);
  const rows = [];
  values.forEach((r, i) => {
    const name = String(r[1] || "").trim();
    if (!name) return;
    const ts = r[0] instanceof Date ? r[0].getTime() : Date.parse(r[0]) || 0;
    rows.push({ id: "r" + (i + 2) + "-" + ts, name: name, note: String(r[2] || ""), at: ts });
  });
  return json_({ ok: true, rows: rows });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const body = JSON.parse((e.postData && e.postData.contents) || "{}");
    if (body.action !== "add") return json_({ ok: false, error: "unknown action" });
    const name = String(body.name || "").trim().slice(0, 80);
    if (!name) return json_({ ok: false, error: "name required" });
    const note = String(body.note || "").trim().slice(0, 300);
    getSheet_().appendRow([new Date(), safe_(name), safe_(note)]);
    return json_({ ok: true });
  } finally {
    lock.releaseLock();
  }
}
