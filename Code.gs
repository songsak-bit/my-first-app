// Google Apps Script สำหรับเก็บรายชื่อใน Google Sheets
// วางโค้ดนี้ใน Extensions > Apps Script ของชีตที่ต้องการ แล้ว Deploy เป็น Web app

const SHEET_NAME = "รายชื่อ";

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(["id", "ชื่อ", "หมายเหตุ", "เวลา (ms)", "เวลาที่ลงชื่อ"]);
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
  const rows = values
    .filter(r => r[0])
    .map(r => ({ id: String(r[0]), name: String(r[1]), note: String(r[2] || ""), at: Number(r[3]) }));
  return json_({ ok: true, rows });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const body = JSON.parse(e.postData.contents || "{}");
    if (body.action !== "add") return json_({ ok: false, error: "unknown action" });
    const name = String(body.name || "").trim().slice(0, 80);
    if (!name) return json_({ ok: false, error: "name required" });
    const note = String(body.note || "").trim().slice(0, 300);
    const at = Date.now();
    const id = String(body.id || at.toString(36)).slice(0, 40);
    getSheet_().appendRow([safe_(id), safe_(name), safe_(note), at, new Date(at)]);
    return json_({ ok: true });
  } finally {
    lock.releaseLock();
  }
}
