// Google Apps Script สำหรับเก็บรายชื่อใน Google Sheets
// ใช้แท็บ "data" คอลัมน์ A = timestamp, B = name, C = note (แถวที่ 1 เป็นหัวตาราง)
// วางโค้ดนี้ใน Extensions > Apps Script ของชีต แล้ว Deploy เป็น Web app
//
// PIN: ตั้งจากเมนู "ลงชื่อ > ตั้ง PIN" ในชีต
// สคริปต์เก็บแค่ค่าแฮช (SHA-256 + salt) ไว้ใน Script Properties ไม่เก็บ PIN จริง

const SHEET_NAME = "data";
const HASH_ROUNDS = 5000;       // แฮชซ้ำหลายรอบ ให้เดา PIN จากค่าแฮชได้ช้าลง
const MAX_FAILS = 10;           // ใส่ PIN ผิดได้กี่ครั้ง
const LOCK_SECONDS = 10 * 60;   // ก่อนถูกล็อกนานเท่านี้

/* ---------- ชีต ---------- */

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

/* ---------- PIN ---------- */

function hashPin_(pin, salt) {
  let bytes = Utilities.newBlob(salt + ":" + pin).getBytes();
  for (let i = 0; i < HASH_ROUNDS; i++) {
    bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, bytes);
  }
  return Utilities.base64Encode(bytes);
}

// เทียบทีละตัวทั้งสตริง ไม่หยุดกลางทางเมื่อเจอตัวที่ต่าง
function sameText_(a, b) {
  if (a.length !== b.length) return false;
  let diff = 0;
  for (let i = 0; i < a.length; i++) diff |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return diff === 0;
}

function checkPin_(pin) {
  const props = PropertiesService.getScriptProperties();
  const salt = props.getProperty("PIN_SALT");
  const hash = props.getProperty("PIN_HASH");
  if (!salt || !hash) return "pin_not_set";

  const cache = CacheService.getScriptCache();
  const fails = Number(cache.get("pin_fails") || 0);
  if (fails >= MAX_FAILS) return "locked";

  if (sameText_(hashPin_(String(pin || ""), salt), hash)) return "ok";
  cache.put("pin_fails", String(fails + 1), LOCK_SECONDS);
  return "bad_pin";
}

// เมนูในชีตสำหรับตั้ง / เปลี่ยน PIN (PIN ไม่ถูกเขียนลงโค้ดหรือชีต)
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("ลงชื่อ")
    .addItem("ตั้ง PIN", "setPinFromMenu")
    .addToUi();
}

function setPinFromMenu() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt("ตั้ง PIN", "พิมพ์ PIN ใหม่ (ตัวเลข 4–8 หลัก)", ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  const pin = res.getResponseText().trim();
  if (!/^\d{4,8}$/.test(pin)) {
    ui.alert("PIN ต้องเป็นตัวเลข 4–8 หลัก ยังไม่ได้เปลี่ยน PIN");
    return;
  }
  const salt = Utilities.getUuid();
  PropertiesService.getScriptProperties().setProperties({
    PIN_SALT: salt,
    PIN_HASH: hashPin_(pin, salt),
  });
  CacheService.getScriptCache().remove("pin_fails");
  ui.alert("ตั้ง PIN เรียบร้อย");
}

/* ---------- Web app ---------- */

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

    const pinResult = checkPin_(body.pin);
    if (pinResult !== "ok") return json_({ ok: false, error: pinResult });

    const name = String(body.name || "").trim().slice(0, 80);
    if (!name) return json_({ ok: false, error: "name required" });
    const note = String(body.note || "").trim().slice(0, 300);
    getSheet_().appendRow([new Date(), safe_(name), safe_(note)]);
    return json_({ ok: true });
  } finally {
    lock.releaseLock();
  }
}
