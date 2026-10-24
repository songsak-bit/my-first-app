// Google Apps Script: ลงชื่อด้วยอีเมล + ดูประวัติของตัวเอง
//
// แท็บ "data" คอลัมน์ A = timestamp, B = name, C = note, D = email (แถวที่ 1 เป็นหัวตาราง)
// วางโค้ดนี้ใน Extensions > Apps Script ของชีต แล้ว Deploy เป็น Web app
//   Execute as: Me   /   Who has access: Anyone
//
// วิธีล็อกอิน: ผู้ใช้ใส่อีเมล > สคริปต์ส่งรหัส 6 หลักไปที่อีเมลนั้น > ใส่รหัสถูก > ได้ session
// รหัสและ session เก็บเป็นค่าแฮชใน CacheService เท่านั้น และหมดอายุเอง

/* ---------- ตั้งค่า ---------- */

const SHEET_NAME = "data";
const APP_NAME = "แผ่นลงชื่อออนไลน์";

// จำกัดโดเมนอีเมลที่ล็อกอินได้ เช่น ["coolinnotech.com"]  เว้นว่าง = อีเมลอะไรก็ได้
const ALLOWED_DOMAINS = [];

const CODE_TTL = 10 * 60;        // รหัสใช้ได้ 10 นาที
const CODE_MAX_TRIES = 5;        // ใส่รหัสผิดได้ 5 ครั้งต่อรหัส
const RESEND_WAIT = 60;          // ขอรหัสใหม่ได้ทุก 60 วินาที
const MAX_CODES_PER_HOUR = 5;    // ขอรหัสได้ไม่เกิน 5 ครั้งต่อชั่วโมงต่ออีเมล
const SESSION_TTL = 6 * 60 * 60; // อยู่ในระบบได้ 6 ชั่วโมง (สูงสุดของ CacheService)
const HISTORY_LIMIT = 200;

/* ---------- เครื่องมือ ---------- */

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(["timestamp", "name", "note", "email"]);
  } else if (String(sh.getRange(1, 4).getValue()).trim() !== "email") {
    sh.getRange(1, 4).setValue("email");
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

function sha_(text) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, text, Utilities.Charset.UTF_8);
  return bytes.map(b => ((b + 256) % 256).toString(16).padStart(2, "0")).join("");
}

// เทียบทั้งสตริงเสมอ ไม่หยุดกลางทาง
function sameText_(a, b) {
  if (a.length !== b.length) return false;
  let diff = 0;
  for (let i = 0; i < a.length; i++) diff |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return diff === 0;
}

function normEmail_(s) {
  return String(s || "").trim().toLowerCase();
}

function emailOk_(email) {
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email) || email.length > 254) return false;
  if (!ALLOWED_DOMAINS.length) return true;
  const domain = email.split("@")[1];
  return ALLOWED_DOMAINS.some(d => domain === d.toLowerCase());
}

function randomCode_() {
  // ตัวเลข 6 หลักจาก UUID แบบสุ่ม
  const hex = Utilities.getUuid().replace(/-/g, "").slice(0, 12);
  return String(parseInt(hex, 16) % 1000000).padStart(6, "0");
}

function cache_() { return CacheService.getScriptCache(); }

function getJson_(key) {
  const v = cache_().get(key);
  return v ? JSON.parse(v) : null;
}

function putJson_(key, obj, ttl) {
  cache_().put(key, JSON.stringify(obj), ttl);
}

/* ---------- ล็อกอิน ---------- */

function requestCode_(body) {
  const email = normEmail_(body.email);
  if (!emailOk_(email)) return { ok: false, error: ALLOWED_DOMAINS.length ? "email_not_allowed" : "bad_email" };

  const id = sha_("email:" + email);
  const now = Date.now();
  const limit = getJson_("rate:" + id) || { sent: [] };
  limit.sent = limit.sent.filter(t => now - t < 3600 * 1000);
  if (limit.sent.length && now - limit.sent[limit.sent.length - 1] < RESEND_WAIT * 1000) return { ok: false, error: "wait" };
  if (limit.sent.length >= MAX_CODES_PER_HOUR) return { ok: false, error: "too_many" };
  if (MailApp.getRemainingDailyQuota() < 1) return { ok: false, error: "mail_quota" };

  const code = randomCode_();
  const salt = Utilities.getUuid();
  putJson_("otp:" + id, { salt: salt, hash: sha_(salt + ":" + code), tries: 0 }, CODE_TTL);
  limit.sent.push(now);
  putJson_("rate:" + id, limit, 3600);

  MailApp.sendEmail({
    to: email,
    subject: "รหัสเข้าสู่ระบบ " + APP_NAME + ": " + code,
    body: "รหัสเข้าสู่ระบบของคุณคือ " + code + "\n\nรหัสนี้ใช้ได้ " + (CODE_TTL / 60) +
      " นาที ถ้าคุณไม่ได้ขอรหัสนี้ ไม่ต้องทำอะไร",
    name: APP_NAME,
  });
  return { ok: true };
}

function verifyCode_(body) {
  const email = normEmail_(body.email);
  const code = String(body.code || "").trim();
  const id = sha_("email:" + email);
  const key = "otp:" + id;
  const otp = getJson_(key);
  if (!otp) return { ok: false, error: "code_expired" };
  if (otp.tries >= CODE_MAX_TRIES) { cache_().remove(key); return { ok: false, error: "code_expired" }; }

  if (!/^\d{6}$/.test(code) || !sameText_(sha_(otp.salt + ":" + code), otp.hash)) {
    otp.tries += 1;
    putJson_(key, otp, CODE_TTL);
    return { ok: false, error: "bad_code", left: CODE_MAX_TRIES - otp.tries };
  }

  cache_().remove(key);
  const token = Utilities.getUuid() + Utilities.getUuid();
  putJson_("sess:" + sha_(token), { email: email }, SESSION_TTL);
  return { ok: true, token: token, email: email };
}

function sessionEmail_(token) {
  if (!token) return null;
  const s = getJson_("sess:" + sha_(String(token)));
  return s ? s.email : null;
}

function logout_(body) {
  if (body.token) cache_().remove("sess:" + sha_(String(body.token)));
  return { ok: true };
}

/* ---------- ลงชื่อ + ประวัติ ---------- */

function sign_(body, email) {
  const name = String(body.name || "").trim().slice(0, 80);
  if (!name) return { ok: false, error: "name required" };
  const note = String(body.note || "").trim().slice(0, 300);
  getSheet_().appendRow([new Date(), safe_(name), safe_(note), email]);
  return { ok: true };
}

function history_(email) {
  const sh = getSheet_();
  const last = sh.getLastRow();
  if (last < 2) return { ok: true, email: email, rows: [] };
  const values = sh.getRange(2, 1, last - 1, 4).getValues();
  const rows = [];
  for (let i = values.length - 1; i >= 0 && rows.length < HISTORY_LIMIT; i--) {
    const r = values[i];
    if (normEmail_(r[3]) !== email) continue;
    const ts = r[0] instanceof Date ? r[0].getTime() : Date.parse(r[0]) || 0;
    rows.push({ id: "r" + (i + 2), name: String(r[1] || ""), note: String(r[2] || ""), at: ts });
  }
  return { ok: true, email: email, rows: rows };
}

/* ---------- Web app ---------- */

function doGet() {
  // ไม่เปิดเผยรายชื่อผ่าน GET
  return json_({ ok: true, app: APP_NAME });
}

function doPost(e) {
  let body;
  try { body = JSON.parse((e.postData && e.postData.contents) || "{}"); }
  catch (err) { return json_({ ok: false, error: "bad_request" }); }

  const action = body.action;
  if (action === "request_code") return json_(withLock_(() => requestCode_(body)));
  if (action === "verify_code") return json_(withLock_(() => verifyCode_(body)));
  if (action === "logout") return json_(logout_(body));

  const email = sessionEmail_(body.token);
  if (!email) return json_({ ok: false, error: "no_session" });
  if (action === "sign") return json_(withLock_(() => sign_(body, email)));
  if (action === "history") return json_(history_(email));
  return json_({ ok: false, error: "unknown action" });
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try { return fn(); } finally { lock.releaseLock(); }
}

// รันฟังก์ชันนี้ 1 ครั้งจากหน้า Apps Script เพื่ออนุญาตสิทธิ์ส่งอีเมลและแก้ชีต
function authorize() {
  getSheet_();
  Logger.log("ส่งอีเมลได้อีก " + MailApp.getRemainingDailyQuota() + " ฉบับวันนี้");
}
