/**
 * API lời chúc cho thiệp cưới - chạy trên Google Apps Script, lưu vào Google Sheet.
 *
 * Cài đặt:
 *  1. Tạo 1 Google Sheet mới -> menu Tiện ích mở rộng (Extensions) -> Apps Script.
 *  2. Xoá nội dung có sẵn, dán toàn bộ file này vào, bấm Lưu.
 *  3. Triển khai (Deploy) -> Tùy chọn triển khai mới (New deployment) -> loại "Ứng dụng web" (Web app):
 *       - Thực thi với tư cách (Execute as): Tôi (Me)
 *       - Người có quyền truy cập (Who has access): Bất kỳ ai (Anyone)
 *     Cấp quyền khi được hỏi, rồi copy "URL ứng dụng web" (dạng https://script.google.com/macros/s/.../exec).
 *  4. Dán URL đó vào thuộc tính data-wish-api của thẻ <body> trong index.html.
 *
 * Khi sửa code này, phải vào Deploy -> Manage deployments -> Edit -> Version: New version
 * thì URL cũ mới chạy code mới.
 *
 * API:
 *  GET  <url>  -> { ok: true, data: [{ name, message, time }] }  (mới nhất trước, chỉ các lời chúc có nội dung)
 *  POST <url>  body JSON (gửi dạng text/plain để tránh CORS preflight): { name, presence, message }
 *              -> { ok: true } | { ok: false, error }
 */

const SHEET_NAME = 'LoiChuc';
const HEADERS = ['Thời gian', 'Họ tên', 'Tham dự', 'Lời chúc'];
const PRESENCE = { 1: 'Sẽ tham dự', 2: 'Không thể tham dự', 3: 'Không chắc' };
const MAX_NAME = 50;
const MAX_MESSAGE = 1000;

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
  }
  return sheet;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// Chặn chèn công thức (=, +, -, @) vào sheet.
function safeCell_(value) {
  const s = String(value);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function doGet() {
  const sheet = getSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return json_({ ok: true, data: [] });
  }

  const rows = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
  const data = rows
    .filter((r) => String(r[3]).trim().length > 0)
    .map((r) => ({
      name: String(r[1]).replace(/^'/, ''),
      message: String(r[3]).replace(/^'/, ''),
      time: r[0] instanceof Date ? r[0].toISOString() : String(r[0]),
    }))
    .reverse();

  return json_({ ok: true, data: data });
}

function doPost(e) {
  let body;
  try {
    body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
  } catch (err) {
    return json_({ ok: false, error: 'invalid json' });
  }

  const name = String(body.name || '').trim();
  const message = String(body.message || '').trim();
  const presence = PRESENCE[Number(body.presence)] || '';

  if (name.length === 0 || name.length > MAX_NAME) {
    return json_({ ok: false, error: 'invalid name' });
  }
  if (message.length > MAX_MESSAGE) {
    return json_({ ok: false, error: 'message too long' });
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    getSheet_().appendRow([new Date(), safeCell_(name), presence, safeCell_(message)]);
  } finally {
    lock.releaseLock();
  }

  return json_({ ok: true });
}
