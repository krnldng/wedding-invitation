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
 * Xoá / sửa lời chúc trực tiếp trên sheet: sửa nội dung ô thì cache tự xoá; xoá cả dòng thì
 * bấm menu "Lời chúc" -> "Làm mới (xoá cache)" (menu hiện sau khi tải lại trang Google Sheet).
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
const CACHE_KEY = 'wishes';
// Cache is cleared on every new wish; this only bounds staleness after editing the sheet by hand.
const CACHE_SECONDS = 300;

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
  // Skip reading the sheet while the cached answer is fresh (cleared on every new wish).
  const cache = CacheService.getScriptCache();
  const hit = cache.get(CACHE_KEY);
  if (hit) {
    return ContentService.createTextOutput(hit).setMimeType(ContentService.MimeType.JSON);
  }

  const body = JSON.stringify({ ok: true, data: readWishes_() });
  try {
    cache.put(CACHE_KEY, body, CACHE_SECONDS);
  } catch (err) {
    // value over 100KB cannot be cached, just serve it.
  }
  return ContentService.createTextOutput(body).setMimeType(ContentService.MimeType.JSON);
}

function readWishes_() {
  const sheet = getSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return [];
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

  return data;
}

// Xoá cache lời chúc. Chạy tay trong Apps Script, hoặc dùng menu "Lời chúc" trên Google Sheet.
function clearCache() {
  CacheService.getScriptCache().remove(CACHE_KEY);
}

// Thêm menu "Lời chúc" khi mở Google Sheet.
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Lời chúc')
    .addItem('Làm mới (xoá cache)', 'clearCacheFromMenu')
    .addToUi();
}

function clearCacheFromMenu() {
  clearCache();
  SpreadsheetApp.getActiveSpreadsheet().toast('Đã xoá cache, thiệp sẽ hiển thị danh sách mới.');
}

// Sửa / xoá nội dung ô trong sheet thì tự xoá cache.
// (Xoá cả dòng không kích hoạt onEdit, khi đó dùng menu "Lời chúc" -> "Làm mới".)
function onEdit() {
  clearCache();
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
    CacheService.getScriptCache().remove(CACHE_KEY);
  } finally {
    lock.releaseLock();
  }

  return json_({ ok: true });
}
