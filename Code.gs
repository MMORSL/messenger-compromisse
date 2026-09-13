/**
 * Провод — бэкенд на Google Apps Script (только текст).
 *
 * Хранит сообщения в Google-таблице и отдаёт их клиенту.
 * Разворачивается как веб-приложение (Deploy → New deployment → Web app),
 * доступ: "Anyone". URL вставляется в index.html в SHEET_API_URL.
 *
 * Лист "messages" создаётся автоматически. Колонки: ts | room | sender | kind | content
 */

const SHEET_NAME = 'messages';

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['ts', 'room', 'sender', 'kind', 'content']);
  }
  return sheet;
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// GET ?room=...&since=<ts>
function doGet(e) {
  const room = String(e.parameter.room || 'default');
  const since = Number(e.parameter.since || 0);
  const values = getSheet().getDataRange().getValues();

  const out = [];
  for (let i = 1; i < values.length; i++) {
    const ts = Number(values[i][0]);
    const r = String(values[i][1]);
    if (r !== room || ts <= since) continue;

    const hasKind = values[i].length >= 5;
    const content = hasKind ? String(values[i][4]) : String(values[i][3]);

    out.push({ ts: ts, sender: String(values[i][2]), content: content });
  }
  return json({ messages: out });
}

// POST {room, sender, text}
function doPost(e) {
  let body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return json({ ok: false, error: 'bad json' });
  }

  const room = String(body.room || 'default');
  const sender = String(body.sender || 'anon');
  const text = String(body.text || '').slice(0, 4000);
  if (!text) return json({ ok: false, error: 'empty' });

  const ts = Date.now();
  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    getSheet().appendRow([ts, room, sender, 'text', text]);
  } finally {
    lock.releaseLock();
  }
  return json({ ok: true, ts: ts });
}
