/**
 * 課程進度同步 - Apps Script Web App
 * 搭配網頁: https://jackytslin8.github.io/gemini-course-chedule/
 *
 * 部署方式（一次性，約 2 分鐘）:
 *   1. 開啟 Google Sheet「課程進度同步 (course-schedule)」:
 *      https://docs.google.com/spreadsheets/d/1mfnpKtXyOgJkoWg5GxQq6FXo83I2kltuYe9Uwg_46_4/edit
 *   2. 選單「Extensions(功能)」→「Apps Script」→ 全選刪除原內容 → 貼上本程式碼 → 上方「Save(儲存)」
 *   3. 右側「Deploy(部署)」→「New deployment(新部署)」→ 齿轮選「Web app(網頁應用程式)」
 *   4. 執行身分: 選「Me(我)」/ 誰可以存取: 選「Anyone(任何人)」→ 按下 Deploy
 *   5. 勾選同意授權 → 完成後把頁面上顯示的
 *      https://script.google.com/macros/s/XXXX/exec 網址貼回給 Agnes
 *
 * 之後網頁開頁面會自動抓資料、改資料會自動上傳，不用再動。
 */
var SECRET = "j5zmhA9tnlhWsCPT9By63g";
var CHUNK = 49000; // Google 儲存格上限 50,000 字，留餘裕

function doGet() {
  var ss = SpreadsheetApp.getActive();
  var data = readChunks(ss);
  var ts = ss.getRange("sync!D2").getValue() || "";
  return jsonOut({ ok: true, data: data, ts: ts });
}

function doPost(e) {
  var body = null;
  try {
    body = JSON.parse((e.postData && e.postData.data) || "{}");
  } catch (err) {
    body = null;
  }
  if (!body || body.token !== SECRET) {
    return jsonOut({ ok: false, error: "bad token" });
  }
  var data = body.data;
  if (!data || typeof data !== "object" || Array.isArray(data)) {
    return jsonOut({ ok: false, error: "bad data" });
  }
  var ss = SpreadsheetApp.getActive();
  var now = new Date().toISOString();
  var jsonStr = JSON.stringify(data);
  writeChunks(ss, jsonStr);
  ss.getRange("sync!D2").setValue(now);
  rebuildRecords(ss, data);
  return jsonOut({ ok: true, ts: now, size: jsonStr.length });
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ---- sync tab: JSON 分塊存在 B 欄 (B2 起往下)，D2 = 最後同步時間 ----
function writeChunks(ss, jsonStr) {
  var sh = ss.getSheetByName("sync");
  var lastRow = sh.getLastRow();
  if (lastRow >= 2) {
    sh.getRange(2, 2, lastRow - 1, 1).clearContent();
  }
  if (!jsonStr) return;
  var chunks = [];
  for (var i = 0; i < jsonStr.length; i += CHUNK) {
    chunks.push([jsonStr.substring(i, i + CHUNK)]);
  }
  sh.getRange(2, 2, chunks.length, 1).setValues(chunks);
}

function readChunks(ss) {
  var sh = ss.getSheetByName("sync");
  var lastRow = sh.getLastRow();
  var parts = [];
  for (var r = 2; r <= lastRow; r++) {
    var v = sh.getRange(r, 2).getValue();
    if (v === "") break;
    parts.push(String(v));
  }
  var joined = parts.join("");
  if (!joined) return null;
  try {
    return JSON.parse(joined);
  } catch (err) {
    return null;
  }
}

// ---- Records tab: 人類可讀鏡像 (日期/班級/進度) ----
// 相容兩種欄位: classId/text (gemini 版) 或 code/content (old 版)
function rebuildRecords(ss, data) {
  var sh = ss.getSheetByName("Records");
  sh.clearContents();
  var rows = [["日期", "班級", "進度"]];
  var keys = Object.keys(data).sort();
  for (var i = 0; i < keys.length; i++) {
    var entries = data[keys[i]] || [];
    if (entries.length === 0) {
      rows.push([keys[i], "", ""]);
      continue;
    }
    for (var j = 0; j < entries.length; j++) {
      var it = entries[j];
      rows.push([
        keys[i],
        String(it.classId !== undefined ? it.classId : (it.code || "")),
        String(it.text !== undefined ? it.text : (it.content || ""))
      ]);
    }
  }
  sh.getRange(1, 1, rows.length, 3).setValues(rows);
}
