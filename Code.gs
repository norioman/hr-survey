const SPREADSHEET_ID = "★ここにスプレッドシートのIDを入れる★";

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.tryLock(10000);
  try {
    const data = JSON.parse(e.postData.contents);
    const sheetName = data.sheet || "votes_1";
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName(sheetName);

    // シートがなければ作成
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      sheet.appendRow(["timestamp", "client_id", "votes_json"]);
    } else if (sheet.getLastRow() === 0) {
      sheet.appendRow(["timestamp", "client_id", "votes_json"]);
    }

    // 同じclient_idの行を上書き
    const lastRow = sheet.getLastRow();
    let updated = false;
    if (lastRow > 1) {
      const ids = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
      for (let i = 0; i < ids.length; i++) {
        if (ids[i][0] === data.clientId) {
          sheet.getRange(i + 2, 1, 1, 3).setValues([
            [new Date().toISOString(), data.clientId, JSON.stringify(data.votes)]
          ]);
          updated = true;
          break;
        }
      }
    }
    if (!updated) {
      sheet.appendRow([new Date().toISOString(), data.clientId, JSON.stringify(data.votes)]);
    }

    return ContentService
      .createTextOutput(JSON.stringify({ status: "ok", updated: updated }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "error", message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const action = e.parameter.action || "results";

  // シート一覧を返す
  if (action === "sheets") {
    const sheets = ss.getSheets()
      .map(s => s.getName())
      .filter(n => n.startsWith("votes_"))
      .sort();
    return ContentService
      .createTextOutput(JSON.stringify({ status: "ok", sheets: sheets }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // 指定シートの集計を返す
  const sheetName = e.parameter.sheet || "votes_1";
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() <= 1) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "ok", count: 0, totals: {} }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  const lastRow = sheet.getLastRow();
  const rows = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
  const totals = {};

  rows.forEach(row => {
    try {
      const votes = JSON.parse(row[2]);
      Object.entries(votes).forEach(([key, val]) => {
        if (!val) return;
        if (!totals[key]) totals[key] = { for: 0, neutral: 0, against: 0 };
        totals[key][val]++;
      });
    } catch(e) {}
  });

  return ContentService
    .createTextOutput(JSON.stringify({ status: "ok", count: rows.length, totals: totals }))
    .setMimeType(ContentService.MimeType.JSON);
}

// 指定シートをリセット（管理者が手動実行）
function resetVotes() {
  const sheetName = "votes_1"; // 対象シートを変更して実行
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ss.insertSheet(sheetName);
  sheet.clearContents();
  sheet.appendRow(["timestamp", "client_id", "votes_json"]);
}
