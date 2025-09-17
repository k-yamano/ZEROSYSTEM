/** IDをキーに行情報を取得する内部関数 */
function findRowById_(id) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const idCol = headers.indexOf('ID');
  if (idCol === -1) return null;

  const idList = sheet.getRange(2, idCol + 1, sheet.getLastRow() - 1, 1).getValues();
  const rowIndex = idList.findIndex(rowId => String(rowId[0]).trim() === String(id).trim());

  if (rowIndex === -1) return null;

  const rowNum = rowIndex + 2;
  const rowValues = sheet.getRange(rowNum, 1, 1, headers.length).getValues()[0];
  const values = {};
  headers.forEach((h, i) => { if (h) values[h] = rowValues[i]; });

  return { sheet, headers, rowNum, values };
}

/** IDをキーにデータを取得し、Webアプリで使いやすいキー名に変換して返す */
function getDataById(id) {
  if (!id) return {};
  const rowData = findRowById_(id);
  if (!rowData) return {};
  const result = {};
  for (const header in rowData.values) {
    result[KEY_MAP[header] || header] = rowData.values[header];
  }
  return result;
}

/** IDをキーに行データを更新する */
function updateRowById(id, data, rowNum = null) {
  const info = rowNum ? { sheet: SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME), headers: SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME).getRange(1, 1, 1, SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME).getLastColumn()).getValues()[0], rowNum } : findRowById_(id);
  if (!info) throw new Error(`ID: ${id} が見つかりません。`);
  
  const { sheet, headers, rowNum: foundRowNum } = info;
  const targetRow = rowNum || foundRowNum;
  const range = sheet.getRange(targetRow, 1, 1, headers.length);
  const currentValues = range.getValues()[0];
  
  headers.forEach((h, i) => { if (data.hasOwnProperty(h)) currentValues[i] = data[h]; });
  range.setValues([currentValues]);
}

/** 修正ログをLogシートに記録する */
function logUpdate(id, item, before, after, evaluator) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOG_SHEET);
    if (!sheet) {
      Logger.log(`シート「${LOG_SHEET}」が見つかりません。`);
      return;
    }
    sheet.appendRow([
      new Date(),
      id,
      item,
      before,
      after,
      evaluator
    ]);
  } catch (e) {
    Logger.log(`ログの記録に失敗しました: ${e.message}`);
  }
}