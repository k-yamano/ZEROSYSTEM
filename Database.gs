/**
 * Finds a row in the sheet by its ID.
 * @private
 * @param {string} id The unique ID of the row to find.
 * @returns {object|null} An object containing sheet, headers, rowNum, and values, or null if not found.
 */
function findRowById_(id) {
  const rowNum = getRowNumberById_(id);
  if (!rowNum) return null;

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const rowValues = sheet.getRange(rowNum, 1, 1, headers.length).getValues()[0];
  const values = {};
  headers.forEach((h, i) => { if (h) values[h] = rowValues[i]; });
  return { sheet, headers, rowNum, values };
}

/**
 * Gets the row number for a given ID, using a cache.
 * @private
 * @param {string} id The ID to search for.
 * @returns {number|null} The row number, or null if not found.
 */
function getRowNumberById_(id) {
  const cache = CacheService.getScriptCache();
  const CACHE_KEY = 'id_row_map';
  let idRowMap = JSON.parse(cache.get(CACHE_KEY) || '{}');
  if (idRowMap[id]) {
    return idRowMap[id];
  }
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const idCol = headers.indexOf('ID');
  if (idCol === -1) return null;
  const idList = sheet.getRange(2, idCol + 1, sheet.getLastRow() - 1, 1).getValues();
  idRowMap = {};
  idList.forEach((rowId, index) => {
    if (rowId[0]) {
      idRowMap[String(rowId[0]).trim()] = index + 2;
    }
  });
  cache.put(CACHE_KEY, JSON.stringify(idRowMap), 21600); // Cache for 6 hours
  return idRowMap[String(id).trim()] || null;
}

/**
 * Clears the ID-to-row number cache.
 */
function clearIdCache() {
  CacheService.getScriptCache().remove('id_row_map');
}

/**
 * Gets data for a given ID and maps header names to friendly keys.
 * @param {string} id The unique ID.
 * @returns {object} The data object for the row.
 */
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

/**
 * Updates a row in the spreadsheet by its ID.
 * @param {string} id The ID of the row to update.
 * @param {object} data The data to update, with keys as header names.
 * @param {number|null} rowNum Optional row number to avoid re-searching.
 */
function updateRowById(id, data, rowNum = null) {
  const info = rowNum ?
    { sheet: SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME), headers: SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME).getRange(1, 1, 1, SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME).getLastColumn()).getValues()[0], rowNum } : findRowById_(id);
  if (!info) throw new Error(`ID: ${id} が見つかりません。`);
  
  const { sheet, headers, rowNum: foundRowNum } = info;
  const targetRow = rowNum || foundRowNum;
  const range = sheet.getRange(targetRow, 1, 1, headers.length);
  const currentValues = range.getValues()[0];
  headers.forEach((h, i) => { if (data.hasOwnProperty(h)) currentValues[i] = data[h]; });
  range.setValues([currentValues]);
}

/**
 * Logs updates to the Log sheet.
 * @param {string} id The incident ID.
 * @param {string} item The item that was changed.
 * @param {*} before The value before the change.
 * @param {*} after The value after the change.
 * @param {string} evaluator The email of the person who made the change.
 */
function logUpdate(id, item, before, after, evaluator) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOG_SHEET);
    if (!sheet) {
      Logger.log(`シート「${LOG_SHEET}」が見つかりません。`);
      return;
    }
    sheet.appendRow([new Date(), id, item, String(before), String(after), evaluator]);
  } catch (e) {
    Logger.log(`ログの記録に失敗しました: ${e.message}`);
  }
}

