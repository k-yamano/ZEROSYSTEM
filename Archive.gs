/**
 * @fileoverview 完了した案件を自動的にアーカイブします。
 */

/**
 * 'Input'シートでステータスが「完了」になっている行を'Archive'シートに移動し、
 * 元のシートからは削除します。
 * この関数をトリガーで定期的に実行することを想定しています。
 */
function archiveCompletedIncidents() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName(SHEET_NAME);
  let archiveSheet = ss.getSheetByName('Archive');

  // Archiveシートが存在しない場合は作成
  if (!archiveSheet) {
    archiveSheet = ss.insertSheet('Archive');
    // ヘッダーをコピー
    const headers = sourceSheet.getRange(1, 1, 1, sourceSheet.getLastColumn());
    headers.copyTo(archiveSheet.getRange(1, 1));
  }

  const headers = sourceSheet.getRange(1, 1, 1, sourceSheet.getLastColumn()).getValues()[0];
  const statusColIndex = headers.indexOf('ステータス');
  if (statusColIndex === -1) {
    Logger.log('ステータス列が見つかりません。');
    return;
  }

  const data = sourceSheet.getDataRange().getValues();
  const rowsToMove = [];
  const rowsToDelete = [];

  // ヘッダー行を除いてループ
  for (let i = 1; i < data.length; i++) {
    if (data[i][statusColIndex] === '完了') {
      rowsToMove.push(data[i]);
      rowsToDelete.push(i + 1); // 1-based index
    }
  }

  if (rowsToMove.length > 0) {
    // Archiveシートに行を追加
    archiveSheet.getRange(archiveSheet.getLastRow() + 1, 1, rowsToMove.length, rowsToMove[0].length).setValues(rowsToMove);
    
    // 元のシートから行を削除（下から削除するのが安全）
    for (let i = rowsToDelete.length - 1; i >= 0; i--) {
      sourceSheet.deleteRow(rowsToDelete[i]);
    }
    
    Logger.log(`${rowsToMove.length} 件の完了案件をアーカイブしました。`);
  } else {
    Logger.log('アーカイブ対象の案件はありませんでした。');
  }
  
  // キャッシュをクリアして行番号の不整合を防ぐ
  clearIdCache();
}