/**
 * @fileoverview データのバックアップを管理します。
 * 定期的にスプレッドシートのデータを保護し、データの損失リスクを低減します。
 */

/**
 * メインのスプレッドシートのコピーを作成してバックアップします。
 * この関数を時間ベースのトリガーで定期的に実行することを想定しています。
 */
function createBackup() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const backupFolderName = 'ZEROSYSTEM_Backups';
    let backupFolder = DriveApp.getFoldersByName(backupFolderName).next();

    if (!backupFolder) {
      backupFolder = DriveApp.createFolder(backupFolderName);
    }

    const timestamp = Utilities.formatDate(new Date(), 'JST', 'yyyy-MM-dd_HH-mm-ss');
    const backupFileName = `${spreadsheet.getName()}_Backup_${timestamp}`;

    spreadsheet.copy(backupFileName).moveTo(backupFolder);
    Logger.log(`バックアップを作成しました: ${backupFileName}`);

  } catch (e) {
    handleError(e, 'createBackup');
  }
}
