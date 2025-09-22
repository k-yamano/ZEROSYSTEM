/**
 * @fileoverview ユーザーの操作ログとシステムの動作ログを記録します。
 * 「いつ、誰が、何をしたか」を追跡可能にし、監査証跡として利用します。
 */

const LOG_SHEET_NAME = 'ActionLog'; // ログを記録するシート名

/**
 * ユーザーのアクションをスプレッドシートに記録します。
 * @param {string} email - 操作を行ったユーザーのメールアドレス。
 * @param {string} action - 'submitImprovementReport' などの操作名。
 * @param {object} details - { id: 'AB-C12', ... } のような操作の詳細情報。
 */
function logAction(email, action, details) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOG_SHEET_NAME);
    if (!sheet) {
      // ログシートが存在しない場合は作成
      const newSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(LOG_SHEET_NAME);
      newSheet.appendRow(['タイムスタンプ', 'ユーザー', 'アクション', '詳細']);
    }

    const timestamp = new Date();
    const detailsString = JSON.stringify(details);
    sheet.appendRow([timestamp, email, action, detailsString]);

  } catch (e) {
    Logger.log(`アクションログの記録に失敗しました: ${e.message}`);
    // ここでErrorHandlerを呼ぶと無限ループの可能性があるので、Logger.logに留める
  }
}
