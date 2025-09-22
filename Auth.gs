/**
 * @fileoverview ユーザー認証と役割ベースの権限管理を担当します。
 * このファイルで、誰がどの操作を実行できるかを集中管理します。
 */

/**
 * ユーザーの役割（ロール）を取得します。
 * @param {string} email - ユーザーのメールアドレス。
 * @returns {string} - '管理者', '担当者', '閲覧者' などの役割名。
 * @description 現状はダミー実装です。将来的にはスプレッドシートや別システムで管理する役割リストを参照します。
 */
function getUserRole(email) {
  // TODO: 実際の役割管理システムと連携する
  const adminUsers = ['your-admin-email@example.com']; // 管理者リスト
  if (adminUsers.includes(email)) {
    return '管理者';
  }
  return '担当者'; // デフォルト
}

/**
 * 指定されたアクションを実行する権限があるかどうかをチェックします。
 * @param {string} email - ユーザーのメールアドレス。
 * @param {string} action - 'submitFinalEvaluation', 'revert' などの操作名。
 * @returns {boolean} - 権限があれば true。
 * @description 役割ごとに許可されるアクションを定義します。
 */
function hasPermission(email, action) {
  const role = getUserRole(email);
  const permissions = {
    '管理者': ['submitSecondaryEvaluation', 'updateScores', 'updatePostImproveScores', 'submitFinalEvaluation', 'revert'],
    '担当者': ['submitImprovementReport'],
    '閲覧者': []
  };

  if (!permissions[role]) {
    return false;
  }
  return permissions[role].includes(action);
}
