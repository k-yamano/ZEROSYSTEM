/**
 * @fileoverview システム全体のエラーハンドリングを担当します。
 * 予期せぬエラーが発生した際に、詳細を記録し、管理者に通知します。
 */

/**
 * エラーを処理するメイン関数。
 * @param {Error} e - 発生したエラーオブジェクト。
 * @param {string} context - エラーが発生した関数名や状況。
 * @returns {string} - UIに表示するための整形されたエラーメッセージ。
 */
function handleError(e, context) {
  const adminEmail = Session.getEffectiveUser().getEmail();
  const timestamp = new Date().toISOString();
  const errorMessage = `
    ZEROSYSTEMでエラーが発生しました。

    日時: ${timestamp}
    コンテキスト: ${context}
    エラーメッセージ: ${e.message}
    スタックトレース:
    ${e.stack}
  `;

  Logger.log(errorMessage);

  // 管理者へメールで通知
  try {
    MailApp.sendEmail(
      adminEmail,
      `[ZEROSYSTEM エラー通知] ${context}`,
      errorMessage
    );
  } catch (mailError) {
    Logger.log(`エラー通知メールの送信に失敗しました: ${mailError.message}`);
  }

  // UIには、ユーザーに分かりやすい一般的なメッセージを返す
  return `処理中にエラーが発生しました。管理者に通知されました。(${context})`;
}
