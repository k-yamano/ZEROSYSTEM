/**
 * フォーム提出者に受付完了メールを送信します。
 */
function notifySubmitter(email, id, subject) {
  const title = `[リスクアセスメント] 受付完了 (ID: ${id})`;
  const body = `ご報告ありがとうございます。\n\nID: ${id}\n件名: ${subject}\nステータス: リスク評価中\n\nAIによる初期評価が完了し次第、担当者へ評価依頼が通知されます。`;
  MailApp.sendEmail(email, title, body);
}

/**
 * AI評価完了を評価者へメールとChatで通知します。
 */
function notifyEvaluator(evaluator, department, id, subject) {
  const title = `[要評価] AI評価完了 (ID: ${id})`;
  const body = `AIによる初期評価が完了しました。二次評価を行ってください。\n\nID: ${id}\n件名: ${subject}\n\n▼評価画面\n${WEBAPP_BASE_URL}?id=${id}`;
  MailApp.sendEmail(evaluator, title, body);
  
  const webhook = getWebhook(department);
  if (webhook) {
    const chatMsg = `[${id}] 新規リスクのAI評価が完了しました。\n担当者(${evaluator})は二次評価をお願いします。`;
    sendChat(webhook, chatMsg, id);
  }
}

/**
 * 評価者へ改善完了を通知します。
 */
function notifyImprovementComplete(evaluator, id, subject) {
  const title = `[要最終評価] 改善報告完了 (ID: ${id})`;
  const body = `改善報告が提出されました。内容を確認し、最終評価を行ってください。\n\nID: ${id}\n件名: ${subject}\n\n▼評価画面\n${WEBAPP_BASE_URL}?id=${id}`;
  MailApp.sendEmail(evaluator, title, body);
}

/**
 * ★追加: 最終評価完了をChat通知します。
 */
function notifyFinalEvaluationComplete(incident) {
  const department = incident.department;
  const webhook = getWebhook(department);
  if (webhook) {
    const chatMsg = `[完了] ID: ${incident.unique_id} のリスクアセスメントが完了しました。\n件名: ${incident.subject}`;
    sendChat(webhook, chatMsg, incident.unique_id);
  }
}

/**
 * 差し戻しを関係者へメールとChatで通知します。
 */
function notifyRevert(incident, targetStatus, reason) {
  let email = '';
  let msg = '';
  if (targetStatus === '改善報告中') { // '差し戻し対応中' の一つ手前
    email = incident.reporter;
    msg = `ID: ${incident.unique_id} の改善報告が差し戻されました。`;
  } else if (targetStatus === 'リスク評価中') {
    email = incident.hopeful_evaluator;
    msg = `ID: ${incident.unique_id} のリスク評価が差し戻されました。`;
  }

  if (email) {
    const title = `[差し戻し] 対応のお願い (ID: ${incident.unique_id})`;
    const body = `${msg}\n\n[理由]\n${reason}\n\n▼対応はこちら\n${WEBAPP_BASE_URL}?id=${incident.unique_id}`;
    MailApp.sendEmail(email, title, body);
  }

  const department = incident.department;
  const webhook = getWebhook(department);
  if (webhook) {
    const chatMsg = `[差し戻し] ${msg}\n担当者は内容を確認し、再対応をお願いします。\n理由: ${reason}`;
    sendChat(webhook, chatMsg, incident.unique_id);
  }
}

/**
 * Google Chatにメッセージを送信します。
 */
function sendChat(webhook, message, id) {
  const payload = { "text": `${message}\n${WEBAPP_BASE_URL}?id=${id}` };
  const options = { method: 'post', contentType: 'application/json; charset=UTF-8', payload: JSON.stringify(payload) };
  try { UrlFetchApp.fetch(webhook, options);
  } catch (e) { Logger.log(`Chat notify failed: ${e.message}`); }
}

/**
 * スクリプトプロパティから部署のWebhook URLを取得します。
 */
function getWebhook(department) {
  try {
    const hooks = JSON.parse(SCRIPT_PROPS.getProperty('DEPT_WEBHOOKS_JSON') || '{}');
    return hooks[department] || hooks['default'] || null;
  } catch (e) { return null; }
}