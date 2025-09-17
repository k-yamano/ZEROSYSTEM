/** 報告者へ受付完了を通知する */
function notifySubmitter(email, id, subject) {
  const title = `[リスクアセスメント] 受付完了 (ID: ${id})`;
  const body = `ご報告ありがとうございます。\n\nID: ${id}\n件名: ${subject}\n\n▼状況確認\n${WEBAPP_BASE_URL}?id=${id}`;
  MailApp.sendEmail(email, title, body);
}

/** AI評価完了を評価者へ通知する */
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

/** 差し戻しを通知する */
function notifyRevert(incident, targetStatus, reason) {
  let email = '';
  let msg = '';
  if (targetStatus === '改善報告中') {
    email = incident.reporter;
    msg = `ID: ${incident.unique_id} の改善報告が差し戻されました。`;
  } else if (targetStatus === 'リスク評価中') {
    email = incident.hopeful_evaluator;
    msg = `ID: ${incident.unique_id} のリスク評価が差し戻されました。`;
  }
  if (!email) return;
  const title = `[差し戻し] 対応のお願い (ID: ${incident.unique_id})`;
  const body = `${msg}\n\n[理由]\n${reason}\n\n▼対応はこちら\n${WEBAPP_BASE_URL}?id=${incident.unique_id}`;
  MailApp.sendEmail(email, title, body);
}

/** Chatに通知を送信する */
function sendChat(webhook, message, id) {
  const payload = { "text": `${message}\n${WEBAPP_BASE_URL}?id=${id}` };
  const options = { method: 'post', contentType: 'application/json; charset=UTF-8', payload: JSON.stringify(payload) };
  try { UrlFetchApp.fetch(webhook, options); } catch (e) { Logger.log(`Chat notify failed: ${e.message}`); }
}

/** 部署のWebhook URLを取得する */
function getWebhook(department) {
  try {
    const hooks = JSON.parse(SCRIPT_PROPS.getProperty('DEPT_WEBHOOKS_JSON') || '{}');
    return hooks[department] || hooks['default'] || null;
  } catch (e) { return null; }
}