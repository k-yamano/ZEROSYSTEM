/**
 * 評価担当者へ初期メールを送信する
 */
function sendInitialEmail_(department, reporterEmail, id, subject, details) {
  if (!reporterEmail) { Logger.log('メールアドレスが空のため、メール通知をスキップしました。'); return; }
  const url = buildIncidentUrl(id);
  const mailSubject = `[リスクアセスメント] 新規報告がありました (ID: ${id})`;
  const body = `${department} の担当者様\n\n新しいインシデント報告がありました。\nAIによる一次評価が完了しましたので、内容を確認してください。\n\nID: ${id}\n件名: ${subject}\n内容:\n${details}\n\n▼評価画面はこちら\n${url}`;
  MailApp.sendEmail(reporterEmail, mailSubject, body);
}

/**
 * 改善報告後に評価者へメール通知
 */
function sendImprovementReportEmail_(evaluatorEmail, id, subject) {
  if (!evaluatorEmail) { Logger.log('評価者メールアドレスが空のため、通知をスキップしました。'); return; }
  const url = buildIncidentUrl(id);
  const mailSubject = `[要対応] 改善報告が提出されました (ID: ${id})`;
  const body = `評価担当者様\n\nID: ${id} の改善報告が提出されました。\n内容を確認し、最終評価を行ってください。\n\n件名: ${subject}\n\n▼評価画面はこちら\n${url}`;
  MailApp.sendEmail(evaluatorEmail, mailSubject, body);
}

/**
 * 差し戻し時にメール通知を送信する
 */
function sendRevertEmail_(notifyEmail, id, subject, message, reason) {
  if (!notifyEmail) { Logger.log('通知先メールアドレスが空のため、差し戻し通知をスキップしました。'); return; }
  const url = buildIncidentUrl(id);
  const mailSubject = `[差し戻し] 対応が必要です (ID: ${id})`;
  const body = `担当者様\n\n${message}\n\n件名: ${subject}\n差し戻し理由:\n${reason}\n\n▼評価画面はこちら\n${url}`;
  MailApp.sendEmail(notifyEmail, mailSubject, body);
}

/**
 * 部署名からWebhook URLを取得する
 */
function getDeptWebhookUrl_(deptName) {
  if (!deptName) return null;
  try {
    const hooks = JSON.parse(DEPT_WEBHOOKS_JSON);
    return hooks[deptName] || null;
  } catch (e) { console.error('DEPT_WEBHOOKS_JSON のパースに失敗', e); return null; }
}

/**
 * Google Chatへ通知を送信する
 */
function sendToChat_(webhookUrl, message) {
  if (!webhookUrl || !message) return;
  try {
    UrlFetchApp.fetch(webhookUrl, { method: 'post', contentType: 'application/json; charset=UTF-8', payload: JSON.stringify({ text: message }),});
  } catch (e) { console.error(`Chat通知に失敗: ${e.message}`); }
}

/**
 * スコアを更新する汎用関数
 */
function updateScores_(data, colNames, dataKeys, modifiedColName) {
  if (!data.id || data.id.includes('<') || data.id.includes('>')) throw new Error('無効なIDが指定されました。ページをリロードしてください。');
  const { id } = data;
  
  const incidentRow = findSheetRowById_(id);
  if (!incidentRow) throw new Error(`ID: ${id} が見つかりません。`);

  const sheetUpdate = {};
  sheetUpdate[colNames[0]] = data[dataKeys[0]];
  sheetUpdate[colNames[1]] = data[dataKeys[1]];
  sheetUpdate[colNames[2]] = data[dataKeys[2]];
  if (modifiedColName) { sheetUpdate[modifiedColName] = new Date(); }
  
  updateSheetRow_(id, sheetUpdate);

  return 'スコアを更新しました。';
}


/** ユニークIDを生成 */
function generateUniqueId_() {
  const rand2Letters = () => Math.random().toString(36).substring(2, 4).toUpperCase();
  const t = new Date();
  return `${rand2Letters()}-${(t.getSeconds()*1000 + t.getMilliseconds()).toString().padStart(4,'0').slice(-4)}`;
}

/** 評価画面のURLを生成 */
function buildIncidentUrl(id){ return `${WEBAPP_BASE_URL || ScriptApp.getService().getUrl()}?id=${encodeURIComponent(id)}`; }
/** レポート画面のURLを生成 */
function buildIncidentReportUrl(id){ return `${WEBAPP_BASE_URL || ScriptApp.getService().getUrl()}?id=${encodeURIComponent(id)}&view=report`; }

/** 列番号をA1形式の列文字に変換する */
function getSheetColumnName(colIndex) {
  let name = ''; let num = colIndex;
  while (num > 0) { let remainder = (num - 1) % 26; name = String.fromCharCode(65 + remainder) + name; num = Math.floor((num - 1) / 26); }
  return name;
}

/** トークンログシートが存在することを確認し、なければ作成する */
function ensureTokenSheet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sh = ss.getSheetByName(TOKEN_SHEET_NAME);
  if (!sh) { sh = ss.insertSheet(TOKEN_SHEET_NAME); sh.getRange(1,1,1,6).setValues([['日時','ID','PromptTokens','CandidateTokens','TotalTokens','関数']]); sh.setFrozenRows(1); }
  return sh;
}

/** Gemini APIのトークン使用量をスプレッドシートに記録する */
function logTokenUsage(insertId, promptTokens, candidateTokens, fnLabel) {
  try {
    const sh = ensureTokenSheet_();
    sh.appendRow([ new Date(), insertId, promptTokens, candidateTokens, (promptTokens || 0) + (candidateTokens || 0), fnLabel ]);
  } catch (e) { console.error(`Tokenログの書き込みに失敗: ${e.message}`); }
}

