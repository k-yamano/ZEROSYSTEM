/**
 * Webアプリ(UI)からの操作を処理するメインの関数
 * @param {string} action - 実行するアクションの種類
 * @param {object} data - UIから送信されたデータ
 * @returns {string} - UIに表示する成功メッセージ
 */
function serverAction(action, data) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) {
    throw new Error('他の処理が実行中のため、少し待ってから再度お試しください。');
  }
  try {
    switch (action) {
      case 'submitSecondaryEvaluation': return submitSecondaryEvaluation(data);
      case 'updateScores': return updateSecondaryScores(data);
      case 'submitImprovementReport': return submitImprovementReport(data);
      case 'updatePostImproveScores': return updatePostImproveScores(data);
      case 'submitFinalEvaluation': return submitFinalEvaluation(data);
      case 'revert': return revert(data);
      default: throw new Error('無効なアクションです。');
    }
  } catch (e) {
    Logger.log(`serverAction error (${action}): ${e.message}\n${e.stack}`);
    throw new Error(`サーバー処理エラー: ${e.message}`);
  } finally {
    lock.releaseLock();
  }
}

/**
 * 二次評価（人）の情報を更新し、ステータスを「改善報告中」へ進める
 */
function submitSecondaryEvaluation(data){
  const { id, department, hopeful_evaluator, secondary_eval_comment, deadline, provisional_budget } = data;
  
  if (!id || !department || !hopeful_evaluator || !secondary_eval_comment) {
    throw new Error('担当部署、評価者、評価者コメントは必須です。');
  }
  
  const updateData = {
    'ステータス': '改善報告中',
    '担当部署': department,
    '希望する評価者': hopeful_evaluator,
    '評価者によるコメント': secondary_eval_comment, // AC列
    '期限': deadline,
    '暫定予算': provisional_budget
  };
  updateRowById(id, updateData);
  
  // Chat通知
  try {
    const incidentData = getDataById(id);
    const message = `【ヒヤリハット 改善依頼】\nID: ${id}\n件名: ${incidentData.subject}\n内容: ${(incidentData.details || '').slice(0,140)}...\nリンク: ${WEBAPP_BASE_URL}?id=${id}`;
    notifyDepartmentChat_(department, message);
  } catch(e) {
    Logger.log('Chat通知失敗: ' + e.message);
  }
  
  return "改善指示を送信しました。";
}

/**
 * AIが評価した二次評価スコアを、評価者が修正する
 */
function updateSecondaryScores(data){
  const { id, frequency_ai, likelihood_ai, severity_ai } = data;
  if (!id) throw new Error('IDが未指定です。');
  
  const beforeData = getDataById(id);
  const evaluator = Session.getActiveUser().getEmail();

  const updateData = {
    '頻度（AI評価）': frequency_ai,
    '発生の可能性（AI評価）': likelihood_ai,
    '重篤度（AI評価）': severity_ai,
    '評価者によるリスク修正': new Date()
  };
  
  // ログ記録
  if (String(beforeData.frequency_ai)  !== String(frequency_ai))  logUpdate(id, '頻度（AI評価）', beforeData.frequency_ai, frequency_ai, evaluator);
  if (String(beforeData.likelihood_ai) !== String(likelihood_ai)) logUpdate(id, '発生の可能性（AI評価）', beforeData.likelihood_ai, likelihood_ai, evaluator);
  if (String(beforeData.severity_ai)   !== String(severity_ai))   logUpdate(id, '重篤度（AI評価）', beforeData.severity_ai, severity_ai, evaluator);
  
  updateRowById(id, updateData);
  
  return 'AI二次評価を更新しました。';
}

/**
 * 改善完了報告を提出し、ステータスを「最終評価中」へ進める
 */
function submitImprovementReport(data) {
  const { id, improvement_details, ojt_completed, ojt_id, cost, effort, effect, team_member_1, team_member_2, team_member_3, reference_url } = data;

  if (!id || !improvement_details) throw new Error('改善内容の詳細は必須です。');
  if (ojt_completed !== 'true') throw new Error('OJT登録のチェックは必須です。');
  if (!ojt_id) throw new Error('OJT IDの入力は必須です。');

  const ai = evalImprovement(id, improvement_details);
  
  const updateData = {
      'ステータス': '最終評価中',
      '改善完了報告': improvement_details,
      '報告者': team_member_1,
      '協力者1': team_member_2,
      '協力者2': team_member_3,
      'URL': reference_url,
      'OJT': true,
      'OJTID': ojt_id,
      '費用': cost,
      '工数': effort,
      '効果': effect,
      'AI改善評価': ai?.comment || "AI講評の取得失敗",
      '頻度（改善評価）': ai?.scores?.frequency || 1,
      '発生の可能性（改善後評価）': ai?.scores?.likelihood || 1,
      '重篤度（改善後評価）': ai?.scores?.severity || 1,
      '最終通知日時': new Date()
  };
  
  updateRowById(id, updateData);

  const incident = getDataById(id);
  if (incident.hopeful_evaluator) {
    notifyImprovementComplete(incident.hopeful_evaluator, id, incident.subject);
  }

  return `ID: ${id} の改善報告を受け付け、評価者へ通知しました。`;
}

/**
 * 改善後のスコアを、評価者が修正する
 */
function updatePostImproveScores(data) {
  const { id, post_frequency, post_likelihood, post_severity } = data;
  if (!id) throw new Error("IDが未指定です。");
  
  const beforeData = getDataById(id);
  const evaluator = Session.getActiveUser().getEmail();
  
  const updateData = {
    '頻度（改善評価）': post_frequency,
    '発生の可能性（改善後評価）': post_likelihood,
    '重篤度（改善後評価）': post_severity,
    '評価者による修正（最終）': new Date()
  };

  if (String(beforeData.post_frequency)  !== String(post_frequency))  logUpdate(id, '頻度（改善評価）', beforeData.post_frequency, post_frequency, evaluator);
  if (String(beforeData.post_likelihood) !== String(post_likelihood)) logUpdate(id, '発生の可能性（改善後評価）', beforeData.post_likelihood, post_likelihood, evaluator);
  if (String(beforeData.post_severity)   !== String(post_severity))   logUpdate(id, '重篤度（改善後評価）', beforeData.post_severity, post_severity, evaluator);

  updateRowById(id, updateData);

  return `改善後評価を更新しました。`;
}

/**
 * 最終評価コメントを保存し、ステータスを「完了」にする
 */
function submitFinalEvaluation(data) {
  const { id, final_eval_comment } = data;
  if (!id || !final_eval_comment) throw new Error("最終評価コメントは必須です。");
  
  const updateData = {
    'ステータス': '完了',
    '最終評価コメント': final_eval_comment,
    '最終通知日時': new Date() // AP列
  };
  
  updateRowById(id, updateData);
  return '最終評価を保存し、完了しました。';
}

/**
 * 差し戻し処理
 */
function revert(data) {
  const { id, targetStatus, reason } = data;
  if (!id || !targetStatus || !reason) throw new Error('差し戻し理由の入力は必須です。');

  const incident = getDataById(id);
  const existingReason = incident.revert_reason ? incident.revert_reason + "\n" : "";
  const newReason = `${existingReason}[${formatDate(new Date())} → ${targetStatus}] ${reason}`;

  updateRowById(id, { 'ステータス': targetStatus, '差し戻し理由': newReason });

  notifyRevert(incident, targetStatus, reason);
  return `ステータスを「${targetStatus}」に差し戻しました。`;
}

/**
 * 部署別Chat通知
 */
function notifyDepartmentChat_(department, message){
  const hooks = JSON.parse(PropertiesService.getScriptProperties().getProperty('DEPT_WEBHOOKS_JSON') || '{}');
  const url = hooks[department] || hooks['default'];
  if (url) {
    UrlFetchApp.fetch(url, {
      method:'post', contentType:'application/json', 
      payload:JSON.stringify({text:message}), muteHttpExceptions:true
    });
  }
}

