/**
 * Web App UIからの操作（フォーム送信など）を一元的に処理します。
 * actionパラメータに応じて、適切な関数を呼び出します。
 */
function serverAction(action, data) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) {
    throw new Error('他の処理が実行中のため、少し待ってから再度お試しください。');
  }
  try {
    logAction(action, data.id);

    switch (action) {
      case 'submitSecondaryEvaluation': return submitSecondaryEvaluation(data);
      case 'updateScores': return updateScores(data);
      case 'submitImprovementReport': return submitImprovementReport(data);
      case 'updatePostImproveScores': return updatePostImproveScores(data);
      case 'submitFinalEvaluation': return submitFinalEvaluation(data);
      case 'revert': return revert(data);
      default: throw new Error('無効な操作です。');
    }
  } catch (e) {
    logError(e, 'serverAction', {action, id: data.id});
    throw new Error(`サーバー処理エラー: ${e.message}`);
  } finally {
    lock.releaseLock();
  }
}

/**
 * 二次評価（リスク評価）を完了し、ステータスを「改善報告中」に更新します。
 */
function submitSecondaryEvaluation(data){
  const { id, department, hopeful_evaluator, secondary_eval_comment, deadline, provisional_budget } = data;
  if (!id || !department || !hopeful_evaluator || !secondary_eval_comment) {
    throw new Error('担当部署、担当者、リスク評価コメントは必須です。');
  }
  
  const updateData = {
    'ステータス': '改善報告中',
    '担当部署': department,
    '希望する評価者': hopeful_evaluator,
    'リスク評価コメント': secondary_eval_comment,
    '期限': deadline,
    '暫定予算': provisional_budget
  };
  updateRowById(id, updateData);
  
  const incidentData = getDataById(id);
  const message = `【ZEROSYSTEM 改善依頼】\nID: ${id}\n件名: ${incidentData.subject}\n担当者: ${hopeful_evaluator}\nリンク: ${WEBAPP_BASE_URL}?id=${id}`;
  notifyDepartmentChat_(department, message);
  
  return "改善指示を送信しました。";
}

/**
 * AI評価スコアを評価者が手動で修正します。
 */
function updateScores(data) {
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
 
  if (String(beforeData.frequency_ai)  !== String(frequency_ai))  logUpdate(id, '頻度（AI評価）', beforeData.frequency_ai, frequency_ai, evaluator);
  if (String(beforeData.likelihood_ai) !== String(likelihood_ai)) logUpdate(id, '発生の可能性（AI評価）', beforeData.likelihood_ai, likelihood_ai, evaluator);
  if (String(beforeData.severity_ai)   !== String(severity_ai))   logUpdate(id, '重篤度（AI評価）', beforeData.severity_ai, severity_ai, evaluator);
  
  updateRowById(id, updateData);
  return 'AI評価スコアを更新しました。';
}

/**
 * 改善完了報告を提出し、ステータスを「最終評価中」に更新します。
 */
function submitImprovementReport(data) {
  const { id, improvement_details, ojt_registered, ojt_implemented, ojt_id } = data;
  if (!id || !improvement_details) throw new Error('改善内容の詳細は必須です。');
  if (ojt_registered !== 'true') throw new Error('OJT登録のチェックは必須です。');
  if (!ojt_id) throw new Error('OJT IDの入力は必須です。');

  const ai = evalImprovement(id, improvement_details);
  const updateData = {
      'ステータス': '最終評価中', '改善完了報告': improvement_details,
      '報告者': data.team_member_1, '協力者1': data.team_member_2, '協力者2': data.team_member_3,
      'URL': data.reference_url, 
      'OJT登録': ojt_registered === 'true', 
      'OJT実施': ojt_implemented === 'true',
      'OJTID': data.ojt_id,
      '費用': data.cost, '工数': data.effort, '効果': data.effect,
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
 * 改善後スコアを評価者が手動で修正します。
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
 * 最終評価を完了し、ステータスを「完了」に更新します。
 */
function submitFinalEvaluation(data) {
  const { id, final_eval_comment, ojt_confirmed_final } = data;
  if (!id || !final_eval_comment) throw new Error("最終評価コメントは必須です。");
  if (ojt_confirmed_final !== 'true') throw new Error("OJT実施状況の確認は必須です。");

  const updateData = {
    'ステータス': '完了',
    '最終評価コメント': final_eval_comment,
    'OJT最終確認': ojt_confirmed_final === 'true',
    '最終通知日時': new Date()
  };
  updateRowById(id, updateData);
  return '最終評価を保存し、完了しました。';
}

/**
 * 案件を前のステータスに差し戻します。
 */
function revert(data) {
  const { id, targetStatus, reason } = data;
  if (!id || !targetStatus || !reason) throw new Error('差し戻し理由の入力は必須です。');

  const incident = getDataById(id);
  let history = [];
  try {
    if (incident.revert_history && String(incident.revert_history).trim().startsWith('[')) {
      history = JSON.parse(incident.revert_history);
    }
  } catch(e) {}
  history.unshift({
    timestamp: new Date().toISOString(),
    reverted_by: Session.getActiveUser().getEmail(),
    reason: reason,
    from_status: incident.status
  });

  updateRowById(id, { 
    'ステータス': targetStatus, 
    '差し戻し履歴': JSON.stringify(history) 
  });

  notifyRevert(incident, targetStatus, reason);
  return `ステータスを「${targetStatus}」に差し戻しました。`;
}

/** 部署別Chat通知 */
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

