/**
 * 二次評価をWebアプリから受け取り処理する
 */
function submitSecondaryEvaluation(data) {
  if (!data.id || data.id.includes('<') || data.id.includes('>')) throw new Error('無効なIDが指定されました。ページをリロードしてください。');
  const { id, department, deadline, provisional_budget, secondary_eval_comment, hopeful_evaluator } = data;
  if (!id || !department || !secondary_eval_comment) throw new Error('必須項目が入力されていません。');
  
  updateSheetRow_(id, { 
    'ステータス': '改善報告中', // ★★★ ステータス名変更
    '担当部署': department, 
    '希望する評価者': hopeful_evaluator, 
    '期限': deadline, 
    '暫定予算': provisional_budget, 
    '評価者によるコメント': secondary_eval_comment 
  });
  
  const incidentData = getIncidentById(id);
  const webhookUrl = getDeptWebhookUrl_(department);
  if (webhookUrl) { sendToChat_(webhookUrl, `[${id}] リスク評価が完了しました。\n所属：${incidentData.department}\n担当部署：${department}\n${buildIncidentUrl(id)}`); }
  return `ID: ${id} のリスク評価を受け付けました。`;
}

/**
 * 改善報告をWebアプリから受け取り処理する
 */
function submitImprovementReport(data) {
  if (!data.id || data.id.includes('<') || data.id.includes('>')) throw new Error('無効なIDが指定されました。ページをリロードしてください。');
  const { id, improvement_details, team_member_1, team_member_2, team_member_3, reference_url, ojt_completed, cost, effort, effect, ojt_id } = data;
  if (!id || !improvement_details) throw new Error('改善内容の詳細は必須です。');

  const incidentData = findSheetRowById_(id);
  if (!incidentData) throw new Error(`ID: ${id} が見つかりません。`);
  
  const aiResult = runPostImprovementAI(id, improvement_details);
  
  const rCol = getSheetColumnName(incidentData.headers.indexOf('リスクの見積もり') + 1);
  const aoCol = getSheetColumnName(incidentData.headers.indexOf('改善後のリスクの見積もり') + 1);
  const reductionFormula = `=IF(AND(ISNUMBER(${rCol}${incidentData.rowNum}), ISNUMBER(${aoCol}${incidentData.rowNum})), ${rCol}${incidentData.rowNum}-${aoCol}${incidentData.rowNum}, "")`;
  
  const sheetUpdateData = {
      'ステータス': '最終評価中', // ★★★ ステータス名変更
      '改善担当部署': incidentData.values['所属'], 
      '改善完了報告': improvement_details,
      '報告者': team_member_1, '協力者1': team_member_2, '協力者2': team_member_3,
      'URL': reference_url, 'OJT': ojt_completed === 'true', '費用': cost, '工数': effort, '効果': effect, 'OJTID': ojt_id,
      'AI改善評価': aiResult.comment,
      'リスク低減値': reductionFormula
  };
  updateSheetRow_(id, sheetUpdateData);
  sendImprovementReportEmail_(incidentData.values['希望する評価者'], id, incidentData.values['件名']);

  return `ID: ${id} の改善報告を受け付けました。`;
}

/**
 * 最終評価をWebアプリから受け取り処理する
 */
function submitFinalEvaluation(data) {
  if (!data.id || data.id.includes('<') || data.id.includes('>')) throw new Error('無効なIDが指定されました。ページをリロードしてください。');
  const { id, final_eval_comment } = data;
  if (!id || !final_eval_comment) throw new Error('最終評価コメントは必須です。');

  updateSheetRow_(id, { 
    'ステータス': '完了', // ★★★ ステータス名変更
    '最終評価コメント': final_eval_comment, 
    '最終通知日時': new Date() 
  });
  
  const incidentData = getIncidentById(id);
  const webhookUrl = getDeptWebhookUrl_(incidentData.department);
  if (webhookUrl) { sendToChat_(webhookUrl, `[${id}] 最終評価が完了しました。\n最終レポート: ${buildIncidentReportUrl(id)}`); }

  return `ID: ${id} の最終評価を受け付けました。`;
}


/**
 * 二次評価スコアを更新する
 */
function updateSecondaryScores(data){
  return updateScores_(data, 
    ['頻度（二次評価）', '発生の可能性（二次評価）', '重篤度（二次評価）'],
    ['frequency_2', 'likelihood_2', 'severity_2'],
    '評価者による修正（二次）'
  );
}

/**
 * 改善後スコアを更新する
 */
function updatePostImproveScores(data) {
  return updateScores_(data, 
    ['頻度（改善評価）', '発生の可能性（改善後評価）', '重篤度（改善後評価）'],
    ['post_frequency', 'post_likelihood', 'post_severity'],
    '評価者による修正（最終）'
  );
}

/**
 * 差し戻し処理を実行する
 */
function revertStatus(data) {
  const { id, targetStatus, reason } = data;
  if (!id || !targetStatus || !reason) throw new Error('差し戻しにはID、対象ステータス、理由が必須です。');

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) throw new Error('他の処理が実行中です。');
  
  try {
    const incidentData = getIncidentById(id);
    let notifyEmail = '';
    let message = '';

    // ★★★ ステータス名変更に対応
    if (targetStatus === '改善報告中') { 
      notifyEmail = incidentData.team_member_1;
      message = `ID: ${id} の改善報告が差し戻されました。修正の上、再提出をお願いします。`;
    } else if (targetStatus === '最終評価中') {
      notifyEmail = incidentData.hopeful_evaluator;
      message = `ID: ${id} の改善後評価が差し戻されました。再評価をお願いします。`;
    } else {
      throw new Error('このステータスへの差し戻しは定義されていません。');
    }

    updateSheetRow_(id, { 'ステータス': targetStatus });
    sendRevertEmail_(notifyEmail, id, incidentData.subject, message, reason);
    
    return `ステータスを「${targetStatus}」に差し戻しました。`;
  } finally {
    lock.releaseLock();
  }
}

