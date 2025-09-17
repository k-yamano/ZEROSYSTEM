/** Webアプリ(UI)からの操作を処理する */
function serverAction(action, data) {
  try {
    switch (action) {
      case 'submitEval': return submitEval(data);
      case 'submitImprovement': return submitImprovement(data);
      case 'submitFinal': return submitFinal(data);
      case 'updateScores': return updateScores(data);
      case 'revert': return revert(data);
      default: throw new Error('Invalid action.');
    }
  } catch (e) {
    Logger.log(`serverAction error (${action}): ${e.message}`);
    throw new Error(`サーバー処理エラー: ${e.message}`);
  }
}

function submitEval(data) {
  const { id, department, deadline, provisional_budget, secondary_eval_comment, hopeful_evaluator } = data;
  if (!id || !department || !secondary_eval_comment) throw new Error('必須項目が未入力です。');
  updateRowById(id, {
    'ステータス': '改善報告中', '担当部署': department, '希望する評価者': hopeful_evaluator,
    '期限': deadline, '暫定予算': provisional_budget, '評価者によるコメント': secondary_eval_comment
  });
  return `ID: ${id} のリスク評価を受け付けました。`;
}

function submitImprovement(data) {
  const { id, improvement_details } = data;
  if (!id || !improvement_details) throw new Error('改善内容の詳細は必須です。');
  const ai = evalImprovement(id, improvement_details);
  updateRowById(id, {
      'ステータス': '最終評価中', '改善完了報告': improvement_details, '報告者': data.team_member_1,
      '協力者1': data.team_member_2, '協力者2': data.team_member_3, 'URL': data.reference_url,
      'OJT': data.ojt_completed === 'true', '費用': data.cost, '工数': data.effort, '効果': data.effect, 'OJTID': data.ojt_id,
      'AI改善評価': ai?.comment || "AI講評の取得失敗",
      '頻度（改善評価）': ai?.scores?.frequency || 1,
      '発生の可能性（改善後評価）': ai?.scores?.likelihood || 1,
      '重篤度（改善後評価）': ai?.scores?.severity || 1,
  });
  return `ID: ${id} の改善報告を受け付けました。`;
}

function submitFinal(data) {
  const { id, final_eval_comment } = data;
  if (!id || !final_eval_comment) throw new Error('最終評価コメントは必須です。');
  updateRowById(id, { 'ステータス': '完了', '最終評価コメント': final_eval_comment, '最終通知日時': new Date() });
  return `ID: ${id} の最終評価を受け付けました。`;
}

/** 評価者がUIからスコアを修正した際の処理（ログ記録機能付き） */
function updateScores(data) {
  const { id, type } = data;
  if (!id || !type) throw new Error('不正なリクエストです。');

  const beforeData = getDataById(id);
  const evaluator = Session.getActiveUser().getEmail();
  let updateData = {};
  let changedItems = {};

  if (type === 'ai') {
    updateData = {
      '頻度（AI評価）': data.frequency_ai, '発生の可能性（AI評価）': data.likelihood_ai,
      '重篤度（AI評価）': data.severity_ai, '評価者によるリスク修正': new Date()
    };
    changedItems = {
        'frequency_ai': '頻度（AI評価）', 'likelihood_ai': '発生の可能性（AI評価）', 'severity_ai': '重篤度（AI評価）'
    };
  } else if (type === 'post') {
    updateData = {
      '頻度（改善評価）': data.post_frequency, '発生の可能性（改善後評価）': data.post_likelihood,
      '重篤度（改善後評価）': data.post_severity, '評価者による修正（最終）': new Date()
    };
     changedItems = {
        'post_frequency': '頻度（改善評価）', 'post_likelihood': '発生の可能性（改善後評価）', 'post_severity': '重篤度（改善後評価）'
    };
  }

  for(const key in changedItems) {
      const sheetHeader = changedItems[key];
      const beforeValue = beforeData[key];
      const afterValue = data[key];
      if(String(beforeValue) !== String(afterValue)){
          logUpdate(id, sheetHeader, beforeValue, afterValue, evaluator);
      }
  }

  updateRowById(id, updateData);
  return `ID: ${id} のスコアを更新しました。`;
}

function revert(data) {
  const { id, targetStatus, reason } = data;
  if (!id || !targetStatus || !reason) throw new Error('差し戻し理由の入力は必須です。');
  updateRowById(id, { 'ステータス': targetStatus, '差し戻し理由': `[${formatDate(new Date())}] ${reason}` });
  const incident = getDataById(id);
  notifyRevert(incident, targetStatus, reason);
  return `ステータスを「${targetStatus}」に差し戻しました。`;
}