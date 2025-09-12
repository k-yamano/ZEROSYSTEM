/**
 * ★★★【追加】★★★
 * IDをキーに行を検索し、行番号やヘッダー情報を含むオブジェクトを返す
 * @param {string} id インシデントID
 * @returns {Object|null} {sh, headers, rowNum, values} or null if not found
 */
function findSheetRowById_(id) {
  const sh = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const idColIndex = headers.indexOf('ID');
  if (idColIndex === -1) return null;

  const data = sh.getRange(2, 1, sh.getLastRow(), sh.getLastColumn()).getValues();
  const rowIndex = data.findIndex(row => String(row[idColIndex]).trim() === String(id).trim());

  if (rowIndex === -1) return null;
  
  const values = {};
  headers.forEach((h, i) => { if(h) values[h] = data[rowIndex][i]; });

  return { sh, headers, rowNum: rowIndex + 2, values };
}


/**
 * IDをキーにスプレッドシートからインシデントデータを取得し、
 * 整形されたオブジェクトとして返す
 */
function getIncidentById(id) {
  if (!id) return {};
  const rowData = findSheetRowById_(id);
  if (!rowData) return {};

  const result = {};
  for (const h in rowData.values) {
    const key = headerToKeyMapping[h] || h;
    result[key] = rowData.values[h];
  }
  return result;
}

// スプレッドシートの列名とプログラムで使うキー名の対応表
const headerToKeyMapping = {
  'ID': 'unique_id', 'タイムスタンプ': 'timestamp', 'メールアドレス': 'reporter_email',
  '所属': 'department', '件名': 'subject', 'ヒヤリハット内容': 'details', 
  '希望する評価者': 'hopeful_evaluator',
  'ステータス': 'status', '担当部署': 'evaluator', '評価者によるコメント': 'secondary_eval_comment',
  '評価理由': 'evaluation_reason', '改善プラン': 'improvement_plan_ai',
  '頻度（二次評価）': 'frequency_2', '発生の可能性（二次評価）': 'likelihood_2', '重篤度（二次評価）': 'severity_2',
  'リスクの見積もり': 'risk_score', '優先順位': 'priority',
  '改善担当部署': 'improvement_department', '改善完了報告': 'improvement_details',
  '報告者': 'team_member_1', '協力者1': 'team_member_2', '協力者2': 'team_member_3',
  '協力者3': 'team_member_4', '協力者4': 'team_member_5',
  'URL': 'reference_url', 'OJT': 'ojt_completed',
  '頻度（改善評価）': 'post_frequency', '発生の可能性（改善後評価）': 'post_likelihood', '重篤度（改善後評価）': 'post_severity',
  '改善後のリスクの見積もり': 'post_risk', '改善後の優先度': 'post_priority',
  'リスク低減値': 'risk_reduction_value',
  'AI改善評価': 'improvement_ai_comment', '最終評価コメント': 'final_eval_comment',
  '評価者による修正（二次）': 'modified_by_evaluator_secondary',
  '評価者による修正（最終）': 'modified_by_evaluator_final',
  '費用': 'cost', '工数': 'effort', '効果': 'effect', 'OJTID': 'ojt_id',
  '期限': 'deadline', '暫定予算': 'provisional_budget',
  '最終通知日時': 'final_notification_timestamp'
};


/**
 * IDをキーにスプレッドシートの指定された行を更新する
 */
function updateSheetRow_(id, updateObject) {
  const rowInfo = findSheetRowById_(id);
  if (!rowInfo) throw new Error(`ID: ${id} がシートに見つかりません。`);
  
  const { sh, headers, rowNum } = rowInfo;
  for (const header in updateObject) {
    const colIndex = headers.indexOf(header);
    if (colIndex !== -1) {
      sh.getRange(rowNum, colIndex + 1).setValue(updateObject[header]);
    } else {
      Logger.log(`警告: 列 "${header}" がスプレッドシートに見つかりません。`);
    }
  }
}

