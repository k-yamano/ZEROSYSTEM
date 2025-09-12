/**
 * Googleフォームからの投稿をスプレッドシートのトリガーで実行
 * @param {Object} e イベントオブジェクト
 */
function onFormSubmitTrigger(e){
  if (!e) return;

  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) return;

  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sh = ss.getSheetByName(SHEET_NAME);
    const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    
    const row = e.range.getRow();
    const idCol = headers.indexOf('ID') + 1;
    if (idCol <= 0) throw new Error('ID 列が見つかりません');

    const existingId = String(sh.getRange(row, idCol).getValue() || '').trim();
    if (existingId) {
      Logger.log(`skip: row ${row} already processed (ID=${existingId})`);
      return;
    }

    const nv = {};
    for (const k in e.namedValues) nv[k] = e.namedValues[k][0];

    const id = generateUniqueId_();
    const ai = runInitialAI(id, nv['件名'], nv['ヒヤリハット内容']);

    const oCol = getSheetColumnName(headers.indexOf('頻度（二次評価）') + 1);
    const pCol = getSheetColumnName(headers.indexOf('発生の可能性（二次評価）') + 1);
    const qCol = getSheetColumnName(headers.indexOf('重篤度（二次評価）') + 1);
    const rCol = getSheetColumnName(headers.indexOf('リスクの見積もり') + 1);

    const riskFormula = oCol && pCol && qCol ? `=SUM(${oCol}${row}, ${pCol}${row}, ${qCol}${row})` : '数式エラー';
    const priorityFormula = rCol ? `=IF(${rCol}${row}>=15, "高", IF(${rCol}${row}>=8, "中", "低"))` : '数式エラー';

    const write = {
      'ID': id,
      'レポート': buildIncidentReportUrl(id),
      'ステータス': 'リスク評価中', // ★★★ ステータス名変更
      '評価理由': ai.reason,
      '改善プラン': ai.plan,
      '頻度（二次評価）': ai.scores.frequency,
      '発生の可能性（二次評価）': ai.scores.likelihood,
      '重篤度（二次評価）': ai.scores.severity,
      '事故の型分類': ai.accident_type,
      '起因物（二次評価）': ai.causal_agent,
      'リスクの見積もり': riskFormula,
      '優先順位': priorityFormula,
    };

    headers.forEach((h, i) => {
      if (h in write) sh.getRange(row, i + 1).setValue(write[h]);
    });

    sendInitialEmail_(nv['所属'], nv['メールアドレス'], id, nv['件名'], nv['ヒヤリハット内容']);

    Logger.log(`processed row=${row}, id=${id}`);
  } finally {
    lock.releaseLock();
  }
}

