/**
 * Googleフォームから新しいインシデントが送信された際に自動的に実行されるトリガーです。
 * AI評価の実行、スプレッドシートへの記録、通知など、ワークフローを開始します。
 */
function onFormSubmit(e) {
  if (!e || !e.range) return;
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) return;

  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const row = e.range.getRow();
    const idCol = headers.indexOf('ID');
    if (idCol !== -1 && sheet.getRange(row, idCol + 1).getValue()) return;

    const formValues = e.namedValues;
    const subject = formValues['件名']?.[0] || '';
    const details = formValues['ヒヤリハット内容']?.[0] || '';
    const id = newId();
    const ai = evalNewIncident(id, subject, details);
    const data = { 
      'ID': id, 
      'ステータス': 'リスク評価中', 
      'レポート': `${WEBAPP_BASE_URL}?id=${id}&view=report`,
      'アプリURL': `${WEBAPP_BASE_URL}?id=${id}` // アプリURLを追加
    };
    if (ai && !ai.error) {
      Object.assign(data, {
        '頻度（AI評価）': ai?.scores?.frequency || 1,
        '発生の可能性（AI評価）': ai?.scores?.likelihood || 1,
        '重篤度（AI評価）': ai?.scores?.severity || 1,
        '評価理由': ai.reason || "評価理由の取得失敗",
        '改善プラン': ai.plan || "改善プランの取得失敗",
        '事故の型分類（AI評価）': ai.accident_type || "分類取得失敗",
        '起因物（AI評価）': ai.causal_agent || "起因物取得失敗"
      });
    } else {
      Object.assign(data, { '評価理由': `AI評価エラー: ${ai?.error || '不明'}` });
    }
    updateRowById(id, data, row);
    setFormulas(sheet, row, headers);
    clearIdCache();

    const submitter = formValues['メールアドレス']?.[0];
    const evaluator = formValues['希望する評価者']?.[0];
    const department = formValues['所属']?.[0];
    if (submitter) notifySubmitter(submitter, id, subject);
    if (evaluator && !ai.error) notifyEvaluator(evaluator, department, id, subject);

  } catch (err) {
    Logger.log(`onFormSubmit error: ${err.message}\n${err.stack}`);
  } finally {
    lock.releaseLock();
  }
}

/**
 * スプレッドシートの指定された行に、リスク値などを自動計算する数式を設定します。
 */
function setFormulas(sheet, row, headers) {
  const getColNotation = (header) => {
    const colIndex = headers.indexOf(header);
    return colIndex !== -1 ? sheet.getRange(row, colIndex + 1).getA1Notation() : null;
  };
  
  // AI評価
  const pCell = getColNotation('頻度（AI評価）');
  const qCell = getColNotation('発生の可能性（AI評価）');
  const rCell = getColNotation('重篤度（AI評価）');
  const riskCell = getColNotation('リスクの見積もり（AI評価）'); // S列
  const priorityCell = getColNotation('優先順位'); // T列

  if (riskCell && pCell && qCell && rCell) {
    sheet.getRange(riskCell).setFormula(`=SUM(${pCell}, ${qCell}, ${rCell})`);
    if (priorityCell) {
      sheet.getRange(priorityCell).setFormula(`=IF(${riskCell}>=15, "高", IF(${riskCell}>=8, "中", "低"))`);
    }
  }

  // 改善後評価
  const arCell = getColNotation('頻度（改善評価）');
  const asCell = getColNotation('発生の可能性（改善後評価）');
  const atCell = getColNotation('重篤度（改善後評価）');
  const postRiskCell = getColNotation('改善後のリスクの見積もり'); // AU列
  const postPriorityCell = getColNotation('改善後の優先度'); // AW列
  
  if (postRiskCell && arCell && asCell && atCell) {
     sheet.getRange(postRiskCell).setFormula(`=SUM(${arCell}, ${asCell}, ${atCell})`);
  }
  if (postPriorityCell && postRiskCell) {
    sheet.getRange(postPriorityCell).setFormula(`=IF(${postRiskCell}>=15, "高", IF(${postRiskCell}>=8, "中", "低"))`);
  }

  // リスク低減値
  const riskReductionCell = getColNotation('リスク低減値'); // AY列
  if (riskReductionCell && riskCell && postRiskCell) {
    sheet.getRange(riskReductionCell).setFormula(`=IF(AND(ISNUMBER(${riskCell}), ISNUMBER(${postRiskCell})), ${riskCell} - ${postRiskCell}, "")`);
  }
}

