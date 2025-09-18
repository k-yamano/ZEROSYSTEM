/** フォーム送信時に実行されるメインの処理 */
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

    // 1. AI評価を先に呼び出す
    const ai = evalNewIncident(id, subject, details);
    
    // 2. AIの評価結果を格納するデータオブジェクトを作成
    const dataForUpdate = { 'ID': id, 'ステータス': 'リスク評価中', 'レポート': `${WEBAPP_BASE_URL}?id=${id}&view=report` };
    if (ai && !ai.error) {
      Object.assign(dataForUpdate, {
        '頻度（AI評価）': ai?.scores?.frequency || 1,
        '発生の可能性（AI評価）': ai?.scores?.likelihood || 1,
        '重篤度（AI評価）': ai?.scores?.severity || 1,
        '評価理由': ai.reason || "評価理由の取得失敗",
        '改善プラン': ai.plan || "改善プランの取得失敗",
        '事故の型分類（AI評価）': ai.accident_type || "分類取得失敗",
        '起因物（AI評価）': ai.causal_agent || "起因物取得失敗"
      });
    } else {
      Object.assign(dataForUpdate, {
        '評価理由': `AI評価エラー: ${ai?.error || '不明'}`, '事故の型分類（AI評価）': 'AI評価エラー', '起因物（AI評価）': 'AI評価エラー'
      });
    }
    
    // 3. AIの結果を書き込む
    updateRowById(id, dataForUpdate, row);

    // 4. 最後に数式を設定する（これにより上書きを防ぐ）
    setFormulas(sheet, row, headers);
    
    // キャッシュクリアと通知
    clearIdCache();
    const submitter = formValues['メールアドレス']?.[0];
    const evaluator = formValues['希望する評価者']?.[0];
    const department = formValues['所属']?.[0];
    if (submitter) notifySubmitter(submitter, id, subject);
    if (evaluator && !ai.error) notifyEvaluator(evaluator, department, id, subject);

  } catch (err) {
    Logger.log(`onFormSubmit error: ${err.message}\n${err.stack}`);
    const adminEmail = Session.getEffectiveUser().getEmail();
    const subject = `[要確認] ZEROSYSTEM onFormSubmitエラー`;
    const body = `フォーム送信処理中にエラーが発生しました。\n\nエラー内容:\n${err.message}\n\nスタックトレース:\n${err.stack}`;
    MailApp.sendEmail(adminEmail, subject, body);
  } finally {
    lock.releaseLock();
  }
}

/** 指定された行に数式を設定する */
function setFormulas(sheet, row, headers) {
  const getColNotation = (header) => {
    const colIndex = headers.indexOf(header);
    return colIndex !== -1 ? sheet.getRange(row, colIndex + 1).getA1Notation() : null;
  };
  
  // 既存の数式設定 (リスク評価)
  const riskCell = getColNotation('リスクの見積もり（AI評価）');
  const pCell = getColNotation('頻度（AI評価）');
  const qCell = getColNotation('発生の可能性（AI評価）');
  const rCell = getColNotation('重篤度（AI評価）');
  const priorityCell = getColNotation('優先順位');

  if (riskCell && pCell && qCell && rCell) {
    sheet.getRange(riskCell).setFormula(`=SUM(${pCell}:${rCell})`);
    if (priorityCell) {
      sheet.getRange(priorityCell).setFormula(`=IF(${riskCell}>=15, "高", IF(${riskCell}>=8, "中", "低"))`);
    }
  }

  // ★★★ 新しい数式を追加 ★★★
  const postRiskCell = getColNotation('改善後のリスクの見積もり'); // AU列相当
  const postPriorityCell = getColNotation('改善後の優先度'); // AW列相当
  const riskReductionCell = getColNotation('リスク低減値'); // AY列相当
  
  // AW列: 改善後優先度
  if (postPriorityCell && postRiskCell) {
    sheet.getRange(postPriorityCell).setFormula(`=IF(${postRiskCell}>=15, "高", IF(${postRiskCell}>=8, "中", "低"))`);
  }

  // AY列: リスク低減値
  if (riskReductionCell && riskCell && postRiskCell) {
    sheet.getRange(riskReductionCell).setFormula(`=IF(AND(ISNUMBER(${riskCell}), ISNUMBER(${postRiskCell})), ${riskCell} - ${postRiskCell}, "")`);
  }
}
