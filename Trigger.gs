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

    setFormulas(sheet, row, headers);

    const formValues = e.namedValues;
    const subject = formValues['件名']?.[0] || '';
    const details = formValues['ヒヤリハット内容']?.[0] || '';
    const id = newId();
    const ai = evalNewIncident(id, subject, details);
    
    const data = { 'ID': id, 'ステータス': 'リスク評価中', 'レポート': `${WEBAPP_BASE_URL}?id=${id}&view=report` };
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
      Object.assign(data, {
        '評価理由': `AI評価エラー: ${ai?.error || '不明'}`, '事故の型分類（AI評価）': 'AI評価エラー', '起因物（AI評価）': 'AI評価エラー'
      });
    }
    updateRowById(id, data, row);

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

/** 指定された行に数式を設定する */
function setFormulas(sheet, row, headers) {
  const riskTargetCol = headers.indexOf('リスクの見積もり（AI評価）') + 1;
  const pCol = headers.indexOf('頻度（AI評価）') + 1;
  const qCol = headers.indexOf('発生の可能性（AI評価）') + 1;
  const rCol = headers.indexOf('重篤度（AI評価）') + 1;
  const priorityCol = headers.indexOf('優先順位') + 1;

  if (riskTargetCol > 0 && pCol > 0 && qCol > 0 && rCol > 0) {
    const startCell = sheet.getRange(row, pCol).getA1Notation();
    const endCell = sheet.getRange(row, rCol).getA1Notation();
    sheet.getRange(row, riskTargetCol).setFormula(`=SUM(${startCell}:${endCell})`);

    if (priorityCol > 0) {
      const riskCell = sheet.getRange(row, riskTargetCol).getA1Notation();
      sheet.getRange(row, priorityCol).setFormula(`=IF(${riskCell}>=15, "高", IF(${riskCell}>=8, "中", "低"))`);
    }
  } else {
      Logger.log("数式の設定に必要な列が見つかりませんでした。");
  }
}