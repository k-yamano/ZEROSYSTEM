/**
 * Executes on form submission from a spreadsheet trigger.
 * @param {Object} e The event object.
 */
function onFormSubmitTrigger(e){
  if (!e || !e.range) return;

  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) return;

  try {
    const sh = e.range.getSheet();
    const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    
    const row = e.range.getRow();
    const idCol = headers.indexOf('ID') + 1;
    if (idCol === 0) throw new Error('ID column not found in spreadsheet.');

    if (sh.getRange(row, idCol).getValue()) {
      Logger.log(`Skipping already processed row: ${row}`);
      return;
    }

    const nv = e.namedValues;
    const data = {};
    for (const key in nv) { data[key] = nv[key][0]; }

    const id = generateUniqueId_();
    const ai = runInitialAI(id, data['件名'], data['ヒヤリハット内容']);

    const oCol = getSheetColumnName(headers.indexOf('頻度（二次評価）') + 1);
    const pCol = getSheetColumnName(headers.indexOf('発生の可能性（二次評価）') + 1);
    const qCol = getSheetColumnName(headers.indexOf('重篤度（二次評価）') + 1);
    const rCol = getSheetColumnName(headers.indexOf('リスクの見積もり') + 1);
    const aoCol = getSheetColumnName(headers.indexOf('改善後のリスクの見積もり') + 1);

    const riskFormula = `=SUM(${oCol}${row}, ${pCol}${row}, ${qCol}${row})`;
    const priorityFormula = `=IF(${rCol}${row}>=15, "高", IF(${rCol}${row}>=8, "中", "低"))`;
    const reductionFormula = `=IF(AND(ISNUMBER(${rCol}${row}), ISNUMBER(${aoCol}${row})), ${rCol}${row}-${aoCol}${row}, "")`;

    const writeData = {
      'ID': id,
      '評価UI': buildIncidentUrl(id),
      '最終レポート': "", // Initially empty
      'ステータス': '二次評価待ち',
      '評価理由': ai.reason,
      '改善プラン': ai.plan,
      '頻度（二次評価）': ai.scores.frequency,
      '発生の可能性（二次評価）': ai.scores.likelihood,
      '重篤度（二次評価）': ai.scores.severity,
      '事故の型分類': ai.accident_type,
      '起因物（二次評価）': ai.causal_agent,
      'リスクの見積もり': riskFormula,
      '優先順位': priorityFormula,
      'リスク低減値': reductionFormula
    };

    headers.forEach((h, i) => {
      if (h in writeData) sh.getRange(row, i + 1).setValue(writeData[h]);
    });

    sendInitialEmail_(data['所属'], data['希望する評価者'], id, data['件名'], data['ヒヤリハット内容']);
    Logger.log(`Processed row=${row}, id=${id}`);
  } finally {
    lock.releaseLock();
  }
}

