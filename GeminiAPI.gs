/** 新規インシデントをAIで評価する */
function evalNewIncident(id, subject, details) {
  const prompt = `
    以下のインシデント報告について、リスク評価を行ってください。
    回答は必ず日本語で、以下のキーを含むJSON形式で生成してください。
    - "reason": 評価理由
    - "plan": 改善プランの提案
    - "accident_type": 事故の型分類
    - "causal_agent": 起因物
    - "scores": {"frequency": 数値, "likelihood": 数値, "severity": 数値}

    ---
    件名: ${subject}
    内容: ${details}
    ---
  `;
  return callGemini(id, prompt, 'evalNewIncident');
}

/** 改善報告をAIで評価する */
function evalImprovement(id, details) {
  const prompt = `
    以下のインシデント改善報告について、評価を行ってください。
    回答は必ず日本語で、以下のキーを含むJSON形式で生成してください。
    - "comment": 改善策に対するAIの講評
    - "scores": {"frequency": 数値, "likelihood": 数値, "severity": 数値}

    ---
    改善内容: ${details}
    ---
  `;
  return callGemini(id, prompt, 'evalImprovement');
}

/** Gemini APIを呼び出す共通関数（リトライ機能付） */
function callGemini(id, prompt, label) {
  if (!API_KEY) throw new Error('API_KEY is not set.');
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${GEMINI_MODEL}:generateContent?key=${API_KEY}`;
  const payload = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: { "responseMimeType": "application/json" }
  };
  const options = { method: 'post', contentType: 'application/json', payload: JSON.stringify(payload), muteHttpExceptions: true };

  for (let i = 0; i < 3; i++) {
    try {
      const res = UrlFetchApp.fetch(url, options);
      if (res.getResponseCode() === 200) {
        const json = JSON.parse(res.getContentText());
        const usage = json.usageMetadata || { promptTokenCount: 0, candidatesTokenCount: 0 };
        logTokens(id, usage.promptTokenCount, usage.candidatesTokenCount, label);
        return JSON.parse(json.candidates[0].content.parts[0].text);
      }
      if (res.getResponseCode() === 503 && i < 2) Utilities.sleep(2000);
      else return { error: `API Error ${res.getResponseCode()}` };
    } catch (e) {
      if (i >= 2) return { error: "Fetch Error" };
      Utilities.sleep(2000);
    }
  }
}

/** トークン使用量を記録する */
function logTokens(id, inTokens, outTokens, label) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TOKEN_SHEET);
    if (!sheet) return;
    const cost = (inTokens / 1000 * 0.000125) + (outTokens / 1000 * 0.000375);
    sheet.appendRow([new Date(), id, inTokens, outTokens, cost.toFixed(6), label]);
  } catch (e) {
    Logger.log(`Token logging failed: ${e.message}`);
  }
}