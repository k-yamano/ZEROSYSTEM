/** 新規インシデントをAIで評価する */
function evalNewIncident(id, subject, details) {
  // ★★★ AIへの指示（プロンプト）を修正 ★★★
  const prompt = `
    以下のインシデント報告について、リスク評価を行ってください。
    回答は必ず日本語で、以下のキーを含むJSON形式で生成してください。

    - "reason": 評価理由
    - "plan": 改善プランの提案
    - "accident_type": 事故の型分類
    - "causal_agent": 起因物
    - "scores": {
        "frequency": 数値,  // 頻度。必ず [1, 2, 4] のいずれかを選択。
        "likelihood": 数値, // 発生の可能性。必ず [1, 2, 4, 6] のいずれかを選択。
        "severity": 数値    // 重篤度。必ず [1, 3, 6, 10] のいずれかを選択。
      }

    ---
    件名: ${subject}
    内容: ${details}
    ---
  `;
  return callGemini(id, prompt, 'evalNewIncident');
}

/** 改善報告をAIで評価する */
function evalImprovement(id, details) {
  // ★★★ こちらのプロンプトも同様に修正 ★★★
  const prompt = `
    以下のインシデント改善報告について、評価を行ってください。
    回答は必ず日本語で、以下のキーを含むJSON形式で生成してください。

    - "comment": 改善策に対するAIの講評
    - "scores": {
        "frequency": 数値,  // 頻度。必ず [1, 2, 4] のいずれかを選択。
        "likelihood": 数値, // 発生の可能性。必ず [1, 2, 4, 6] のいずれかを選択。
        "severity": 数値    // 重篤度。必ず [1, 3, 6, 10] のいずれかを選択。
      }

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
      const responseCode = res.getResponseCode();
      
      if (responseCode === 200) {
        const json = JSON.parse(res.getContentText());
        const usage = json.usageMetadata || { promptTokenCount: 0, candidatesTokenCount: 0 };
        logTokens(id, usage.promptTokenCount, usage.candidatesTokenCount, label);
        return JSON.parse(json.candidates[0].content.parts[0].text);
      }

      if (responseCode === 503 && i < 2) {
        Logger.log(`Gemini API returned ${responseCode}. Retrying in 2 seconds... (Attempt ${i + 1})`);
        Utilities.sleep(2000);
        continue;
      }
      
      const errorContent = res.getContentText();
      Logger.log(`Gemini API Error. Status: ${responseCode}, Response: ${errorContent}`);
      throw new Error(`API request failed with status ${responseCode}`);

    } catch (e) {
      Logger.log(`Fetch Error during Gemini API call: ${e.message} (Attempt ${i + 1})`);
      if (i < 2) {
        Utilities.sleep(2000);
      } else {
        throw new Error(`Failed to call Gemini API after 3 attempts. Last error: ${e.message}`);
      }
    }
  }
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
      const responseCode = res.getResponseCode();
      
      if (responseCode === 200) {
        const json = JSON.parse(res.getContentText());
        const usage = json.usageMetadata || { promptTokenCount: 0, candidatesTokenCount: 0 };
        logTokens(id, usage.promptTokenCount, usage.candidatesTokenCount, label);
        return JSON.parse(json.candidates[0].content.parts[0].text);
      }

      // 503 Service Unavailable の場合は少し待ってリトライ
      if (responseCode === 503 && i < 2) {
        Logger.log(`Gemini API returned ${responseCode}. Retrying in 2 seconds... (Attempt ${i + 1})`);
        Utilities.sleep(2000);
        continue; // 次のループへ
      }
      
      // その他のエラーの場合は、詳細をログに記録してエラーをスロー
      const errorContent = res.getContentText();
      Logger.log(`Gemini API Error. Status: ${responseCode}, Response: ${errorContent}`);
      throw new Error(`API request failed with status ${responseCode}`);

    } catch (e) {
      Logger.log(`Fetch Error during Gemini API call: ${e.message} (Attempt ${i + 1})`);
      if (i < 2) {
        Utilities.sleep(2000); // ネットワークエラーなどの場合もリトライ
      } else {
        // 3回試行しても失敗した場合
        throw new Error(`Failed to call Gemini API after 3 attempts. Last error: ${e.message}`);
      }
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