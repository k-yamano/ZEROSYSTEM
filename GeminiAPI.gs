/**
 * Evaluates a new incident using the Gemini API.
 * @param {string} id The incident ID.
 * @param {string} subject The subject of the incident.
 * @param {string} details The details of the incident.
 * @returns {object} The parsed JSON response from the API.
 */
function evalNewIncident(id, subject, details) {
  // ★修正: ユーザー提案の新しいプロンプト（加算方式）に更新
  const prompt = `
    以下のインシデント報告について、リスク評価を行ってください。

    # 評価対象
    件名: ${subject}
    内容: ${details}

    # 回答形式
    必ず日本語で、以下の形式のJSONで回答してください：
    {
      "reason": "評価理由の詳細説明（200文字以内）",
      "plan": "具体的な改善プランの提案（300文字以内）",
      "accident_type": "事故の型分類",
      "causal_agent": "起因物の特定",
      "additional_classification": "品質・環境案件の場合は4M分析結果、それ以外は空欄",
      "scores": {
        "frequency": "頻度スコア",
        "likelihood": "可能性スコア", 
        "severity": "重篤度スコア"
      },
      "risk_level": "総合リスクレベル（ランクⅠ～Ⅳ）",
      "priority": "対応優先度（A～D）"
    }

    # 評価基準
    ## スコア選択肢（必須）
    - 頻度（frequency）: [1, 2, 4] から選択 (1:稀, 2:時々, 4:頻繁)
    - 可能性（likelihood）: [1, 2, 4, 6] から選択 (1:極めて低い, 2:低い, 4:中程度, 6:高い)
    - 重篤度（severity）: [1, 3, 6, 10] から選択 (1:軽微, 3:中程度, 6:重大, 10:致命的)

    ## 総合リスクレベル
    リスクスコア = 頻度 + 可能性 + 重篤度 (0～20点)
    - ランクⅣ (12～20点): 直ちに解決すべき問題がある
    - ランクⅢ (9～11点): 重大な問題がある
    - ランクⅡ (6～8点): 多少問題がある
    - ランクⅠ (5点以下): 必要に応じて低減措置

    ## 事故型分類（厚生労働省基準）
    以下から最も適切なものを1つ選択：
    墜落・転落, 転倒, 激突, 飛来・落下, 崩壊・倒壊, 挟まれ・巻き込まれ, 切れ・こすれ, 踏み抜き, おぼれ, 高温・低温の物との接触, 有害物質等との接触, 爆発, 火災, 交通事故（道路）, 交通事故（その他）, その他
  `;
  return callGemini(id, prompt, 'evalNewIncident');
}

/**
 * Evaluates an improvement report using the Gemini API.
 */
function evalImprovement(id, details) {
  const prompt = `
    以下のインシデント改善報告について、評価を行ってください。
    回答は必ず日本語で、以下のキーを含むJSON形式で生成してください。
    - "comment": 改善策に対するAIの講評
    - "scores": { "frequency": 数値, "likelihood": 数値, "severity": 数値 }
    頻度は[1,2,4]、可能性は[1,2,4,6]、重篤度は[1,3,6,10]から必ず選択してください。
    ---
    改善内容: ${details}
    ---
  `;
  return callGemini(id, prompt, 'evalImprovement');
}

/**
 * Common function to call the Gemini API with retry logic.
 * @param {string} id The incident ID for logging.
 * @param {string} prompt The prompt to send to the model.
 * @param {string} label A label for logging the token usage.
 * @returns {object} The parsed JSON response from the API.
 */
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

      if (responseCode >= 500 && i < 2) { // Retry on server errors
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

/**
 * Logs token usage to a spreadsheet.
 * @param {string} id The incident ID.
 * @param {number} inTokens The number of input tokens.
 * @param {number} outTokens The number of output tokens.
 * @param {string} label A label for the API call.
 */
function logTokens(id, inTokens, outTokens, label) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TOKEN_SHEET);
    if (!sheet) return;
    const cost = (inTokens / 1000 * 0.000125) + (outTokens / 1000 * 0.000375); // Example pricing
    sheet.appendRow([new Date(), id, inTokens, outTokens, cost.toFixed(6), label]);
  } catch (e) {
    Logger.log(`Token logging failed: ${e.message}`);
  }
}

