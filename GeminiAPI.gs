// Gemini APIを呼び出す際のプロンプト設定

/**
 * 新規インシデントの初期AI評価を実行する
 */
function runInitialAI(id, subject, details) {
  const prompt = `
    以下のインシデント報告について、リスク評価を行ってください。
    回答は必ず日本語で、以下のキーを含むJSON形式で生成してください。
    - "reason": 評価理由（なぜこのリスクが重要か）
    - "plan": 改善プランの提案
    - "accident_type": 事故の型分類（例: 墜落・転落, 飛来・落下, 激突, 崩壊・倒壊, 感電, 火災, 爆発, 破裂, 有害物との接触, 切れ・こすれ, 踏み抜き, 溺れ, 高温・低温物との接触, 動作の反動・無理な動作）
    - "causal_agent": 起因物（二次評価）（例: 機械, 構造物, 物質, 環境, 人）
    - "scores": 以下の3つの評価項目を数値で評価
      - "frequency": 発生頻度 (1:ほとんどない, 2:たまにある, 4:よくある)
      - "likelihood": 発生可能性 (1:ほとんどない, 2:たまにある, 4:よくある, 6:ほぼ確実)
      - "severity": 影響度 (1:軽微, 3:中程度, 6:重大, 10:致命的)

    ---
    件名: ${subject}
    内容: ${details}
    ---
  `;
  const result = callGeminiAPI_(id, prompt, 'runInitialAI');
  return {
    reason: result.reason || "AIによる評価理由の取得に失敗",
    plan: result.plan || "AIによる改善プランの取得に失敗",
    accident_type: result.accident_type || "AIによる分類取得に失敗",
    causal_agent: result.causal_agent || "AIによる起因物取得に失敗",
    scores: {
      frequency: result.scores ? (result.scores.frequency || 1) : 1,
      likelihood: result.scores ? (result.scores.likelihood || 1) : 1,
      severity: result.scores ? (result.scores.severity || 1) : 1,
    }
  };
}

/**
 * 改善報告を基にAIによる改善後評価を実行する
 */
function runPostImprovementAI(id, improvement_details) {
  const prompt = `
    以下のインシデント改善報告について、評価を行ってください。
    回答は必ず日本語で、以下のキーを含むJSON形式で生成してください。
    - "comment": 改善策に対するAIの講評
    - "scores": 改善後のリスクを3項目で再評価
      - "frequency": 発生頻度 (1:ほとんどない, 2:たまにある, 4:よくある)
      - "likelihood": 発生可能性 (1:ほとんどない, 2:たまにある, 4:よくある, 6:ほぼ確実)
      - "severity": 影響度 (1:軽微, 3:中程度, 6:重大, 10:致命的)

    ---
    改善内容: ${improvement_details}
    ---
  `;
  const result = callGeminiAPI_(id, prompt, 'runPostImprovementAI');
  return {
    comment: result.comment || "AIによる講評の取得に失敗",
    scores: {
      frequency: result.scores ? (result.scores.frequency || 1) : 1,
      likelihood: result.scores ? (result.scores.likelihood || 1) : 1,
      severity: result.scores ? (result.scores.severity || 1) : 1,
    }
  };
}

/**
 * Gemini APIを呼び出す共通関数
 */
function callGeminiAPI_(id, prompt, functionLabel) {
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${GEMINI_MODEL}:generateContent?key=${API_KEY}`;
  const payload = {
    contents: [{ parts: [{ text: prompt }] }],
    "generationConfig": {
      "responseMimeType": "application/json",
    }
  };
  const options = { method: 'post', contentType: 'application/json', payload: JSON.stringify(payload), muteHttpExceptions: true };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const responseBody = response.getContentText();

    if (responseCode === 200) {
      const json = JSON.parse(responseBody);
      const candidate = json.candidates[0];
      const usage = json.usageMetadata || { promptTokenCount: 0, candidatesTokenCount: 0, totalTokenCount: 0 };
      logTokenUsage(id, usage.promptTokenCount, usage.candidatesTokenCount, usage.totalTokenCount, functionLabel);
      return JSON.parse(candidate.content.parts[0].text);
    } else {
      console.error(`Gemini API Error (${responseCode}): ${responseBody}`);
      return {};
    }
  } catch (e) {
    console.error(`Gemini API Fetch Error: ${e}`);
    return {};
  }
}

