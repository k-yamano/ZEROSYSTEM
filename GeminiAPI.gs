/**
 * ★追加：Vertex AIのモデルごとの料金表（1000文字あたりの米ドル）
 * 最新の料金はGoogle Cloudの公式ページでご確認ください。
 * https://cloud.google.com/vertex-ai/pricing
 */
const PRICE_LIST = {
  // Gemini 1.5 Flash
  'gemini-1.5-flash-001': { input: 0.000125, output: 0.000375 },
  'gemini-1.5-flash-latest': { input: 0.000125, output: 0.000375 },
  // Gemini 1.0 Pro
  'gemini-1.0-pro': { input: 0.000125, output: 0.000375 },
  'gemini-1.0-pro-001': { input: 0.000125, output: 0.000375 },
  // 他のモデルも必要に応じてここに追加
};


/**
 * AIの応答テキストからJSON部分のみを抽出するヘルパー関数
 * @param {string} text - AIからの応答テキスト全体
 * @returns {object|null} - 抽出されたJSONオブジェクト、または見つからない場合はnull
 */
function extractJson_(text) {
  const match = text.match(/```json\s*([\s\S]*?)\s*```|({[\s\S]*})/);
  if (match) {
    const jsonString = match[1] || match[2];
    try {
      return JSON.parse(jsonString);
    } catch (e) {
      Logger.log(`JSON Parse Error after extraction: ${e.message}`);
      return null;
    }
  }
  return null;
}

/**
 * Evaluates a new incident using the Gemini API.
 */
function evalNewIncident(id, subject, details) {
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
  return callVertexAI(id, prompt, 'evalNewIncident');
}

/**
 * ★★★ 全面的に修正: 改善評価プロンプトの最適化 ★★★
 * Evaluates an improvement report using the Gemini API.
 */
function evalImprovement(id, improvementDetails) {
  // 元のインシデントデータを取得して、評価の文脈をAIに提供する
  const incidentData = getDataById(id);
  if (!incidentData || !incidentData.subject) {
    throw new Error(`ID ${id} の元データが見つかりませんでした。`);
  }

  const prompt = `
    以下のインシデントとその改善報告について、改善後のリスクを再評価してください。

    # 当初のインシデント内容
    ## 件名
    ${incidentData.subject}

    ## 詳細
    ${incidentData.details}

    # 実施された改善策
    ${improvementDetails}

    # 評価タスク
    1. 実施された改善策が、当初の問題の根本原因に対して効果的か評価してください。
    2. 改善後のリスク（頻度、可能性、重篤度）を再評価し、新しいスコアを付けてください。
    3. 評価の理由や、もし残存リスクがあればそれを講評として記述してください。

    # 回答形式
    必ず日本語で、以下のキーを含むJSON形式で回答してください：
    {
      "comment": "改善策に対するAIの講評（根本原因への対応、残存リスクなど）",
      "scores": { 
        "frequency": "改善後の頻度スコア", 
        "likelihood": "改善後の可能性スコア", 
        "severity": "改善後の重篤度スコア" 
      }
    }

    # 評価基準（スコア選択肢）
    - 頻度 (frequency): [1, 2, 4]
    - 可能性 (likelihood): [1, 2, 4, 6]
    - 重篤度 (severity): [1, 3, 6, 10]
  `;
  return callVertexAI(id, prompt, 'evalImprovement');
}


/**
 * Vertex AI API を呼び出す共通関数
 */
function callVertexAI(id, prompt, label) {
  if (!GCP_PROJECT_ID) throw new Error('スクリプトプロパティに GCP_PROJECT_ID が設定されていません。');

  const url = `https://${GCP_REGION}-aiplatform.googleapis.com/v1/projects/${GCP_PROJECT_ID}/locations/${GCP_REGION}/publishers/google/models/${GEMINI_MODEL}:generateContent`;

  const options = {
    method: 'post',
    contentType: 'application/json',
    muteHttpExceptions: true,
    headers: {
      'Authorization': 'Bearer ' + ScriptApp.getOAuthToken()
    },
    payload: JSON.stringify({
      contents: [{
        "role": "user",
        "parts": [{ "text": prompt }]
      }],
      generationConfig: {
        "responseMimeType": "application/json"
      }
    })
  };

  for (let i = 0; i < 3; i++) {
    try {
      const res = UrlFetchApp.fetch(url, options);
      const responseCode = res.getResponseCode();
      const responseBody = res.getContentText();

      if (responseCode === 200) {
        const json = JSON.parse(responseBody);
        if (json.candidates && json.candidates[0].content && json.candidates[0].content.parts[0]) {
          const responseText = json.candidates[0].content.parts[0].text;
          const extractedJson = extractJson_(responseText);
          if (extractedJson) {
            const usage = json.usageMetadata || { promptTokenCount: 0, candidatesTokenCount: 0 };
            logTokens(id, usage.promptTokenCount, usage.candidatesTokenCount, label);
            return extractedJson;
          } else {
            throw new Error('AIからの応答をJSONとして解釈できませんでした。');
          }
        } else {
          throw new Error('AIからの応答が空でした。');
        }
      }

      Logger.log(`Vertex AI API Error. Status: ${responseCode}, Response: ${responseBody}`);
      if (responseCode >= 500 && i < 2) {
        Utilities.sleep(2000);
        continue;
      }
      throw new Error(`API request failed with status ${responseCode}`);
    } catch (e) {
      Logger.log(`Fetch Error during Vertex AI API call: ${e.message} (Attempt ${i + 1})`);
      if (i < 2) {
        Utilities.sleep(2000);
      } else {
        throw new Error(`Failed to call Vertex AI API after 3 attempts. Last error: ${e.message}`);
      }
    }
  }
}

/**
 * Vertex AIの料金体系に合わせてトークン使用量を記録する関数
 */
function logTokens(id, inTokens, outTokens, label) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TOKEN_SHEET);
    if (!sheet) return;

    const prices = PRICE_LIST[GEMINI_MODEL];
    let cost = 0;

    if (prices) {
      const inputCost = (inTokens / 1000) * prices.input;
      const outputCost = (outTokens / 1000) * prices.output;
      cost = inputCost + outputCost;
    } else {
      Logger.log(`モデル '${GEMINI_MODEL}' の価格がPRICE_LISTに見つかりません。`);
    }

    sheet.appendRow([
      new Date(),
      id,
      inTokens,
      outTokens,
      cost.toFixed(6),
      label,
      GEMINI_MODEL
    ]);
  } catch (e) {
    Logger.log(`Token logging failed: ${e.message}`);
  }
}

