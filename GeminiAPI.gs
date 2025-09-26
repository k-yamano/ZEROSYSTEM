const PRICE_LIST = {
  // Gemini 1.5 Flash
  'gemini-1.5-flash-001': { input: 0.000125, output: 0.000375 },
  'gemini-1.5-flash-latest': { input: 0.000125, output: 0.000375 },
  // Gemini 1.0 Pro
  'gemini-1.0-pro': { input: 0.000125, output: 0.000375 },
  'gemini-1.0-pro-001': { input: 0.000125, output: 0.000375 },
  // 他のモデルも必要に応じてここに追加
};

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
    // ★★★ 最終修正 ★★★
    // contentsオブジェクトに "role": "user" を追加
    payload: JSON.stringify({
      contents: [{
        "role": "user", // この行を追加
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
            logTokens(id, 0, 0, label); 
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

function logTokens(id, inTokens, outTokens, label) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TOKEN_SHEET);
    if (!sheet) return;

    // 現在のモデルの価格を取得
    const prices = PRICE_LIST[GEMINI_MODEL];
    let cost = 0;

    if (prices) {
      // 1000トークンあたりの価格で計算
      const inputCost = (inTokens / 1000) * prices.input;
      const outputCost = (outTokens / 1000) * prices.output;
      cost = inputCost + outputCost;
    } else {
      Logger.log(`モデル '${GEMINI_MODEL}' の価格がPRICE_LISTに見つかりません。`);
    }

    // スプレッドシートに記録
    sheet.appendRow([
      new Date(),       // タイムスタンプ
      id,               // 案件ID
      inTokens,         // 入力トークン数
      outTokens,        // 出力トークン数
      cost.toFixed(6),  // 計算されたコスト（米ドル）
      label,            // 処理ラベル
      GEMINI_MODEL      // 使用したモデル名
    ]);
  } catch (e) {
    Logger.log(`Token logging failed: ${e.message}`);
  }
}