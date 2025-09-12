/**
 * Runs the initial AI evaluation for a new incident.
 */
function runInitialAI(id, subject, details) {
  const prompt = `
    Evaluate the risk for the following incident report.
    Respond strictly in JSON format with these keys:
    - "reason": Why this risk is significant.
    - "plan": A proposed improvement plan.
    - "accident_type": Classification of the accident (e.g., Fall, Struck by object, Fire).
    - "causal_agent": The causal agent (e.g., Machine, Structure, Environment, Human).
    - "scores": An object with numeric scores for:
      - "frequency": (1: Rare, 2: Occasional, 4: Frequent)
      - "likelihood": (1: Unlikely, 2: Possible, 4: Likely, 6: Almost certain)
      - "severity": (1: Minor, 3: Moderate, 6: Serious, 10: Catastrophic)

    ---
    Subject: ${subject}
    Details: ${details}
    ---
  `;
  const result = callGeminiAPI_(id, prompt, 'runInitialAI');
  return {
    reason: result.reason || "Failed to get AI reason.",
    plan: result.plan || "Failed to get AI plan.",
    accident_type: result.accident_type || "AI classification failed.",
    causal_agent: result.causal_agent || "AI causal agent analysis failed.",
    scores: result.scores || { frequency: 1, likelihood: 1, severity: 1 }
  };
}

/**
 * Runs the post-improvement AI evaluation.
 */
function runPostImprovementAI(id, improvement_details) {
  const prompt = `
    Evaluate the following incident improvement report.
    Respond strictly in JSON format with these keys:
    - "comment": AI's commentary on the improvement measures.
    - "scores": An object with new numeric risk scores for:
      - "frequency": (1: Rare, 2: Occasional, 4: Frequent)
      - "likelihood": (1: Unlikely, 2: Possible, 4: Likely, 6: Almost certain)
      - "severity": (1: Minor, 3: Moderate, 6: Serious, 10: Catastrophic)

    ---
    Improvement Details: ${improvement_details}
    ---
  `;
  const result = callGeminiAPI_(id, prompt, 'runPostImprovementAI');
  return {
    comment: result.comment || "Failed to get AI commentary.",
    scores: result.scores || { frequency: 1, likelihood: 1, severity: 1 }
  };
}

/**
 * Generic function to call the Gemini API.
 */
function callGeminiAPI_(id, prompt, functionLabel) {
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${GEMINI_MODEL}:generateContent?key=${API_KEY}`;
  const payload = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: { responseMimeType: "application/json" }
  };
  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const responseBody = response.getContentText();

    if (responseCode === 200) {
      const json = JSON.parse(responseBody);
      const candidate = json.candidates[0];
      const usage = json.usageMetadata || { promptTokenCount: 0, candidatesTokenCount: 0 };
      logTokenUsage(id, usage.promptTokenCount, usage.candidatesTokenCount, functionLabel);
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

