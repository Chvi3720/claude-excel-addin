/* global CustomFunctions, Office, fetch */

// This file runs in Excel's hidden custom-functions runtime (loaded by functions.html).
// It reads settings (API key + model) that were saved from the task pane (taskpane.js).
// Those settings live in the workbook itself (Office.context.document.settings), so both
// files see the same values without any extra wiring.

const DEFAULT_MODEL = "claude-sonnet-5";
const API_URL = "https://api.anthropic.com/v1/messages";

const readyPromise = Office.onReady();

async function getSetting(name, fallback) {
  await readyPromise;
  const value = Office.context.document.settings.get(name);
  return value === undefined || value === null || value === "" ? fallback : value;
}

function matrixToText(context) {
  if (context === undefined || context === null) return "";
  if (Array.isArray(context)) {
    return context
      .map((row) => (Array.isArray(row) ? row.join("\t") : String(row)))
      .join("\n");
  }
  return String(context);
}

/**
 * Ask Claude any question or instruction and get the answer back in the cell.
 * @customfunction ASK
 * @param {string} question Your question or instruction, e.g. "summarize this in 5 words".
 * @param {any[][]} [context] Optional: a cell or range to use as data/context for the question.
 * @returns {string} Claude's answer.
 */
async function ASK(question, context) {
  if (!question || !String(question).trim()) {
    throw new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidValue, "Enter a question.");
  }

  const apiKey = await getSetting("claudeApiKey", "");
  if (!apiKey) {
    throw new CustomFunctions.Error(
      CustomFunctions.ErrorCode.invalidValue,
      "No API key set. Click the Claude Settings button on the Home tab and add your key."
    );
  }
  const model = await getSetting("claudeModel", DEFAULT_MODEL);

  const contextText = matrixToText(context);
  const userContent = contextText
    ? `${question}\n\n---\nContext data from the spreadsheet:\n${contextText}`
    : String(question);

  let response;
  try {
    response = await fetch(API_URL, {
      method: "POST",
      headers: {
        "content-type": "application/json",
        "x-api-key": apiKey,
        "anthropic-version": "2023-06-01",
        "anthropic-dangerous-direct-browser-access": "true"
      },
      body: JSON.stringify({
        model: model,
        max_tokens: 1024,
        messages: [{ role: "user", content: userContent }]
      })
    });
  } catch (err) {
    throw new CustomFunctions.Error(CustomFunctions.ErrorCode.notAvailable, "Network error reaching Claude.");
  }

  if (!response.ok) {
    let detail = "";
    try {
      const errJson = await response.json();
      detail = errJson && errJson.error && errJson.error.message ? errJson.error.message : "";
    } catch (e) {
      // ignore parse failure
    }
    if (response.status === 401) {
      throw new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidValue, "Invalid API key.");
    }
    throw new CustomFunctions.Error(
      CustomFunctions.ErrorCode.invalidValue,
      `Claude API error (${response.status})${detail ? ": " + detail : ""}`
    );
  }

  const data = await response.json();
  const text = (data.content || [])
    .map((block) => (block && block.type === "text" ? block.text : ""))
    .join("")
    .trim();

  return text || "(empty response)";
}

CustomFunctions.associate("ASK", ASK);
