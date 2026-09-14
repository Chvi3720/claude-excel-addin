/* global Office, fetch */

const API_URL = "https://api.anthropic.com/v1/messages";

let apiKeyEl, modelEl, saveBtn, saveStatus, quickQuestionEl, askBtn, askStatus, answerBox;

Office.onReady(() => {
  apiKeyEl = document.getElementById("apiKey");
  modelEl = document.getElementById("model");
  saveBtn = document.getElementById("saveBtn");
  saveStatus = document.getElementById("saveStatus");
  quickQuestionEl = document.getElementById("quickQuestion");
  askBtn = document.getElementById("askBtn");
  askStatus = document.getElementById("askStatus");
  answerBox = document.getElementById("answerBox");

  loadSettings();

  saveBtn.addEventListener("click", onSave);
  askBtn.addEventListener("click", onAsk);
});

function loadSettings() {
  const settings = Office.context.document.settings;
  const key = settings.get("claudeApiKey");
  const model = settings.get("claudeModel");
  if (key) apiKeyEl.value = key;
  if (model) modelEl.value = model;
}

function onSave() {
  const settings = Office.context.document.settings;
  settings.set("claudeApiKey", apiKeyEl.value.trim());
  settings.set("claudeModel", modelEl.value);
  saveBtn.disabled = true;
  settings.saveAsync((result) => {
    saveBtn.disabled = false;
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      setStatus(saveStatus, "Saved. You can now use =CLAUDE.ASK(\"question\") in any cell.", "ok");
    } else {
      setStatus(saveStatus, "Could not save settings: " + (result.error ? result.error.message : "unknown error"), "error");
    }
  });
}

async function onAsk() {
  const question = quickQuestionEl.value.trim();
  if (!question) {
    setStatus(askStatus, "Type a question first.", "error");
    return;
  }
  const apiKey = apiKeyEl.value.trim();
  if (!apiKey) {
    setStatus(askStatus, "Add your API key above and click Save first.", "error");
    return;
  }
  const model = modelEl.value;

  askBtn.disabled = true;
  answerBox.style.display = "none";
  setStatus(askStatus, "Asking Claude...", "");

  try {
    const response = await fetch(API_URL, {
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
        messages: [{ role: "user", content: question }]
      })
    });

    if (!response.ok) {
      let detail = "";
      try {
        const errJson = await response.json();
        detail = errJson && errJson.error && errJson.error.message ? errJson.error.message : "";
      } catch (e) {
        // ignore
      }
      setStatus(askStatus, `Error (${response.status})${detail ? ": " + detail : ""}`, "error");
      return;
    }

    const data = await response.json();
    const text = (data.content || [])
      .map((block) => (block && block.type === "text" ? block.text : ""))
      .join("")
      .trim();

    answerBox.textContent = text || "(empty response)";
    answerBox.style.display = "block";
    setStatus(askStatus, "", "");
  } catch (err) {
    setStatus(askStatus, "Network error reaching Claude.", "error");
  } finally {
    askBtn.disabled = false;
  }
}

function setStatus(el, message, kind) {
  el.textContent = message;
  el.className = "status" + (kind ? " " + kind : "");
}
