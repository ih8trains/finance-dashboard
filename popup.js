const syncBtn = document.getElementById("syncBtn");
const statusEl = document.getElementById("status");
const progressWrap = document.getElementById("progressWrap");
const progressEl = document.getElementById("progress");

function renderState(state) {
  if (!state) return;
  statusEl.textContent = state.message || "Ожидание...";
  statusEl.style.color = state.status === "error" ? "#dc2626" :
    state.status === "success" ? "#15803d" : "#475569";

  const total = Number(state.totalOperations || 0);
  const sent = Number(state.sentOperations || 0);
  if (state.status === "running" && total > 0) {
    progressWrap.style.display = "block";
    progressEl.style.width = `${Math.min(100, sent / total * 100)}%`;
  } else if (state.status === "success") {
    progressWrap.style.display = "block";
    progressEl.style.width = "100%";
  } else {
    progressWrap.style.display = "none";
    progressEl.style.width = "0%";
  }

  syncBtn.disabled = state.status === "running";
  syncBtn.textContent = state.status === "running" ? "Синхронизация выполняется..." :
    state.status === "error" && state.stage === "sending" ? "Продолжить синхронизацию" :
    "Синхронизировать";
}

const passwordInput = document.getElementById("password");
const tokenInput = document.getElementById("token");

async function restoreCredentials() {
  const saved = await chrome.storage.local.get(["githubToken", "archivePassword"]);
  tokenInput.value = saved.githubToken || "";
  passwordInput.value = saved.archivePassword || "";
}

let saveTimer;
function persistCredentials() {
  clearTimeout(saveTimer);
  saveTimer = setTimeout(() => {
    chrome.storage.local.set({
      githubToken: tokenInput.value.trim(),
      archivePassword: passwordInput.value
    });
  }, 150);
}
[tokenInput, passwordInput].forEach(input => input.addEventListener("input", persistCredentials));

async function refreshStatus() {
  try {
    const response = await chrome.runtime.sendMessage({ type: "GET_STATUS" });
    renderState(response?.state);
  } catch (_) {
    statusEl.textContent = "Не удалось получить статус фоновой задачи.";
  }
}

syncBtn.addEventListener("click", async () => {
  syncBtn.disabled = true;
  syncBtn.textContent = "Запуск...";
  try {
    const response = await chrome.runtime.sendMessage({
      type: "START_SYNC",
      archivePassword: passwordInput.value,
      githubToken: tokenInput.value
    });
    if (!response || !response.ok) throw new Error(response?.error || "Не удалось запустить синхронизацию.");
    renderState(response.state);
    if (response.state?.status === "success") {
      passwordInput.value = "";
      tokenInput.value = "";
    }
  } catch (error) {
    statusEl.textContent = error.message || String(error);
    statusEl.style.color = "#dc2626";
    syncBtn.disabled = false;
    syncBtn.textContent = "Синхронизировать";
  }
});

chrome.runtime.onMessage.addListener(message => {
  if (message?.type === "SYNC_STATE") renderState(message.state);
});

restoreCredentials();
refreshStatus();
setInterval(refreshStatus, 1500);
