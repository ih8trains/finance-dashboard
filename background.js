const OWNER = "ih8trains";
const REPO = "finance-dashboard";
const BRANCH = "main";
const API_BASE = `https://api.github.com/repos/${OWNER}/${REPO}/contents/`;
const TIME_ZONE = "Europe/Moscow";
const PERIOD_DAYS = 30;
const ITERATIONS = 600000;
const STATE_KEY = "tbank_sync_state_v4";
let running = false;

const b64 = bytes => {
  let s = "";
  for (let i = 0; i < bytes.length; i += 0x8000)
    s += String.fromCharCode(...bytes.subarray(i, Math.min(i + 0x8000, bytes.length)));
  return btoa(s);
};
const unb64 = s => Uint8Array.from(atob(s), c => c.charCodeAt(0));
const hex = bytes => [...bytes].map(x => x.toString(16).padStart(2, "0")).join("");
async function sha256(bytes) {
  return hex(new Uint8Array(await crypto.subtle.digest("SHA-256", bytes)));
}
async function deriveKey(pass, salt, usages, iterations = ITERATIONS) {
  const material = await crypto.subtle.importKey("raw", new TextEncoder().encode(pass), "PBKDF2", false, ["deriveKey"]);
  return crypto.subtle.deriveKey(
    { name: "PBKDF2", salt, iterations: Number(iterations) || ITERATIONS, hash: "SHA-256" },
    material, { name: "AES-GCM", length: 256 }, false, usages
  );
}
async function encryptPayload(payload, pass) {
  const plain = new TextEncoder().encode(JSON.stringify(payload));
  const salt = crypto.getRandomValues(new Uint8Array(16));
  const iv = crypto.getRandomValues(new Uint8Array(12));
  const key = await deriveKey(pass, salt, ["encrypt"]);
  const cipher = new Uint8Array(await crypto.subtle.encrypt({ name: "AES-GCM", iv }, key, plain));
  return {
    month: payload.month, count: payload.rows.length, sha256: await sha256(plain),
    algorithm: "AES-GCM", kdf: "PBKDF2-SHA-256", iterations: ITERATIONS,
    salt: b64(salt), iv: b64(iv), ciphertext: b64(cipher)
  };
}
async function decryptEnvelope(envelope, pass, expectedMonth) {
  if (envelope.month !== expectedMonth || envelope.algorithm !== "AES-GCM" ||
      envelope.kdf !== "PBKDF2-SHA-256" || !envelope.salt || !envelope.iv || !envelope.ciphertext)
    throw new Error(`Некорректный формат архива ${expectedMonth}.`);
  const key = await deriveKey(pass, unb64(envelope.salt), ["decrypt"], envelope.iterations);
  let plain;
  try {
    plain = new Uint8Array(await crypto.subtle.decrypt(
      { name: "AES-GCM", iv: unb64(envelope.iv) }, key, unb64(envelope.ciphertext)
    ));
  } catch (error) {
    if (error && error.name === "OperationError") {
      throw new Error(`Не удалось расшифровать архив ${expectedMonth}. Проверь именно пароль архивов дашборда: он должен быть тем же, которым открываются месячные архивы. Архив не изменён.`);
    }
    throw error;
  }
  if (await sha256(plain) !== envelope.sha256) throw new Error(`SHA-256 архива ${expectedMonth} не совпал.`);
  const payload = JSON.parse(new TextDecoder().decode(plain));
  if (payload.month !== expectedMonth || !Array.isArray(payload.headers) ||
      !Array.isArray(payload.rows) || payload.rows.length !== Number(envelope.count))
    throw new Error(`Содержимое архива ${expectedMonth} повреждено.`);
  return payload;
}
function flattenJSON(data) {
  const out = {};
  function walk(value, key) {
    if (value === null || typeof value !== "object") out[key] = value;
    else if (Array.isArray(value)) {
      if (!value.length) out[key] = [];
      value.forEach((v, i) => walk(v, `${key}[${i}]`));
    } else {
      const keys = Object.keys(value);
      if (!keys.length && key) out[key] = {};
      keys.forEach(k => walk(value[k], key ? `${key}.${k}` : k));
    }
  }
  walk(data, "");
  return out;
}
function moscowParts(timestamp) {
  const parts = new Intl.DateTimeFormat("en-CA", {
    timeZone: TIME_ZONE, year: "numeric", month: "2-digit", day: "2-digit",
    hour: "2-digit", minute: "2-digit", second: "2-digit", hourCycle: "h23"
  }).formatToParts(new Date(timestamp));
  return Object.fromEntries(parts.map(p => [p.type, p.value]));
}
function monthKey(timestamp) {
  const p = moscowParts(timestamp);
  return `${p.year}-${p.month}`;
}
function moscowDate(timestamp) {
  const p = moscowParts(timestamp);
  return `${p.day}.${p.month}.${p.year} ${p.hour}:${p.minute}:${p.second}`;
}
function operationRow(op, headers) {
  const flat = flattenJSON(op);
  const id = flat.id;
  const time = Number(flat["operationTime.milliseconds"]);
  if (id === undefined || id === null || String(id) === "" || !Number.isFinite(time) || time <= 0) return null;
  return headers.map(h => {
    if (h === "Моя группа") return "";
    if (h === "ID операции") return String(id);
    if (h === "Дата и время") return moscowDate(time);
    const v = flat[h];
    return v === undefined || v === null ? "" : (typeof v === "object" ? JSON.stringify(v) : String(v));
  });
}
async function getState() {
  const data = await chrome.storage.local.get(STATE_KEY);
  return data[STATE_KEY] || { status: "idle", message: "Готов к синхронизации." };
}
async function setState(patch) {
  const current = await getState();
  const state = { ...current, ...patch, updatedAt: Date.now() };
  await chrome.storage.local.set({ [STATE_KEY]: state });
  chrome.runtime.sendMessage({ type: "SYNC_STATE", state }).catch(() => {});
  return state;
}
async function secrets() {
  // Credentials are intentionally persisted in storage.local by popup.js.
  // Read from the same storage area; storage.session is empty after popup closes/restarts.
  const s = await chrome.storage.local.get(["githubToken", "archivePassword"]);
  if (!s.githubToken || !s.archivePassword) {
    throw new Error("Не найдены сохранённые данные. Открой блок «Доступ к архивам», проверь оба поля и нажми «Синхронизировать».");
  }
  return s;
}
async function githubRequest(path, options = {}) {
  const { githubToken } = await secrets();
  const response = await fetch(API_BASE + path.split("/").map(encodeURIComponent).join("/"), {
    ...options,
    headers: {
      Authorization: `Bearer ${githubToken}`,
      Accept: "application/vnd.github+json",
      "X-GitHub-Api-Version": "2022-11-28",
      ...(options.headers || {})
    }
  });
  if (!response.ok) {
    const body = await response.json().catch(() => ({}));
    throw new Error(`GitHub ${response.status}: ${body.message || response.statusText}`);
  }
  return response.status === 204 ? null : response.json();
}
async function putFile(path, content, message, sha = null) {
  const body = { message, content: btoa(unescape(encodeURIComponent(content))), branch: BRANCH };
  if (sha) body.sha = sha;
  return githubRequest(path, {
    method: "PUT", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body)
  });
}
async function getContentFile(path) {
  try {
    return await githubRequest(path);
  } catch (e) {
    if (String(e.message).includes("GitHub 404")) return null;
    throw e;
  }
}
async function readContentJson(path) {
  const file = await getContentFile(path);
  if (!file || !file.content) throw new Error(`Не удалось прочитать ${path} через GitHub Contents API.`);
  if (file.encoding !== "base64") throw new Error(`Неожиданный формат ответа GitHub для ${path}: ${file.encoding || "без encoding"}.`);
  const encoded = file.content.replace(/\n/g, "");
  const decoded = decodeURIComponent(escape(atob(encoded)));
  try {
    return JSON.parse(decoded);
  } catch (error) {
    throw new Error(`Файл ${path} получен из GitHub, но содержит некорректный JSON: ${error.message}`);
  }
}
async function getBankOrigin() {
  const [tab] = await chrome.tabs.query({ active: true, lastFocusedWindow: true });
  if (!tab?.url) throw new Error("Открой вкладку Т-Банка и авторизуйся.");
  const url = new URL(tab.url);
  if (!/(^|\.)tbank\.ru$|(^|\.)tbank-online\.com$/i.test(url.hostname))
    throw new Error("Активная вкладка не принадлежит Т-Банку.");
  return url.origin;
}
async function fetchOperations() {
  const origin = await getBankOrigin();
  const end = Date.now(), start = end - PERIOD_DAYS * 86400000;
  await setState({ status: "running", stage: "fetching", message: "Получаю операции за последние 30 дней…" });
  const response = await fetch(`${origin}/api/common/v1/operations?start=${start}&end=${end}`, {
    method: "GET", credentials: "include", headers: { Accept: "application/json" }
  });
  if (!response.ok) throw new Error(`Т-Банк вернул HTTP ${response.status}. Проверь авторизацию.`);
  const data = await response.json();
  const ops = data?.payload;
  if (!Array.isArray(ops)) throw new Error("Т-Банк вернул неожиданный формат операций.");
  return ops;
}
async function sync() {
  if (running) return;
  running = true;
  try {
    const { archivePassword } = await secrets();
    const operations = await fetchOperations();
    if (!operations.length) throw new Error("За последние 30 дней банк не вернул операций.");
    await setState({ message: `Получено ${operations.length} операций. Готовлю архивы…` });

    const flatOps = operations.map(flattenJSON);
    const allKeys = new Set();
    flatOps.forEach(o => Object.keys(o).forEach(k => allKeys.add(k)));
    allKeys.delete("id");
    allKeys.delete("operationTime.milliseconds");
    const headers = ["Моя группа", "ID операции", "Дата и время", ...allKeys];

    const incomingByMonth = new Map();
    let invalid = 0;
    for (const op of operations) {
      const flat = flattenJSON(op);
      const id = flat.id;
      const time = Number(flat["operationTime.milliseconds"]);
      if (id === undefined || id === null || !String(id) || !Number.isFinite(time) || time <= 0) { invalid++; continue; }
      const month = monthKey(time);
      const row = operationRow(op, headers);
      if (!incomingByMonth.has(month)) incomingByMonth.set(month, []);
      incomingByMonth.get(month).push(row);
    }
    if (!incomingByMonth.size) throw new Error("Не нашлось корректных операций с ID и датой.");

    await setState({ message: `Обрабатываю ${incomingByMonth.size} месяцев. Проверяю текущие архивы…` });
    const manifestFile = await getContentFile("data/manifest.json");
    if (!manifestFile) throw new Error("В репозитории не найден data/manifest.json.");
    const manifest = JSON.parse(decodeURIComponent(escape(atob(manifestFile.content.replace(/\n/g, "")))));
    if (!Array.isArray(manifest.archives)) throw new Error("manifest.json имеет неизвестную структуру.");

    const manifestByMonth = new Map(manifest.archives.map(a => [a.month, a]));
    let addedTotal = 0, updatedTotal = 0, touched = 0, processedOperations = 0;
    for (const [month, incoming] of [...incomingByMonth.entries()].sort((a,b)=>a[0].localeCompare(b[0]))) {
      const existingMeta = manifestByMonth.get(month);
      let oldPayload = { month, headers: [], rows: [] };
      let archiveSha = null;
      let currentEnvelope = null;
      if (existingMeta) {
        if (existingMeta.status !== "verified") throw new Error(`Архив ${month} не имеет статуса verified. Остановил синхронизацию.`);
        currentEnvelope = await readContentJson(existingMeta.path);
        if (currentEnvelope.sha256 !== existingMeta.sha256 || Number(currentEnvelope.count) !== Number(existingMeta.count))
          throw new Error(`Метаданные архива ${month} не совпадают с manifest. Ничего не перезаписано.`);
        oldPayload = await decryptEnvelope(currentEnvelope, archivePassword, month);
        const contentMeta = await getContentFile(existingMeta.path);
        archiveSha = contentMeta?.sha;
        if (!archiveSha) throw new Error(`Не удалось получить SHA файла ${month}.`);
      }

      const mergedHeaders = [...oldPayload.headers];
      headers.forEach(h => { if (!mergedHeaders.includes(h)) mergedHeaders.push(h); });
      const idIndexOld = oldPayload.headers.indexOf("ID операции");
      if (oldPayload.rows.length && idIndexOld < 0) throw new Error(`В архиве ${month} отсутствует ID операции.`);
      const oldById = new Map();
      oldPayload.rows.forEach(row => {
        const id = String(row[idIndexOld]);
        const normalized = mergedHeaders.map(h => {
          const i = oldPayload.headers.indexOf(h);
          return i >= 0 ? (row[i] ?? "") : "";
        });
        oldById.set(id, normalized);
      });
      let added = 0, updated = 0;
      const incomingIds = new Set();
      for (const newRow of incoming) {
        const id = String(newRow[headers.indexOf("ID операции")]);
        if (incomingIds.has(id)) continue;
        incomingIds.add(id);
        const previous = oldById.get(id);
        if (!previous) added++;
        else updated++;
        const normalized = mergedHeaders.map(h => {
          const newIndex = headers.indexOf(h);
          const oldIndex = oldPayload.headers.indexOf(h);
          if (newIndex >= 0) {
            const value = newRow[newIndex];
            // Не перезаписываем колонку ручной разметки из банковского ответа.
            if (h === "Моя группа" && previous) return previous[oldIndex] ?? "";
            return value ?? "";
          }
          return oldIndex >= 0 && previous ? previous[oldIndex] ?? "" : "";
        });
        oldById.set(id, normalized);
      }
      const rows = [...oldById.values()];
      const payload = { month, headers: mergedHeaders, rows };
      const envelope = await encryptPayload(payload, archivePassword);
      const archivePath = existingMeta?.path || `data/${month}.json.enc`;

      if (existingMeta) {
        const backupPath = `data/backups/sync-${Date.now()}-${month}/${month}.json.enc`;
        await putFile(backupPath, JSON.stringify(currentEnvelope), `Backup ${month} before sync`);
        // Retrieve backup SHA for verification before touching the original.
        const backupContent = await getContentFile(backupPath);
        if (!backupContent) throw new Error(`Не удалось создать резервную копию ${month}.`);
        const backupRaw = await readContentJson(backupPath);
        if (JSON.stringify(backupRaw) !== JSON.stringify(currentEnvelope)) throw new Error(`Резервная копия ${month} не совпала с исходником.`);
        await putFile(archivePath, JSON.stringify(envelope), `Update encrypted operations ${month}`, archiveSha);
        const verify = await readContentJson(archivePath);
        const checked = await decryptEnvelope(verify, archivePassword, month);
        if (checked.rows.length !== rows.length || verify.sha256 !== envelope.sha256) throw new Error(`Проверка нового архива ${month} не пройдена.`);
        existingMeta.count = rows.length;
        existingMeta.sha256 = envelope.sha256;
        existingMeta.salt = envelope.salt;
        existingMeta.iv = envelope.iv;
        existingMeta.iterations = ITERATIONS;
        existingMeta.backupPath = backupPath;
        existingMeta.uploadedAt = new Date().toISOString();
        existingMeta.verifiedAt = new Date().toISOString();
        existingMeta.status = "verified";
      } else {
        await putFile(archivePath, JSON.stringify(envelope), `Create encrypted operations ${month}`);
        const verify = await readContentJson(archivePath);
        const checked = await decryptEnvelope(verify, archivePassword, month);
        if (checked.rows.length !== rows.length) throw new Error(`Проверка нового архива ${month} не пройдена.`);
        manifest.archives.push({
          month, path: archivePath, count: rows.length, sha256: envelope.sha256,
          algorithm: "AES-GCM", kdf: "PBKDF2-SHA-256", iterations: ITERATIONS,
          salt: envelope.salt, iv: envelope.iv, status: "verified", backupPath: "",
          uploadedAt: new Date().toISOString(), verifiedAt: new Date().toISOString()
        });
        manifestByMonth.set(month, manifest.archives[manifest.archives.length - 1]);
      }
      addedTotal += added; updatedTotal += updated; touched++;
      processedOperations += incoming.length;
      manifest.archives.sort((a,b)=>a.month.localeCompare(b.month));
      const latestManifestFile = await getContentFile("data/manifest.json");
      await putFile("data/manifest.json", JSON.stringify(manifest, null, 2),
        `Update manifest after syncing ${month}`, latestManifestFile.sha);
      const manifestCheck = await readContentJson("data/manifest.json");
      const checkedMonth = manifestCheck.archives?.find(x => x.month === month);
      if (!checkedMonth || checkedMonth.sha256 !== envelope.sha256 || Number(checkedMonth.count) !== rows.length)
        throw new Error(`Проверка manifest после ${month} не пройдена.`);
      await setState({
        status: "running", stage: "writing",
        message: `${month}: добавлено ${added}, обновлено ${updated}. Проверено месяцев: ${touched}.`,
        totalOperations: operations.length, sentOperations: processedOperations
      });
    }

    const verifyManifest = await readContentJson("data/manifest.json");
    if (!Array.isArray(verifyManifest.archives) || !incomingByMonth.size)
      throw new Error("Проверка manifest.json не пройдена.");

    await chrome.storage.local.remove(["githubToken", "archivePassword"]);
    await setState({
      status: "success", stage: "done",
      message: `Синхронизация завершена.\nПериод: 30 дней.\nПолучено: ${operations.length}.\nДобавлено: ${addedTotal}.\nОбновлено: ${updatedTotal}.\nНекорректных: ${invalid}.\nМесяцев обновлено: ${touched}.`
    });
  } catch (error) {
    const detail = error && error.name === "OperationError"
      ? "Ошибка криптографии (OperationError). Чаще всего это неверный пароль архивов или несовместимые параметры шифрования. Исходные архивы не должны измениться на этом этапе."
      : (error.message || String(error));
    await setState({ status: "error", stage: "error", message: `Ошибка синхронизации: ${detail}` });
  } finally {
    running = false;
  }
}
chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  if (message?.type === "GET_STATUS") {
    getState().then(state => sendResponse({ state })).catch(e => sendResponse({ error: e.message }));
    return true;
  }
  if (message?.type === "START_SYNC") {
    (async () => {
      if (running) return { ok: false, error: "Синхронизация уже выполняется." };
      const token = String(message.githubToken || "").trim();
      const pass = String(message.archivePassword || "");
      if (!token || !pass) return { ok: false, error: "Нужны GitHub token и пароль архивов." };
      await chrome.storage.local.set({ githubToken: token, archivePassword: pass });
      await setState({ status: "running", stage: "starting", message: "Запускаю синхронизацию…" });
      await sync();
      return { ok: true, state: await getState() };
    })().then(sendResponse).catch(e => sendResponse({ ok: false, error: e.message }));
    return true;
  }
});
