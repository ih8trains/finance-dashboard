/**
 * Архив операций — серверная часть, этап 1.
 *
 * Этот файл добавляется в тот же Apps Script-проект, что и Code.gs.
 * Здесь нет паролей и токенов: GitHub PAT читается из Script Properties.
 *
 * Действия:
 *   archive_list_candidates — список закрытых месяцев-кандидатов;
 *   archive_get_month       — выгрузка исходных строк месяца;
 *   archive_upload          — публикация уже зашифрованного браузером архива.
 *
 * Удаления строк из Google Sheets здесь намеренно НЕТ.
 */

const ARCHIVE_CONFIG = {
  GITHUB_OWNER: "ih8trains",
  GITHUB_REPO: "finance-dashboard",
  GITHUB_BRANCH: "main",
  DATA_DIR: "data",
  MANIFEST_PATH: "data/manifest.json",
  GITHUB_API: "https://api.github.com",
  MIN_AGE_DAYS: 30,
  MONTHS_RU: [
    "Январь", "Февраль", "Март", "Апрель",
    "Май", "Июнь", "Июль", "Август",
    "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"
  ]
};


/**
 * Вставить в doPost(e) из Code.gs перед обработчиком по умолчанию:
 *
 * if (String(request.action || "").indexOf("archive_") === 0) {
 *   return jsonResponse(handleArchiveAction_(request));
 * }
 */
function handleArchiveAction_(request) {
  switch (request.action) {
    case "archive_list_candidates":
      return archiveListCandidates_();

    case "archive_get_month":
      return archiveGetMonth_(request.month);

    case "archive_upload":
      return archiveUpload_(request);

    default:
      throw new Error("Неизвестное действие архива.");
  }
}


function archiveListCandidates_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const now = new Date();
  const cutoff = new Date(
    now.getTime() - ARCHIVE_CONFIG.MIN_AGE_DAYS * 24 * 60 * 60 * 1000
  );

  const manifest = archiveReadManifest_();
  const archived = new Set(
    (manifest.archives || [])
      .filter(item => item && item.status === "ready_unverified")
      .map(item => item.month)
  );

  const candidates = [];

  ss.getSheets().forEach(sheet => {
    const parsed = archiveParseMonthName_(sheet.getName());
    if (!parsed) return;

    const monthEnd = new Date(
      parsed.year,
      parsed.monthIndex + 1,
      0,
      23, 59, 59, 999
    );

    if (monthEnd > cutoff) return;

    const lastRow = sheet.getLastRow();
    const lastColumn = sheet.getLastColumn();
    if (lastRow < 2 || lastColumn < 1) return;

    const headers = sheet.getRange(1, 1, 1, lastColumn)
      .getDisplayValues()[0].map(String);

    if (!headers.includes(CONFIG.ID_COLUMN)) return;

    const count = lastRow - 1;
    candidates.push({
      month: sheet.getName(),
      monthKey: parsed.year + "-" + String(parsed.monthIndex + 1).padStart(2, "0"),
      rows: count,
      alreadyUploaded: archived.has(
        parsed.year + "-" + String(parsed.monthIndex + 1).padStart(2, "0")
      )
    });
  });

  candidates.sort((a, b) => a.monthKey.localeCompare(b.monthKey));

  return {
    status: "success",
    cutoff: cutoff.toISOString(),
    candidates: candidates
  };
}


function archiveGetMonth_(monthName) {
  if (typeof monthName !== "string" || !archiveParseMonthName_(monthName)) {
    throw new Error("Некорректное название месяца.");
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(monthName);
  if (!sheet) throw new Error("Лист месяца не найден.");

  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();

  if (lastRow < 2 || lastColumn < 1) {
    throw new Error("В месяце нет операций для архивации.");
  }

  const values = sheet.getRange(1, 1, lastRow, lastColumn).getValues();
  const headers = values[0].map(String);

  if (!headers.includes(CONFIG.ID_COLUMN)) {
    throw new Error("На листе отсутствует колонка ID операции.");
  }

  const rows = values.slice(1).filter(row =>
    row.some(value => value !== "" && value !== null)
  ).map(row => row.map(archiveSerializeValue_));

  return {
    status: "success",
    month: monthName,
    monthKey: archiveMonthKey_(monthName),
    headers: headers,
    rows: rows,
    rowCount: rows.length
  };
}


/**
 * Принимает только зашифрованный клиентом payload.
 * Само шифрование и пароль архива остаются в браузере.
 */
function archiveUpload_(request) {
  const token = PropertiesService.getScriptProperties()
    .getProperty("GITHUB_TOKEN");

  if (!token) {
    throw new Error("В свойствах скрипта не найден GITHUB_TOKEN.");
  }

  const monthName = String(request.month || "");
  const monthKey = archiveMonthKey_(monthName);

  if (!monthKey) throw new Error("Некорректный месяц архива.");

  const parsed = archiveParseMonthName_(monthName);
  const monthEnd = new Date(parsed.year, parsed.monthIndex + 1, 0, 23, 59, 59, 999);
  const cutoff = new Date(Date.now() - ARCHIVE_CONFIG.MIN_AGE_DAYS * 86400000);
  if (monthEnd > cutoff) {
    throw new Error("Этот месяц ещё не вышел за безопасное 30-дневное окно.");
  }

  const encrypted = String(request.ciphertext || "");
  const salt = String(request.salt || "");
  const iv = String(request.iv || "");
  const sha256 = String(request.sha256 || "");
  const iterations = Number(request.iterations || 0);
  const rowCount = Number(request.rowCount || 0);

  if (!encrypted || !salt || !iv || !/^[a-f0-9]{64}$/i.test(sha256)) {
    throw new Error("Не хватает данных зашифрованного архива.");
  }
  if (!Number.isInteger(iterations) || iterations < 100000 || iterations > 2000000) {
    throw new Error("Некорректное число итераций PBKDF2.");
  }
  if (!Number.isInteger(rowCount) || rowCount < 1) {
    throw new Error("Некорректное количество строк.");
  }

  const envelope = {
    version: 1,
    algorithm: "AES-GCM",
    kdf: "PBKDF2-SHA-256",
    iterations: iterations,
    salt: salt,
    iv: iv,
    sha256: sha256.toLowerCase(),
    rowCount: rowCount,
    ciphertext: encrypted
  };

  const archivePath = ARCHIVE_CONFIG.DATA_DIR + "/" + monthKey + ".json.enc";
  const archiveText = JSON.stringify(envelope);

  const uploaded = githubPutFile_(
    token,
    archivePath,
    archiveText,
    "Archive encrypted operations " + monthKey
  );

  const manifest = archiveReadManifest_();
  const archives = Array.isArray(manifest.archives) ? manifest.archives : [];
  const record = {
    month: monthKey,
    sheetName: monthName,
    path: archivePath,
    rows: rowCount,
    sha256: sha256.toLowerCase(),
    status: "ready_unverified",
    uploadedAt: new Date().toISOString(),
    commit: uploaded.commit
  };

  const oldIndex = archives.findIndex(item => item && item.month === monthKey);
  if (oldIndex >= 0) archives[oldIndex] = record;
  else archives.push(record);

  archives.sort((a, b) => String(a.month).localeCompare(String(b.month)));

  const newManifest = {
    version: 1,
    updatedAt: new Date().toISOString(),
    archives: archives
  };

  githubPutFile_(
    token,
    ARCHIVE_CONFIG.MANIFEST_PATH,
    JSON.stringify(newManifest, null, 2),
    "Update archive manifest " + monthKey
  );

  return {
    status: "success",
    month: monthKey,
    path: archivePath,
    rows: rowCount,
    sha256: sha256.toLowerCase(),
    archiveCommit: uploaded.commit,
    verified: false,
    deletedFromSheets: false,
    message: "Архив опубликован. Проверь его расшифровку в браузере; строки Google Sheets не удалялись."
  };
}


function archiveReadManifest_() {
  const token = PropertiesService.getScriptProperties().getProperty("GITHUB_TOKEN");
  if (!token) throw new Error("В свойствах скрипта не найден GITHUB_TOKEN.");

  const result = githubGetFile_(token, ARCHIVE_CONFIG.MANIFEST_PATH);
  if (!result) return { version: 1, archives: [] };

  try {
    return JSON.parse(result.content);
  } catch (error) {
    throw new Error("Файл data/manifest.json повреждён или содержит некорректный JSON.");
  }
}


function githubGetFile_(token, path) {
  const url = ARCHIVE_CONFIG.GITHUB_API + "/repos/" +
    ARCHIVE_CONFIG.GITHUB_OWNER + "/" + ARCHIVE_CONFIG.GITHUB_REPO +
    "/contents/" + path + "?ref=" + encodeURIComponent(ARCHIVE_CONFIG.GITHUB_BRANCH);

  const response = UrlFetchApp.fetch(url, {
    method: "get",
    muteHttpExceptions: true,
    headers: {
      Authorization: "Bearer " + token,
      Accept: "application/vnd.github+json",
      "X-GitHub-Api-Version": "2022-11-28"
    }
  });

  const code = response.getResponseCode();
  if (code === 404) return null;
  if (code < 200 || code >= 300) {
    throw new Error("GitHub read error " + code + ": " + response.getContentText().slice(0, 500));
  }

  const json = JSON.parse(response.getContentText());
  const bytes = Utilities.base64Decode(String(json.content || "").replace(/\s/g, ""));
  return {
    sha: json.sha,
    content: Utilities.newBlob(bytes).getDataAsString("UTF-8")
  };
}


function githubPutFile_(token, path, content, message) {
  const current = githubGetFile_(token, path);
  const bytes = Utilities.newBlob(content, "application/json").getBytes();
  const base64 = Utilities.base64Encode(bytes);

  const body = {
    message: message,
    content: base64,
    branch: ARCHIVE_CONFIG.GITHUB_BRANCH
  };
  if (current && current.sha) body.sha = current.sha;

  const url = ARCHIVE_CONFIG.GITHUB_API + "/repos/" +
    ARCHIVE_CONFIG.GITHUB_OWNER + "/" + ARCHIVE_CONFIG.GITHUB_REPO +
    "/contents/" + path;

  const response = UrlFetchApp.fetch(url, {
    method: "put",
    contentType: "application/json",
    payload: JSON.stringify(body),
    muteHttpExceptions: true,
    headers: {
      Authorization: "Bearer " + token,
      Accept: "application/vnd.github+json",
      "X-GitHub-Api-Version": "2022-11-28"
    }
  });

  const code = response.getResponseCode();
  if (code < 200 || code >= 300) {
    throw new Error("GitHub write error " + code + ": " + response.getContentText().slice(0, 700));
  }

  const json = JSON.parse(response.getContentText());
  return {
    commit: json.commit && json.commit.sha ? json.commit.sha : ""
  };
}


function archiveParseMonthName_(name) {
  const match = String(name || "").match(/^(.+)\s+(20\d{2})$/);
  if (!match) return null;

  const monthIndex = ARCHIVE_CONFIG.MONTHS_RU.indexOf(match[1]);
  if (monthIndex < 0) return null;

  return {
    year: Number(match[2]),
    monthIndex: monthIndex
  };
}


function archiveMonthKey_(monthName) {
  const parsed = archiveParseMonthName_(monthName);
  if (!parsed) return "";
  return parsed.year + "-" + String(parsed.monthIndex + 1).padStart(2, "0");
}


function archiveSerializeValue_(value) {
  if (value instanceof Date) return value.toISOString();
  if (value === null || value === undefined) return "";
  return value;
}
