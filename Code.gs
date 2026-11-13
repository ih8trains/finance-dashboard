const SECRET_PASS = "";

const CONFIG = {
  MANUAL_COLUMNS: [
    "Моя группа"
  ],

  ID_COLUMN: "ID операции",
  DATE_COLUMN: "Дата и время",
  CATEGORY_COLUMN: "spendingCategory.name",

  REFERENCE_SHEET: "Справочник",

  INSERT_BATCH_SIZE: 500
};


// =====================================================
// ОСНОВНОЙ ОБРАБОТЧИК
// =====================================================

function doPost(e) {
  try {
    if (
      !SECRET_PASS ||
      SECRET_PASS === "ВСТАВЬ_НОВЫЙ_СЕКРЕТ"
    ) {
      throw new Error("Не задан пароль Apps Script.");
    }

    if (!e || !e.postData || !e.postData.contents) {
      throw new Error("Пустой запрос.");
    }

    const request = JSON.parse(e.postData.contents);

    if (request.password !== SECRET_PASS) {
      throw new Error("Неверный пароль.");
    }

    if (request.action === "sync_batch") {
      if (
        !request.month ||
        !Array.isArray(request.headers) ||
        !Array.isArray(request.rows)
      ) {
        throw new Error("Некорректный пакет.");
      }

      const lock = LockService.getScriptLock();

      if (!lock.tryLock(30000)) {
        throw new Error(
          "Таблица занята другой синхронизацией. Повтори пакет."
        );
      }

      try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();

        const result = syncMonth(
          ss,
          request.month,
          request.headers,
          request.rows
        );

        SpreadsheetApp.flush();

        return jsonResponse({
          status: "success",
          message: "Пакет записан.",
          added: result.added,
          updated: result.updated,
          skipped: result.skipped
        });

      } finally {
        lock.releaseLock();
      }
    }

    if (request.action === "sync_raw") {
      const result = syncRawOperations(
        request.data || {}
      );

      return jsonResponse(result);
    }

    if (request.action === "get_meta") {
      return jsonResponse(getDashboardMeta());
    }

    if (request.action === "save_label") {
      return jsonResponse(
        saveOperationLabel(request.label || {})
      );
    }

    if (request.action === "save_labels_batch") {
      return jsonResponse(
        saveOperationLabelsBatch(request.labels || [])
      );
    }

    if (request.action === "save_settings") {
      return jsonResponse(
        saveDashboardSettings(request.settings || {})
      );
    }

    if (request.action === "delete_rule") {
      return jsonResponse(
        deleteDashboardRule(request.id)
      );
    }

    if (String(request.action || "").indexOf("archive_") === 0) {
      return jsonResponse(handleArchiveAction_(request));
    }

    return getDashboardData();

  } catch (error) {
    return jsonResponse({
      status: "error",
      message: error.message || String(error)
    });
  }
}


// =====================================================
// JSON RESPONSE
// =====================================================

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}


// =====================================================
// СТАРАЯ СИНХРОНИЗАЦИЯ
// =====================================================

function syncRawOperations(dataByMonth) {
  const lock = LockService.getScriptLock();

  if (!lock.tryLock(30000)) {
    throw new Error("Таблица занята другой синхронизацией.");
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    let totalAdded = 0;
    let totalUpdated = 0;
    let totalSkipped = 0;

    const monthResults = {};

    for (const month in dataByMonth) {
      const monthData = dataByMonth[month];

      if (
        !monthData ||
        !Array.isArray(monthData.headers) ||
        !Array.isArray(monthData.rows)
      ) {
        continue;
      }

      const result = syncMonth(
        ss,
        month,
        monthData.headers,
        monthData.rows
      );

      totalAdded += result.added;
      totalUpdated += result.updated;
      totalSkipped += result.skipped;

      monthResults[month] = result;
    }

    SpreadsheetApp.flush();

    return {
      status: "success",
      message: "Синхронизация завершена",
      added: totalAdded,
      updated: totalUpdated,
      skipped: totalSkipped,
      months: monthResults
    };

  } finally {
    lock.releaseLock();
  }
}


// =====================================================
// СИНХРОНИЗАЦИЯ МЕСЯЦА
// =====================================================

function syncMonth(
  ss,
  monthName,
  incomingHeaders,
  incomingRows
) {
  if (!incomingHeaders.includes(CONFIG.ID_COLUMN)) {
    throw new Error(
      `Во входящих данных отсутствует колонка "${CONFIG.ID_COLUMN}".`
    );
  }

  let sheet = ss.getSheetByName(monthName);

  if (!sheet) {
    sheet = ss.insertSheet(monthName);

    sheet.getRange(
      1,
      1,
      1,
      incomingHeaders.length
    ).setValues([incomingHeaders]);

    sheet.getRange(
      1,
      1,
      1,
      incomingHeaders.length
    )
    .setFontWeight("bold")
    .setHorizontalAlignment("center");
  }

  let lastColumn = Math.max(
    sheet.getLastColumn(),
    incomingHeaders.length,
    1
  );

  let existingHeaders = sheet
    .getRange(1, 1, 1, lastColumn)
    .getValues()[0]
    .map(value => String(value));

  if (!existingHeaders.includes(CONFIG.ID_COLUMN)) {
    throw new Error(
      `На листе "${monthName}" отсутствует колонка "${CONFIG.ID_COLUMN}".`
    );
  }

  const newHeaders = incomingHeaders.filter(
    header => !existingHeaders.includes(header)
  );

  if (newHeaders.length > 0) {
    const startColumn = existingHeaders.length + 1;

    sheet.getRange(
      1,
      startColumn,
      1,
      newHeaders.length
    ).setValues([newHeaders]);

    sheet.getRange(
      1,
      startColumn,
      1,
      newHeaders.length
    )
    .setFontWeight("bold")
    .setHorizontalAlignment("center");

    existingHeaders = existingHeaders.concat(newHeaders);
  }

  const columnCount = existingHeaders.length;

  const incomingIndex = {};

  incomingHeaders.forEach((header, index) => {
    incomingIndex[header] = index;
  });

  const existingIndex = {};

  existingHeaders.forEach((header, index) => {
    existingIndex[header] = index;
  });

  const idColumnIndex = existingIndex[CONFIG.ID_COLUMN];

  const lastRow = sheet.getLastRow();
  const dataRowCount = Math.max(0, lastRow - 1);

  const existingData = dataRowCount > 0
    ? sheet.getRange(
        2,
        1,
        dataRowCount,
        columnCount
      ).getValues()
    : [];

  const existingIds = new Map();

  existingData.forEach((row, index) => {
    const id = String(
      row[idColumnIndex] || ""
    ).trim();

    if (id && !existingIds.has(id)) {
      existingIds.set(id, index + 2);
    }
  });

  const manualColumns = new Set(
    CONFIG.MANUAL_COLUMNS
  );

  const processedIds = new Set();
  const changedRows = new Map();
  const newOperations = [];

  let skipped = 0;

  for (const sourceRow of incomingRows) {
    if (!Array.isArray(sourceRow)) {
      skipped++;
      continue;
    }

    const id = String(
      sourceRow[incomingIndex[CONFIG.ID_COLUMN]] || ""
    ).trim();

    if (!id || processedIds.has(id)) {
      skipped++;
      continue;
    }

    processedIds.add(id);

    const existingSheetRow = existingIds.get(id);

    if (existingSheetRow) {
      const dataIndex = existingSheetRow - 2;

      let targetRow = changedRows.get(
        existingSheetRow
      );

      if (!targetRow) {
        targetRow = existingData[dataIndex].slice();
      }

      incomingHeaders.forEach((header, sourceIndex) => {
        if (manualColumns.has(header)) return;

        const targetIndex = existingIndex[header];

        if (targetIndex === undefined) return;

        targetRow[targetIndex] =
          sourceRow[sourceIndex] !== undefined
            ? sourceRow[sourceIndex]
            : "";
      });

      changedRows.set(
        existingSheetRow,
        targetRow
      );

    } else {
      const newRow = new Array(columnCount).fill("");

      incomingHeaders.forEach((header, sourceIndex) => {
        if (manualColumns.has(header)) return;

        const targetIndex = existingIndex[header];

        if (targetIndex === undefined) return;

        newRow[targetIndex] =
          sourceRow[sourceIndex] !== undefined
            ? sourceRow[sourceIndex]
            : "";
      });

      newOperations.push(newRow);
    }
  }

  // Обновление существующих операций

  if (changedRows.size > 0) {
    const affectedRows = Array.from(
      changedRows.keys()
    ).sort((a, b) => a - b);

    const minRow = affectedRows[0];
    const maxRow = affectedRows[affectedRows.length - 1];

    const writableColumnRuns = [];

    let runStart = null;

    for (let col = 0; col < columnCount; col++) {
      const isManual = manualColumns.has(
        existingHeaders[col]
      );

      if (!isManual && runStart === null) {
        runStart = col;
      }

      if (
        (isManual || col === columnCount - 1) &&
        runStart !== null
      ) {
        const runEnd = isManual
          ? col - 1
          : col;

        writableColumnRuns.push([
          runStart,
          runEnd
        ]);

        runStart = null;
      }
    }

    for (const [startCol, endCol] of writableColumnRuns) {
      const matrix = [];

      for (
        let sheetRow = minRow;
        sheetRow <= maxRow;
        sheetRow++
      ) {
        const changed = changedRows.get(sheetRow);

        const source = changed ||
          existingData[sheetRow - 2];

        matrix.push(
          source.slice(startCol, endCol + 1)
        );
      }

      sheet.getRange(
        minRow,
        startCol + 1,
        matrix.length,
        endCol - startCol + 1
      ).setValues(matrix);
    }
  }

  // Добавление новых операций

  let added = 0;

  if (newOperations.length > 0) {
    const firstNewRow = sheet.getLastRow() + 1;

    for (
      let i = 0;
      i < newOperations.length;
      i += CONFIG.INSERT_BATCH_SIZE
    ) {
      const batch = newOperations.slice(
        i,
        i + CONFIG.INSERT_BATCH_SIZE
      );

      sheet.getRange(
        firstNewRow + i,
        1,
        batch.length,
        columnCount
      ).setValues(batch);

      added += batch.length;
    }

    const groupIndex = existingIndex["Моя группа"];

    const categoryIndex = existingIndex[
      CONFIG.CATEGORY_COLUMN
    ];

    if (
      groupIndex !== undefined &&
      categoryIndex !== undefined
    ) {
      const groupColumn = groupIndex + 1;

      const categoryColumn = columnToLetter(
        categoryIndex + 1
      );

      const formulas = newOperations.map(
        (row, index) => {
          const currentRow = firstNewRow + index;

          return [
            `=IFERROR(VLOOKUP(${categoryColumn}${currentRow};'${CONFIG.REFERENCE_SHEET}'!A:B;2;FALSE);"Не размечено")`
          ];
        }
      );

      sheet.getRange(
        firstNewRow,
        groupColumn,
        formulas.length,
        1
      ).setFormulas(formulas);
    }
  }

  return {
    added: added,
    updated: changedRows.size,
    skipped: skipped
  };
}


// =====================================================
// НОМЕР КОЛОНКИ → БУКВА
// =====================================================

function columnToLetter(column) {
  let letter = "";

  while (column > 0) {
    const remainder = (column - 1) % 26;

    letter = String.fromCharCode(
      65 + remainder
    ) + letter;

    column = Math.floor(
      (column - 1) / 26
    );
  }

  return letter;
}


// =====================================================
// ДАННЫЕ ДЛЯ ДАШБОРДА
// =====================================================

function getDashboardData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();

  let allHeaders = [];
  const allData = [];

  // Каждый лист читается только один раз

  for (const sheet of sheets) {
    const sheetName = sheet.getName();

    if (
      sheetName === CONFIG.REFERENCE_SHEET ||
      Object.values(META_SHEETS).includes(sheetName)
    ) {
      continue;
    }

    const values = sheet.getDataRange().getValues();

    if (values.length <= 1 || values[0].length === 0) {
      continue;
    }

    const headers = values[0].map(value => String(value));

    if (
      !headers.includes(CONFIG.ID_COLUMN) ||
      !headers.includes("Моя группа")
    ) {
      continue;
    }

    headers.forEach(header => {
      if (!allHeaders.includes(header)) {
        allHeaders.push(header);
      }
    });

    allData.push({
      headers: headers,
      rows: values.slice(1)
    });
  }

  if (allData.length === 0) {
    return jsonResponse({
      status: "error",
      message: "Не найдено листов с выгрузкой Т-Банка."
    });
  }

  // Объединение операций в единую структуру

  const normalizedData = [];

  for (const item of allData) {
    const indexMap = allHeaders.map(
      header => item.headers.indexOf(header)
    );

    for (const row of item.rows) {
      normalizedData.push(
        indexMap.map(index =>
          index === -1 ? "" : row[index]
        )
      );
    }
  }

  // Метаданные

  const meta = getDashboardMeta();

  return jsonResponse({
    status: "success",
    headers: allHeaders,
    data: normalizedData,
    labels: meta.labels || [],
    settings: meta.settings || {},
    rules: meta.rules || []
  });
}


// =====================================================
// ПОЛЬЗОВАТЕЛЬСКИЕ ДАННЫЕ ДАШБОРДА
// =====================================================

const META_SHEETS = {
  LABELS: "UserLabels",
  SETTINGS: "DashboardSettings",
  RULES: "DashboardRules"
};


function ensureMetaSheet_(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let sheet = ss.getSheetByName(name);

  if (!sheet) {
    sheet = ss.insertSheet(name);
  }

  if (sheet.getLastRow() === 0) {
    sheet.getRange(
      1,
      1,
      1,
      headers.length
    ).setValues([headers]);
  }

  return sheet;
}


function getDashboardMeta() {
  const labelsSheet = ensureMetaSheet_(
    META_SHEETS.LABELS,
    [
      "operation_id",
      "custom_category",
      "tags",
      "spending_scope",
      "operation_role",
      "related_person",
      "project",
      "reimbursable",
      "reimbursed_amount",
      "comment",
      "updated_at"
    ]
  );

  const settingsSheet = ensureMetaSheet_(
    META_SHEETS.SETTINGS,
    [
      "key",
      "value",
      "updated_at"
    ]
  );

  const rulesSheet = ensureMetaSheet_(
    META_SHEETS.RULES,
    [
      "id",
      "field",
      "match",
      "value",
      "enabled",
      "updated_at"
    ]
  );

  const labels = readObjects_(labelsSheet);
  const settingsRows = readObjects_(settingsSheet);

  const rules = readObjects_(rulesSheet).filter(
    r => String(r.enabled).toLowerCase() !== "false"
  );

  const settings = {};

  settingsRows.forEach(r => {
    if (r.key) {
      try {
        settings[r.key] = JSON.parse(r.value);
      } catch (e) {
        settings[r.key] = r.value;
      }
    }
  });

  return {
    status: "success",
    labels: labels,
    settings: settings,
    rules: rules
  };
}


function readObjects_(sheet) {
  const values = sheet.getDataRange().getValues();

  if (values.length < 2) return [];

  const headers = values[0].map(String);

  return values
    .slice(1)
    .filter(row =>
      row.some(v => v !== "" && v !== null)
    )
    .map(row => {
      const obj = {};

      headers.forEach((h, i) => {
        obj[h] = row[i] instanceof Date
          ? row[i].toISOString()
          : row[i];
      });

      return obj;
    });
}


// =====================================================
// СОХРАНЕНИЕ РАЗМЕТКИ
// =====================================================

function saveOperationLabel(label) {
  if (!label.operation_id) {
    throw new Error("Не указан ID операции.");
  }

  const sheet = ensureMetaSheet_(
    META_SHEETS.LABELS,
    [
      "operation_id",
      "custom_category",
      "tags",
      "spending_scope",
      "operation_role",
      "related_person",
      "project",
      "reimbursable",
      "reimbursed_amount",
      "comment",
      "updated_at"
    ]
  );

  const headers = sheet
    .getRange(1, 1, 1, sheet.getLastColumn())
    .getValues()[0]
    .map(String);

  const idCol = headers.indexOf("operation_id") + 1;
  const last = sheet.getLastRow();

  let rowNum = 0;

  if (last > 1) {
    const ids = sheet
      .getRange(2, idCol, last - 1, 1)
      .getDisplayValues()
      .flat();

    const idx = ids.indexOf(String(label.operation_id));

    if (idx >= 0) {
      rowNum = idx + 2;
    }
  }

  const obj = {
    operation_id: String(label.operation_id),
    custom_category: String(label.custom_category || ""),
    tags: Array.isArray(label.tags)
      ? label.tags.join(",")
      : String(label.tags || ""),
    spending_scope: String(label.spending_scope || "personal"),
    operation_role: String(label.operation_role || ""),
    related_person: String(label.related_person || ""),
    project: String(label.project || ""),
    reimbursable: String(
      label.reimbursable === true ||
      label.reimbursable === "true"
    ),
    reimbursed_amount: Number(label.reimbursed_amount || 0),
    comment: String(label.comment || ""),
    updated_at: new Date()
  };

  const row = headers.map(
    h => obj[h] !== undefined ? obj[h] : ""
  );

  if (rowNum) {
    sheet.getRange(
      rowNum,
      1,
      1,
      headers.length
    ).setValues([row]);
  } else {
    sheet.appendRow(row);
  }

  return {
    status: "success",
    operation_id: obj.operation_id
  };
}


function saveOperationLabelsBatch(labels) {
  if (!Array.isArray(labels) || labels.length > 500) {
    throw new Error(
      "Пакет разметки должен содержать до 500 операций."
    );
  }

  if (labels.length === 0) {
    return {
      status: "success",
      saved: 0
    };
  }

  const sheet = ensureMetaSheet_(
    META_SHEETS.LABELS,
    [
      "operation_id",
      "custom_category",
      "tags",
      "spending_scope",
      "operation_role",
      "related_person",
      "project",
      "reimbursable",
      "reimbursed_amount",
      "comment",
      "updated_at"
    ]
  );

  const headers = sheet
    .getRange(1, 1, 1, sheet.getLastColumn())
    .getValues()[0]
    .map(String);

  const idIndex = headers.indexOf("operation_id");

  if (idIndex < 0) {
    throw new Error(
      "В UserLabels отсутствует колонка operation_id."
    );
  }

  const lastRow = sheet.getLastRow();

  const existing = lastRow > 1
    ? sheet.getRange(
        2,
        1,
        lastRow - 1,
        headers.length
      ).getValues()
    : [];

  const rowById = new Map();

  existing.forEach((row, index) => {
    rowById.set(String(row[idIndex]), index);
  });

  const newRows = [];
  const now = new Date();

  labels.forEach(label => {
    if (!label || !label.operation_id) return;

    const obj = {
      operation_id: String(label.operation_id),
      custom_category: String(label.custom_category || ""),
      tags: Array.isArray(label.tags)
        ? label.tags.join(",")
        : String(label.tags || ""),
      spending_scope: String(label.spending_scope || "personal"),
      operation_role: String(label.operation_role || ""),
      related_person: String(label.related_person || ""),
      project: String(label.project || ""),
      reimbursable: String(
        label.reimbursable === true ||
        label.reimbursable === "true"
      ),
      reimbursed_amount: Number(label.reimbursed_amount || 0),
      comment: String(label.comment || ""),
      updated_at: now
    };

    const row = headers.map(
      header => obj[header] !== undefined
        ? obj[header]
        : ""
    );

    const existingIndex = rowById.get(obj.operation_id);

    if (existingIndex !== undefined) {
      if (existingIndex < existing.length) {
        existing[existingIndex] = row;
      } else {
        newRows[existingIndex - existing.length] = row;
      }
    } else {
      rowById.set(
        obj.operation_id,
        existing.length + newRows.length
      );

      newRows.push(row);
    }
  });

  if (existing.length) {
    sheet.getRange(
      2,
      1,
      existing.length,
      headers.length
    ).setValues(existing);
  }

  if (newRows.length) {
    sheet.getRange(
      sheet.getLastRow() + 1,
      1,
      newRows.length,
      headers.length
    ).setValues(newRows);
  }

  return {
    status: "success",
    saved: labels.length
  };
}


// =====================================================
// СОХРАНЕНИЕ НАСТРОЕК
// =====================================================

function saveDashboardSettings(settings) {
  const sheet = ensureMetaSheet_(
    META_SHEETS.SETTINGS,
    [
      "key",
      "value",
      "updated_at"
    ]
  );

  Object.keys(settings).forEach(key => {
    const value = JSON.stringify(settings[key]);
    const last = sheet.getLastRow();

    let rowNum = 0;

    if (last > 1) {
      const keys = sheet
        .getRange(2, 1, last - 1, 1)
        .getDisplayValues()
        .flat();

      const idx = keys.indexOf(key);

      if (idx >= 0) {
        rowNum = idx + 2;
      }
    }

    const row = [
      key,
      value,
      new Date()
    ];

    if (rowNum) {
      sheet.getRange(
        rowNum,
        1,
        1,
        3
      ).setValues([row]);
    } else {
      sheet.appendRow(row);
    }
  });

  return {
    status: "success"
  };
}


// =====================================================
// УДАЛЕНИЕ ПРАВИЛ
// =====================================================

function deleteDashboardRule(id) {
  const sheet = ensureMetaSheet_(
    META_SHEETS.RULES,
    [
      "id",
      "field",
      "match",
      "value",
      "enabled",
      "updated_at"
    ]
  );

  const last = sheet.getLastRow();

  if (last < 2) {
    return {
      status: "success",
      deleted: false
    };
  }

  const ids = sheet
    .getRange(2, 1, last - 1, 1)
    .getDisplayValues()
    .flat();

  const idx = ids.indexOf(String(id));

  if (idx >= 0) {
    sheet.deleteRow(idx + 2);
  }

  return {
    status: "success",
    deleted: idx >= 0
  };
}


// =====================================================
// ЗАЩИЩЁННОЕ АРХИВИРОВАНИЕ В GITHUB
// Первый этап: выгрузка, шифрование в браузере,
// загрузка и проверка. Удаления из Sheets здесь нет.
// =====================================================

const ARCHIVE_CONFIG_ = {
  REPOSITORY: "ih8trains/finance-dashboard",
  BRANCH: "main",
  MANIFEST_PATH: "data/manifest.json",
  ARCHIVE_FOLDER: "data",
  ROLLING_DAYS: 30,
  STATEMENT_DAY: 19
};


function handleArchiveAction_(request) {
  const action = String(request.action || "");

  if (action === "archive_list_candidates") {
    return archiveListCandidates_();
  }

  if (action === "archive_get_month") {
    return archiveGetMonth_(String(request.month || ""));
  }

  if (action === "archive_upload") {
    return archiveUpload_(request);
  }

  if (action === "archive_mark_verified") {
    return archiveMarkVerified_(request);
  }

  if (action === "archive_get_manifest") {
    return {
      status: "success",
      manifest: archiveReadManifest_()
    };
  }

  throw new Error("Неизвестная команда архивирования: " + action);
}


function archiveListCandidates_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const now = new Date();
  const cutoff = new Date(now.getTime());
  cutoff.setDate(cutoff.getDate() - ARCHIVE_CONFIG_.ROLLING_DAYS);

  const manifest = archiveReadManifest_();
  const verifiedMonths = new Set(
    manifest.archives
      .filter(item => item.status === "verified")
      .map(item => item.month)
  );

  const candidates = [];

  sheets.forEach(sheet => {
    const sheetName = sheet.getName();

    if (
      sheetName === CONFIG.REFERENCE_SHEET ||
      Object.values(META_SHEETS).includes(sheetName)
    ) {
      return;
    }

    const monthInfo = archiveParseSheetMonth_(sheetName);
    if (!monthInfo) return;
    if (verifiedMonths.has(monthInfo.key)) return;

    const monthEnd = new Date(
      monthInfo.year,
      monthInfo.monthIndex + 1,
      0,
      23, 59, 59
    );

    const statementDate = new Date(
      monthInfo.year,
      monthInfo.monthIndex + 1,
      ARCHIVE_CONFIG_.STATEMENT_DAY,
      23, 59, 59
    );

    // Архивируем только после 19-го числа следующего месяца
    // и после выхода всего месяца за rolling 30-day window.
    if (now < statementDate || monthEnd > cutoff) return;

    const lastRow = sheet.getLastRow();
    const lastColumn = sheet.getLastColumn();

    if (lastRow < 2 || lastColumn < 1) return;

    const headers = sheet
      .getRange(1, 1, 1, lastColumn)
      .getValues()[0]
      .map(value => String(value));

    const idIndex = headers.indexOf(CONFIG.ID_COLUMN);
    if (idIndex < 0) return;

    const values = sheet
      .getRange(2, 1, lastRow - 1, lastColumn)
      .getValues();

    const count = values.filter(row =>
      String(row[idIndex] || "").trim() !== ""
    ).length;

    if (count > 0) {
      candidates.push({
        month: monthInfo.key,
        sheetName: sheetName,
        count: count,
        rowsCount: count
      });
    }
  });

  candidates.sort((a, b) => a.month.localeCompare(b.month));

  return {
    status: "success",
    cutoff: Utilities.formatDate(
      cutoff,
      Session.getScriptTimeZone(),
      "yyyy-MM-dd"
    ),
    candidates: candidates
  };
}


function archiveParseSheetMonth_(sheetName) {
  const normalized = String(sheetName || "")
    .trim()
    .toLowerCase()
    .replace(/ё/g, "е");

  const numeric = normalized.match(/^(\d{4})-(\d{2})$/);

  if (numeric) {
    const year = Number(numeric[1]);
    const monthNumber = Number(numeric[2]);

    if (monthNumber >= 1 && monthNumber <= 12) {
      return {
        year: year,
        monthIndex: monthNumber - 1,
        key: numeric[1] + "-" + numeric[2]
      };
    }
  }

  const monthNames = {
    "январь": 0,
    "февраль": 1,
    "март": 2,
    "апрель": 3,
    "май": 4,
    "июнь": 5,
    "июль": 6,
    "август": 7,
    "сентябрь": 8,
    "октябрь": 9,
    "ноябрь": 10,
    "декабрь": 11
  };

  const match = normalized.match(/^([а-я]+)\s+(\d{4})$/);
  if (!match) return null;

  const monthIndex = monthNames[match[1]];
  const year = Number(match[2]);

  if (monthIndex === undefined || !year) return null;

  return {
    year: year,
    monthIndex: monthIndex,
    key: year + "-" + String(monthIndex + 1).padStart(2, "0")
  };
}


function archiveGetMonth_(monthKey) {
  if (!/^\d{4}-\d{2}$/.test(monthKey)) {
    throw new Error("Некорректный месяц архива.");
  }

  const candidateResult = archiveListCandidates_();
  const candidate = candidateResult.candidates.find(
    item => item.month === monthKey
  );

  if (!candidate) {
    throw new Error(
      "Месяц " + monthKey +
      " пока нельзя архивировать: он не прошёл условия безопасности."
    );
  }

  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(candidate.sheetName);

  if (!sheet) {
    throw new Error("Лист месяца не найден: " + candidate.sheetName);
  }

  const values = sheet.getDataRange().getValues();

  if (values.length < 2) {
    throw new Error("В листе нет операций для архивации.");
  }

  const headers = values[0].map(value => String(value));
  const idIndex = headers.indexOf(CONFIG.ID_COLUMN);

  if (idIndex < 0) {
    throw new Error("В листе отсутствует колонка ID операции.");
  }

  const rows = values.slice(1).filter(row =>
    String(row[idIndex] || "").trim() !== ""
  );

  const expectedCount = Number(candidate.count ?? candidate.rowsCount);
  if (Number.isFinite(expectedCount) && expectedCount >= 0 && rows.length !== expectedCount) {
    throw new Error(
      "Количество операций изменилось во время подготовки архива. " +
      "В списке: " + expectedCount + ", фактически: " + rows.length + ". " +
      "Повтори получение месяца."
    );
  }

  return {
    status: "success",
    payload: {
      schemaVersion: 1,
      month: monthKey,
      sheetName: candidate.sheetName,
      exportedAt: new Date().toISOString(),
      headers: headers,
      rows: rows,
      count: rows.length
    }
  };
}


function archiveUpload_(request) {
  const month = String(request.month || "");
  const envelope = request.envelope;

  if (!/^\d{4}-\d{2}$/.test(month)) {
    throw new Error("Некорректный месяц архива.");
  }

  if (
    !envelope ||
    envelope.algorithm !== "AES-GCM" ||
    envelope.kdf !== "PBKDF2-SHA-256" ||
    !envelope.salt ||
    !envelope.iv ||
    !envelope.ciphertext ||
    !envelope.sha256 ||
    !Number.isFinite(Number(envelope.count))
  ) {
    throw new Error("Некорректный зашифрованный архив.");
  }

  if (Number(envelope.count) <= 0) {
    throw new Error("Пустой архив загружать нельзя.");
  }

  const path = ARCHIVE_CONFIG_.ARCHIVE_FOLDER +
    "/" + month + ".json.enc";

  const fileContent = JSON.stringify({
    schemaVersion: 1,
    month: month,
    algorithm: envelope.algorithm,
    kdf: envelope.kdf,
    iterations: Number(envelope.iterations),
    salt: String(envelope.salt),
    iv: String(envelope.iv),
    sha256: String(envelope.sha256),
    count: Number(envelope.count),
    ciphertext: String(envelope.ciphertext)
  });

  const existing = githubGetFile_(path);

  githubPutFile_(
    path,
    fileContent,
    "Add encrypted finance archive " + month,
    existing ? existing.sha : null
  );

  const manifest = archiveReadManifest_();
  const item = {
    month: month,
    path: path,
    count: Number(envelope.count),
    sha256: String(envelope.sha256),
    algorithm: envelope.algorithm,
    kdf: envelope.kdf,
    iterations: Number(envelope.iterations),
    salt: String(envelope.salt),
    iv: String(envelope.iv),
    status: "uploaded",
    uploadedAt: new Date().toISOString()
  };

  const existingIndex = manifest.archives.findIndex(
    entry => entry.month === month
  );

  if (existingIndex >= 0) {
    manifest.archives[existingIndex] = item;
  } else {
    manifest.archives.push(item);
  }

  manifest.archives.sort((a, b) =>
    a.month.localeCompare(b.month)
  );

  githubWriteManifest_(manifest);

  return {
    status: "success",
    month: month,
    path: path,
    count: item.count,
    sha256: item.sha256,
    archiveUrl:
      "https://raw.githubusercontent.com/" +
      ARCHIVE_CONFIG_.REPOSITORY + "/" +
      ARCHIVE_CONFIG_.BRANCH + "/" + path
  };
}


function archiveMarkVerified_(request) {
  const month = String(request.month || "");
  const sha256 = String(request.sha256 || "");
  const count = Number(request.count);

  const manifest = archiveReadManifest_();
  const item = manifest.archives.find(
    entry => entry.month === month
  );

  if (!item) {
    throw new Error("Архив месяца не найден в manifest.");
  }

  if (
    item.sha256 !== sha256 ||
    Number(item.count) !== count
  ) {
    throw new Error(
      "Проверка не совпала с метаданными загруженного архива."
    );
  }

  item.status = "verified";
  item.verifiedAt = new Date().toISOString();

  githubWriteManifest_(manifest);

  return {
    status: "success",
    month: month,
    count: count,
    archiveStatus: "verified"
  };
}


function archiveReadManifest_() {
  const file = githubGetFile_(ARCHIVE_CONFIG_.MANIFEST_PATH);

  if (!file) {
    return {
      schemaVersion: 1,
      archives: []
    };
  }

  let manifest;

  try {
    manifest = JSON.parse(file.content);
  } catch (error) {
    throw new Error("Файл data/manifest.json содержит некорректный JSON.");
  }

  if (!Array.isArray(manifest.archives)) {
    manifest.archives = [];
  }

  return manifest;
}


function githubGetFile_(path) {
  const token = PropertiesService
    .getScriptProperties()
    .getProperty("GITHUB_TOKEN");

  if (!token) {
    throw new Error(
      "Не найден GITHUB_TOKEN в свойствах скрипта."
    );
  }

  const url =
    "https://api.github.com/repos/" +
    ARCHIVE_CONFIG_.REPOSITORY +
    "/contents/" +
    path.split("/").map(encodeURIComponent).join("/") +
    "?ref=" + encodeURIComponent(ARCHIVE_CONFIG_.BRANCH);

  const response = UrlFetchApp.fetch(url, {
    method: "get",
    headers: {
      Authorization: "Bearer " + token,
      Accept: "application/vnd.github+json",
      "X-GitHub-Api-Version": "2022-11-28"
    },
    muteHttpExceptions: true
  });

  const code = response.getResponseCode();

  if (code === 404) return null;

  if (code < 200 || code >= 300) {
    throw new Error(
      "GitHub GET " + path + " вернул HTTP " + code +
      ": " + response.getContentText().slice(0, 500)
    );
  }

  const result = JSON.parse(response.getContentText());
  const encoded = String(result.content || "").replace(/\s/g, "");
  const bytes = Utilities.base64Decode(encoded);

  return {
    sha: result.sha,
    content: Utilities.newBlob(bytes).getDataAsString("UTF-8")
  };
}


function githubPutFile_(path, content, message, sha) {
  const token = PropertiesService
    .getScriptProperties()
    .getProperty("GITHUB_TOKEN");

  if (!token) {
    throw new Error(
      "Не найден GITHUB_TOKEN в свойствах скрипта."
    );
  }

  const url =
    "https://api.github.com/repos/" +
    ARCHIVE_CONFIG_.REPOSITORY +
    "/contents/" +
    path.split("/").map(encodeURIComponent).join("/");

  const payload = {
    message: message,
    content: Utilities.base64Encode(
      content,
      Utilities.Charset.UTF_8
    ),
    branch: ARCHIVE_CONFIG_.BRANCH
  };

  if (sha) payload.sha = sha;

  const response = UrlFetchApp.fetch(url, {
    method: "put",
    contentType: "application/json",
    headers: {
      Authorization: "Bearer " + token,
      Accept: "application/vnd.github+json",
      "X-GitHub-Api-Version": "2022-11-28"
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const code = response.getResponseCode();

  if (code < 200 || code >= 300) {
    throw new Error(
      "GitHub PUT " + path + " вернул HTTP " + code +
      ": " + response.getContentText().slice(0, 700)
    );
  }

  return JSON.parse(response.getContentText());
}


function githubWriteManifest_(manifest) {
  const path = ARCHIVE_CONFIG_.MANIFEST_PATH;
  const existing = githubGetFile_(path);

  githubPutFile_(
    path,
    JSON.stringify(manifest, null, 2),
    "Update encrypted finance archive manifest",
    existing ? existing.sha : null
  );
}
