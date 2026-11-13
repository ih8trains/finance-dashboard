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
      \`Во входящих данных отсутствует колонка "\${CONFIG.ID_COLUMN}".\`
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
      \`На листе "\${monthName}" отсутствует колонка "\${CONFIG.ID_COLUMN}".\`
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
            \`=IFERROR(VLOOKUP(\${categoryColumn}\${currentRow};'\${CONFIG.REFERENCE_SHEET}'!A:B;2;FALSE);"Не размечено")\`
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
