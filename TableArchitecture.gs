/**
 * Семантический слой доступа к листу «База».
 *
 * Главный инвариант: бизнес-логика обращается к полям по стабильным field_key,
 * а не по буквам/номерам столбцов. Физическое расположение определяется из
 * скрытого листа «Схема» и строки заголовков «База» при первом обращении.
 */
const BASE_SHEET_NAME = 'База';
const BASE_SCHEMA_SHEET_NAME = 'Схема';
const SPEC_RELATIONS_SHEET_NAME = 'Связи спецификации';

const REQUIRED_BASE_FIELDS = Object.freeze([
  'date', 'drawing_link', 'order_link', 'uuid', 'item_name', 'destination',
  'note', 'quantity', 'stock', 'deficit', 'need_status', 'need_flag',
  'stage_design', 'stage_technical_control', 'stage_deployment',
  'stage_technologist', 'stage_technology_check', 'stage_picker',
  'stage_picking', 'stage_stock', 'stage_production_stock', 'stage_production',
  'stage_quality_check', 'stage_acceptance', 'stage_ready', 'status',
  'production_mode', 'salesdrive_id', 'trello_id', 'salesdrive_order_id',
  'actual_made', 'responsible_technologist', 'designer', 'supply_ordered',
  'tied_capital', 'salesdrive_stock', 'abc', 'abc_components',
  'abc_components_2', 'priority', 'tooling', 'deadline', 'cost',
  'should_order', 'supplier', 'order_produced_received'
]);

const STAGE_FIELD_KEYS = Object.freeze([
  'stage_design',
  'stage_technical_control',
  'stage_deployment',
  'stage_technologist',
  'stage_technology_check',
  'stage_picker',
  'stage_picking',
  'stage_stock',
  'stage_production_stock',
  'stage_production',
  'stage_quality_check',
  'stage_acceptance',
  'stage_ready'
]);

let _baseArchitectureCache = new Map();
let _dispatcherRuntimeSpreadsheet = null;

function setDispatcherRuntimeSpreadsheet_(ss) {
  _dispatcherRuntimeSpreadsheet = ss || null;
}

function getDispatcherRuntimeSpreadsheet_() {
  return _dispatcherRuntimeSpreadsheet || SpreadsheetApp.getActiveSpreadsheet();
}

function clearTableArchitectureCache_(ss) {
  if (!ss) {
    _baseArchitectureCache = new Map();
    return;
  }
  _baseArchitectureCache.delete(ss.getId());
}

/** Читает «Схема» и сопоставляет field_key с текущим физическим столбцом «База». */
function getBaseArchitecture_(ss) {
  ss = ss || getDispatcherRuntimeSpreadsheet_();
  if (!ss) throw new Error('Не удалось определить таблицу Диспетчера.');

  const cacheKey = ss.getId();
  if (_baseArchitectureCache.has(cacheKey)) return _baseArchitectureCache.get(cacheKey);

  const schemaSheet = ss.getSheetByName(BASE_SCHEMA_SHEET_NAME);
  const baseSheet = ss.getSheetByName(BASE_SHEET_NAME);
  if (!schemaSheet) throw new Error(`Не найден служебный лист «${BASE_SCHEMA_SHEET_NAME}».`);
  if (!baseSheet) throw new Error(`Не найден лист «${BASE_SHEET_NAME}».`);

  const schemaLastRow = schemaSheet.getLastRow();
  if (schemaLastRow < 2) throw new Error('Лист «Схема» не содержит описания полей.');

  // A=field_key, B=header, C=kind, D=computed.
  const schemaRows = schemaSheet.getRange(2, 1, schemaLastRow - 1, 4).getValues();
  const keyToHeader = new Map();
  const headerToKey = new Map();
  const computedKeys = new Set();

  for (const row of schemaRows) {
    const key = String(row[0] ?? '').trim();
    const header = String(row[1] ?? '').trim();
    if (!key && !header) continue;
    if (!key || !header) throw new Error(`Некорректная строка «Схема»: key/header должны быть заполнены вместе.`);
    if (keyToHeader.has(key)) throw new Error(`Дублирующий field_key «${key}» на листе «Схема».`);
    if (headerToKey.has(header)) throw new Error(`Дублирующий заголовок «${header}» на листе «Схема».`);

    keyToHeader.set(key, header);
    headerToKey.set(header, key);
    if (row[3] === true || String(row[3]).toUpperCase() === 'TRUE') computedKeys.add(key);
  }

  const lastColumn = baseSheet.getLastColumn();
  const headers = baseSheet.getRange(1, 1, 1, lastColumn).getDisplayValues()[0]
    .map(value => String(value ?? '').trim());
  const headerToColumn = new Map();

  headers.forEach((header, index) => {
    if (!header) return;
    if (headerToColumn.has(header)) {
      throw new Error(`В «База» найден повторяющийся заголовок «${header}».`);
    }
    headerToColumn.set(header, index + 1);
  });

  const keyToColumn = new Map();
  const columnToKey = new Map();
  for (const [key, header] of keyToHeader) {
    const column = headerToColumn.get(header);
    if (!column) continue;
    keyToColumn.set(key, column);
    columnToKey.set(column, key);
  }

  const missing = REQUIRED_BASE_FIELDS.filter(key => !keyToColumn.has(key));
  if (missing.length) {
    throw new Error(`В «База» отсутствуют обязательные поля: ${missing.join(', ')}`);
  }

  const architecture = Object.freeze({
    spreadsheetId: cacheKey,
    baseSheet,
    schemaSheet,
    keyToHeader,
    keyToColumn,
    columnToKey,
    computedKeys,
    lastColumn
  });
  _baseArchitectureCache.set(cacheKey, architecture);
  return architecture;
}

function baseCol_(ss, fieldKey) {
  const architecture = getBaseArchitecture_(ss);
  const column = architecture.keyToColumn.get(fieldKey);
  if (!column) throw new Error(`Поле «${fieldKey}» не найдено в семантической схеме «База».`);
  return column;
}

function baseHeader_(ss, fieldKey) {
  const header = getBaseArchitecture_(ss).keyToHeader.get(fieldKey);
  if (!header) throw new Error(`field_key «${fieldKey}» отсутствует в «Схема».`);
  return header;
}

function baseKeyForColumn_(ss, column) {
  return getBaseArchitecture_(ss).columnToKey.get(Number(column)) || '';
}

function isComputedBaseField_(ss, fieldKey) {
  return getBaseArchitecture_(ss).computedKeys.has(fieldKey);
}

/**
 * Читает произвольный набор семантических полей одной строки одним физическим
 * чтением диапазона min..max. Результат — объект {fieldKey: value}.
 */
function readBaseFields_(baseSheet, row, fieldKeys) {
  if (!baseSheet || !fieldKeys || !fieldKeys.length) return {};
  const ss = baseSheet.getParent();
  const columns = fieldKeys.map(key => ({ key, col: baseCol_(ss, key) }));
  const minCol = Math.min(...columns.map(item => item.col));
  const maxCol = Math.max(...columns.map(item => item.col));
  const values = baseSheet.getRange(row, minCol, 1, maxCol - minCol + 1).getValues()[0];
  const result = {};
  for (const item of columns) result[item.key] = values[item.col - minCol];
  return result;
}

/** Группирует отсортированные физические столбцы в непрерывные блоки. */
function groupContiguousColumns_(items) {
  if (!items.length) return [];
  const sorted = [...items].sort((a, b) => a.col - b.col);
  const groups = [];
  let current = [sorted[0]];

  for (let i = 1; i < sorted.length; i++) {
    if (sorted[i].col === current[current.length - 1].col + 1) {
      current.push(sorted[i]);
    } else {
      groups.push(current);
      current = [sorted[i]];
    }
  }
  groups.push(current);
  return groups;
}

/**
 * Записывает несколько полей одной строки. Вычисляемые spill-поля защищены от
 * случайной записи; их можно менять только явно через allowComputed=true.
 */
function writeBaseFields_(baseSheet, row, valuesByKey, options = {}) {
  if (!baseSheet || !valuesByKey) return;
  const ss = baseSheet.getParent();
  const allowComputed = options.allowComputed === true;
  const items = Object.keys(valuesByKey).map(key => {
    if (!allowComputed && isComputedBaseField_(ss, key)) {
      throw new Error(`Попытка записать вычисляемое поле «${key}». Оно обслуживается spill-формулой.`);
    }
    return { key, col: baseCol_(ss, key), value: valuesByKey[key] };
  });

  for (const group of groupContiguousColumns_(items)) {
    baseSheet
      .getRange(row, group[0].col, 1, group.length)
      .setValues([group.map(item => item.value)]);
  }
}

/**
 * Пакетно записывает новые строки как массив объектов {field_key: value}.
 * Вычисляемые поля не трогаются: их значения разливаются из формул строки 2.
 */
function writeBaseRows_(baseSheet, startRow, rows) {
  if (!baseSheet || !rows || !rows.length) return;
  const ss = baseSheet.getParent();
  const neededEndRow = startRow + rows.length - 1;
  const maxRows = baseSheet.getMaxRows();
  if (neededEndRow > maxRows) baseSheet.insertRowsAfter(maxRows, neededEndRow - maxRows);

  const keySet = new Set();
  rows.forEach(row => Object.keys(row || {}).forEach(key => keySet.add(key)));
  const items = [...keySet].map(key => {
    if (isComputedBaseField_(ss, key)) {
      throw new Error(`Новая строка содержит вычисляемое поле «${key}». Уберите его из данных записи.`);
    }
    return { key, col: baseCol_(ss, key) };
  });

  for (const group of groupContiguousColumns_(items)) {
    const matrix = rows.map(row => group.map(item =>
      Object.prototype.hasOwnProperty.call(row, item.key) ? row[item.key] : ''
    ));
    baseSheet.getRange(startRow, group[0].col, rows.length, group.length).setValues(matrix);
  }
}

function findBaseRowByUuid_(baseSheet, itemId) {
  if (!baseSheet || itemId === null || typeof itemId === 'undefined') return null;
  const target = String(itemId).trim();
  if (!target) return null;

  const ss = baseSheet.getParent();
  const uuidCol = baseCol_(ss, 'uuid');
  const lastRow = baseLastDataRow_(baseSheet);
  if (lastRow < 2) return null;

  const match = baseSheet
    .getRange(2, uuidCol, lastRow - 1, 1)
    .createTextFinder(target)
    .matchEntireCell(true)
    .findNext();
  return match ? match.getRow() : null;
}

function baseLastDataRow_(baseSheet) {
  if (!baseSheet) return 0;
  const ss = baseSheet.getParent();
  const uuidCol = baseCol_(ss, 'uuid');
  const maxRows = baseSheet.getMaxRows();
  if (maxRows < 2) return 1;
  return baseSheet
    .getRange(maxRows, uuidCol)
    .getNextDataCell(SpreadsheetApp.Direction.UP)
    .getRow();
}

/**
 * Внутренняя проверка архитектуры. Публичный диагностический wrapper находится
 * в Diagnostics.gs, чтобы все ручные проверки оставались собраны в одном месте.
 */
function validateTableArchitecture_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  setDispatcherRuntimeSpreadsheet_(ss);
  clearTableArchitectureCache_(ss);
  const architecture = getBaseArchitecture_(ss);
  const issues = [];

  // В вычисляемых столбцах должна быть одна формула-якорь в строке 2 и не должно
  // быть независимых формул ниже — это защищает от formula drift.
  const lastDataRow = baseLastDataRow_(architecture.baseSheet);
  for (const key of architecture.computedKeys) {
    const col = architecture.keyToColumn.get(key);
    if (!col) continue;
    if (!architecture.baseSheet.getRange(2, col).getFormula()) {
      issues.push(`Нет формулы-якоря: ${key} (${architecture.keyToHeader.get(key)})`);
      continue;
    }
    if (lastDataRow >= 3) {
      const formulasBelow = architecture.baseSheet.getRange(3, col, lastDataRow - 2, 1).getFormulas();
      if (formulasBelow.some(row => Boolean(row[0]))) {
        issues.push(`Найдены построчные формулы ниже якоря: ${key}`);
      }
    }
  }

  const settingsSheet = ss.getSheetByName('Настройки');
  if (!settingsSheet) {
    issues.push('Нет листа «Настройки».');
  } else {
    const rows = settingsSheet.getRange(2, 1, Math.max(settingsSheet.getLastRow() - 1, 1), 2).getValues();
    for (const [sheetNameRaw, fieldKeyRaw] of rows) {
      const sheetName = String(sheetNameRaw ?? '').trim();
      const fieldKey = String(fieldKeyRaw ?? '').trim();
      if (!sheetName && !fieldKey) continue;
      if (!sheetName || !fieldKey) issues.push(`Неполная настройка этапа: «${sheetName}» / «${fieldKey}».`);
      else if (!architecture.keyToColumn.has(fieldKey)) issues.push(`Этап «${sheetName}»: неизвестный field_key «${fieldKey}».`);
    }
  }

  const relationSheet = ss.getSheetByName(SPEC_RELATIONS_SHEET_NAME);
  if (!relationSheet) issues.push(`Нет листа «${SPEC_RELATIONS_SHEET_NAME}».`);

  const result = {
    ok: issues.length === 0,
    spreadsheetId: ss.getId(),
    mappedFields: architecture.keyToColumn.size,
    computedFields: architecture.computedKeys.size,
    lastDataRow,
    issues
  };
  Logger.log(JSON.stringify(result, null, 2));
  return result;
}
