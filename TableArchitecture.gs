/**
 * Семантический слой доступа к листу «База».
 * Бизнес-логика использует стабильные field_key, физический столбец определяется
 * по скрытому листу «Схема» и текущим заголовкам «База».
 */
const BASE_SHEET_NAME = 'База';
const BASE_SCHEMA_SHEET_NAME = 'Схема';

const REQUIRED_BASE_FIELDS = Object.freeze([
  'date', 'drawing_link', 'order_link', 'uuid', 'item_name', 'destination',
  'note', 'quantity', 'minimum_batch', 'order_quantity', 'stock', 'deficit',
  'need_status', 'need_flag', 'manual_status',
  'stage_design', 'stage_technical_control', 'stage_deployment',
  'stage_technologist', 'stage_technology_check', 'stage_picker',
  'stage_picking', 'stage_stock', 'stage_production_stock', 'stage_production',
  'stage_quality_check', 'stage_acceptance', 'stage_ready', 'status',
  'production_mode', 'supplier', 'supplier_request_date',
  'salesdrive_id', 'trello_id', 'salesdrive_order_id', 'actual_made',
  'responsible_technologist', 'designer', 'supply_ordered', 'tied_capital',
  'salesdrive_stock', 'abc', 'abc_components', 'abc_components_2', 'priority',
  'tooling', 'deadline', 'cost', 'should_order', 'order_produced_received',
  'parent_uuid', 'root_uuid', 'parent_name', 'root_name'
]);

const STAGE_FIELD_KEYS = Object.freeze([
  'stage_design', 'stage_technical_control', 'stage_deployment',
  'stage_technologist', 'stage_technology_check', 'stage_picker',
  'stage_picking', 'stage_stock', 'stage_production_stock', 'stage_production',
  'stage_quality_check', 'stage_acceptance', 'stage_ready'
]);

let _baseArchitectureCache = new Map();
let _dispatcherRuntimeSpreadsheet = null;

function setDispatcherRuntimeSpreadsheet_(ss) { _dispatcherRuntimeSpreadsheet = ss || null; }
function getDispatcherRuntimeSpreadsheet_() { return _dispatcherRuntimeSpreadsheet || SpreadsheetApp.getActiveSpreadsheet(); }

function clearTableArchitectureCache_(ss) {
  if (!ss) { _baseArchitectureCache = new Map(); return; }
  _baseArchitectureCache.delete(ss.getId());
}

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

  const schemaRows = schemaSheet.getRange(2, 1, schemaLastRow - 1, 4).getValues();
  const keyToHeader = new Map();
  const headerToKey = new Map();
  const computedKeys = new Set();
  for (const row of schemaRows) {
    const key = String(row[0] ?? '').trim();
    const header = String(row[1] ?? '').trim();
    if (!key && !header) continue;
    if (!key || !header) throw new Error('Некорректная строка «Схема»: key/header должны быть заполнены вместе.');
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
    if (headerToColumn.has(header)) throw new Error(`В «База» найден повторяющийся заголовок «${header}».`);
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
  if (missing.length) throw new Error(`В «База» отсутствуют обязательные поля: ${missing.join(', ')}`);

  const architecture = Object.freeze({
    spreadsheetId: cacheKey, baseSheet, schemaSheet, keyToHeader, keyToColumn,
    columnToKey, computedKeys, lastColumn
  });
  _baseArchitectureCache.set(cacheKey, architecture);
  return architecture;
}

function baseCol_(ss, fieldKey) {
  const column = getBaseArchitecture_(ss).keyToColumn.get(fieldKey);
  if (!column) throw new Error(`Поле «${fieldKey}» не найдено в семантической схеме «База».`);
  return column;
}
function baseHeader_(ss, fieldKey) {
  const header = getBaseArchitecture_(ss).keyToHeader.get(fieldKey);
  if (!header) throw new Error(`field_key «${fieldKey}» отсутствует в «Схема».`);
  return header;
}
function baseKeyForColumn_(ss, column) { return getBaseArchitecture_(ss).columnToKey.get(Number(column)) || ''; }
function isComputedBaseField_(ss, fieldKey) { return getBaseArchitecture_(ss).computedKeys.has(fieldKey); }

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

function groupContiguousColumns_(items) {
  if (!items.length) return [];
  const sorted = [...items].sort((a, b) => a.col - b.col);
  const groups = [];
  let current = [sorted[0]];
  for (let i = 1; i < sorted.length; i++) {
    if (sorted[i].col === current[current.length - 1].col + 1) current.push(sorted[i]);
    else { groups.push(current); current = [sorted[i]]; }
  }
  groups.push(current);
  return groups;
}

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
    baseSheet.getRange(row, group[0].col, 1, group.length)
      .setValues([group.map(item => item.value)]);
  }
}

function resolveParentRootUuid_(baseSheet, parentUuid, pendingRootByUuid) {
  if (!parentUuid) return '';
  if (pendingRootByUuid.has(parentUuid)) return pendingRootByUuid.get(parentUuid);
  const parentRow = findBaseRowByUuid_(baseSheet, parentUuid);
  if (!parentRow) return parentUuid;
  const parent = readBaseFields_(baseSheet, parentRow, ['uuid', 'root_uuid']);
  return String(parent.root_uuid || parent.uuid || parentUuid).trim();
}

/**
 * Variant B: одна строка = одно применение компонента у одного родителя.
 * Новые генераторы задают parent_uuid напрямую. _specRelation поддерживается
 * только как вход старой версии bound-кода при атомарной миграции файлов.
 */
function prepareBaseHierarchyRows_(baseSheet, rows) {
  const pendingRootByUuid = new Map();
  for (const row of rows) {
    if (!row) continue;
    const uuid = String(row.uuid || '').trim();
    if (!uuid) continue;
    const relation = row._specRelation || null;
    const relationParent = relation ? String(relation.parentUuid || '').trim() : '';
    const parentUuid = String(row.parent_uuid || relationParent || '').trim();
    let rootUuid = String(row.root_uuid || '').trim();

    if (!rootUuid) {
      rootUuid = parentUuid
        ? resolveParentRootUuid_(baseSheet, parentUuid, pendingRootByUuid)
        : uuid;
    }

    row.parent_uuid = parentUuid;
    row.root_uuid = rootUuid || parentUuid || uuid;
    pendingRootByUuid.set(uuid, row.root_uuid);
  }
}

/** Пакетно записывает новые строки; вычисляемые поля не материализуются. */
function writeBaseRows_(baseSheet, startRow, rows) {
  if (!baseSheet || !rows || !rows.length) return;
  const ss = baseSheet.getParent();
  const neededEndRow = startRow + rows.length - 1;
  const maxRows = baseSheet.getMaxRows();
  if (neededEndRow > maxRows) baseSheet.insertRowsAfter(maxRows, neededEndRow - maxRows);

  prepareBaseHierarchyRows_(baseSheet, rows);

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
  const uuidCol = baseCol_(baseSheet.getParent(), 'uuid');
  const lastRow = baseLastDataRow_(baseSheet);
  if (lastRow < 2) return null;
  const match = baseSheet.getRange(2, uuidCol, lastRow - 1, 1)
    .createTextFinder(target).matchEntireCell(true).findNext();
  return match ? match.getRow() : null;
}

/** Возвращает последнюю реально занятую строку «База» по UUID. */
function baseLastDataRow_(baseSheet) {
  if (!baseSheet) return 0;
  const upperBound = Math.max(baseSheet.getLastRow(), 1);
  if (upperBound < 2) return 1;

  const uuidCol = baseCol_(baseSheet.getParent(), 'uuid');
  const values = baseSheet.getRange(2, uuidCol, upperBound - 1, 1).getDisplayValues();
  for (let i = values.length - 1; i >= 0; i--) {
    if (String(values[i][0] || '').trim()) return i + 2;
  }
  return 1;
}

function findHierarchyCycles_(uuids, parents) {
  const parentByUuid = new Map();
  uuids.forEach((rawUuid, i) => {
    const uuid = String(rawUuid || '').trim();
    if (uuid) parentByUuid.set(uuid, String(parents[i] || '').trim());
  });

  const cycleNodes = new Set();
  for (const start of parentByUuid.keys()) {
    const seenAt = new Map();
    const path = [];
    let current = start;
    while (current && parentByUuid.has(current)) {
      if (seenAt.has(current)) {
        for (let i = seenAt.get(current); i < path.length; i++) cycleNodes.add(path[i]);
        cycleNodes.add(current);
        break;
      }
      seenAt.set(current, path.length);
      path.push(current);
      current = parentByUuid.get(current);
    }
  }
  return cycleNodes;
}

function validateTableArchitecture_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  setDispatcherRuntimeSpreadsheet_(ss);
  clearTableArchitectureCache_(ss);
  const architecture = getBaseArchitecture_(ss);
  const issues = [];
  const warnings = [];
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
      if (formulasBelow.some(row => Boolean(row[0]))) issues.push(`Найдены построчные формулы ниже якоря: ${key}`);
    }
  }

  if (lastDataRow >= 2) {
    const rowCount = lastDataRow - 1;
    const base = architecture.baseSheet;
    const uuids = base.getRange(2, baseCol_(ss, 'uuid'), rowCount, 1).getDisplayValues().flat();
    const parents = base.getRange(2, baseCol_(ss, 'parent_uuid'), rowCount, 1).getDisplayValues().flat();
    const roots = base.getRange(2, baseCol_(ss, 'root_uuid'), rowCount, 1).getDisplayValues().flat();
    const rootByUuid = new Map();
    const uuidCounts = new Map();
    uuids.forEach((uuidRaw, i) => {
      const uuid = String(uuidRaw || '').trim();
      if (!uuid) return;
      rootByUuid.set(uuid, roots[i]);
      uuidCounts.set(uuid, (uuidCounts.get(uuid) || 0) + 1);
    });

    let missingRoots = 0;
    let invalidTopRoots = 0;
    let selfParents = 0;
    let rootMismatches = 0;
    let orphanParents = 0;
    for (let i = 0; i < uuids.length; i++) {
      const uuid = String(uuids[i] || '').trim();
      if (!uuid) continue;
      const parent = String(parents[i] || '').trim();
      const root = String(roots[i] || '').trim();
      if (!root) missingRoots++;
      if (!parent && root && root !== uuid) invalidTopRoots++;
      if (parent && parent === uuid) selfParents++;
      if (parent) {
        if (!rootByUuid.has(parent)) orphanParents++;
        else {
          const expectedRoot = String(rootByUuid.get(parent) || parent).trim();
          if (root && expectedRoot && root !== expectedRoot) rootMismatches++;
        }
      }
    }
    const duplicateUuids = [...uuidCounts.values()].filter(count => count > 1).length;
    const cycleNodes = findHierarchyCycles_(uuids, parents);
    if (missingRoots) issues.push(`Строки без root_uuid: ${missingRoots}.`);
    if (invalidTopRoots) issues.push(`Корневые строки с root_uuid != собственному UUID: ${invalidTopRoots}.`);
    if (selfParents) issues.push(`Строки с parent_uuid = собственному UUID: ${selfParents}.`);
    if (rootMismatches) issues.push(`Несогласованный root_uuid относительно существующего родителя: ${rootMismatches}.`);
    if (duplicateUuids) issues.push(`Дублирующиеся UUID: ${duplicateUuids}.`);
    if (cycleNodes.size) issues.push(`Циклы parent_uuid: ${cycleNodes.size} узлов.`);
    if (orphanParents) warnings.push(`Исторические parent_uuid, отсутствующие в текущей База: ${orphanParents}.`);
  }

  const settingsSheet = ss.getSheetByName('Настройки');
  if (!settingsSheet) issues.push('Нет листа «Настройки».');
  else {
    const count = Math.max(settingsSheet.getLastRow() - 1, 1);
    const rows = settingsSheet.getRange(2, 1, count, 2).getValues();
    for (const [sheetNameRaw, fieldKeyRaw] of rows) {
      const sheetName = String(sheetNameRaw ?? '').trim();
      const fieldKey = String(fieldKeyRaw ?? '').trim();
      if (!sheetName && !fieldKey) continue;
      if (!sheetName || !fieldKey) issues.push(`Неполная настройка этапа: «${sheetName}» / «${fieldKey}».`);
      else if (!architecture.keyToColumn.has(fieldKey)) issues.push(`Этап «${sheetName}»: неизвестный field_key «${fieldKey}».`);
    }
  }

  const result = {
    ok: issues.length === 0,
    spreadsheetId: ss.getId(),
    mappedFields: architecture.keyToColumn.size,
    computedFields: architecture.computedKeys.size,
    lastDataRow,
    issues,
    warnings
  };
  Logger.log(JSON.stringify(result, null, 2));
  return result;
}
