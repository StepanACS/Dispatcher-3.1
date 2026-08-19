/** Создаёт устанавливаемый onEdit-триггер, если такого триггера ещё нет. */
function installEditTrigger() {
  const existing = ScriptApp.getProjectTriggers().some(trigger =>
    trigger.getHandlerFunction() === 'onEditHandler' &&
    trigger.getEventType() === ScriptApp.EventType.ON_EDIT
  );
  if (existing) {
    Logger.log('Триггер onEditHandler уже установлен.');
    return;
  }
  ScriptApp.newTrigger('onEditHandler')
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onEdit()
    .create();
}

/**
 * Базовый объект новой строки «База». Формулы сюда принципиально не входят:
 * вычисляемые столбцы обслуживаются единственными spill-якорями строки 2.
 */
function buildBaseRowTemplate() {
  const row = { uuid: Utilities.getUuid() };
  STAGE_FIELD_KEYS.forEach(key => { row[key] = false; });
  return row;
}

function columnLetterToNumber(letter) {
  const normalized = String(letter || '').trim().toUpperCase();
  if (!/^[A-Z]+$/.test(normalized)) return 0;
  let column = 0;
  for (let i = 0; i < normalized.length; i++) column = column * 26 + normalized.charCodeAt(i) - 64;
  return column;
}

function checkboxOk(range) {
  range.setValue(new Date());
}

function extractSheetId(link) {
  const match = String(link || '').match(/\/d\/([a-zA-Z0-9_-]+)/);
  return match ? match[1] : null;
}

/**
 * Ищет UUID. На «База» физический UUID-столбец определяется по schema field_key;
 * на пользовательских листах интерфейс стабилен и UUID остаётся в D.
 */
function findRowById(sheet, itemId) {
  if (!sheet || itemId === null || typeof itemId === 'undefined') return null;
  const target = String(itemId).trim();
  if (!target) return null;
  if (sheet.getName() === BASE_SHEET_NAME) return findBaseRowByUuid_(sheet, target);

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;
  const match = sheet
    .getRange(2, ROLE_COL_UUID, lastRow - 1, 1)
    .createTextFinder(target)
    .matchEntireCell(true)
    .findNext();
  return match ? match.getRow() : null;
}

function capitalizeFirstLetter(str) {
  const value = String(str || '').trim();
  return value ? value.charAt(0).toUpperCase() + value.slice(1) : '';
}

const ROLE_EDIT_TO_BASE_FIELD = Object.freeze({
  [ROLE_COL_ITEMNAME]: 'item_name',
  [ROLE_COL_DESTINATION]: 'destination',
  [ROLE_COL_NOTE]: 'note',
  [ROLE_COL_QUANTITY]: 'quantity'
});

/** Переносит ручное изменение с листа роли в соответствующее семантическое поле «База». */
function handleNote(ss, sheet, range, row, col, val) {
  const fieldKey = ROLE_EDIT_TO_BASE_FIELD[col];
  if (!fieldKey) return;

  const value = String(val ?? '').trim();
  range.setValue('');
  const itemId = String(sheet.getRange(row, ROLE_COL_UUID).getValue() || '').trim();
  if (!itemId) {
    ss.toast(`Идентификатор отсутствует, строка ${row}`, 'Ошибка', 5);
    return;
  }

  const baseSheet = ss.getSheetByName(BASE_SHEET_NAME);
  const baseRow = findBaseRowByUuid_(baseSheet, itemId);
  if (!baseRow) {
    ss.toast(`Изделие ${itemId} не найдено в Базе`, 'Ошибка', 5);
    return;
  }

  writeBaseFields_(baseSheet, baseRow, { [fieldKey]: value });
  ss.toast('Изменение перенесено в базу', 'Информация', 5);
}

/** Возвращает изделие на предыдущий этап по field_key этапа. */
function handleReturn(ss, sheet, range, row, clearFieldKey) {
  range.setValue(false);
  const itemId = String(sheet.getRange(row, ROLE_COL_UUID).getValue() || '').trim();
  if (!itemId) {
    ss.toast(`Идентификатор отсутствует, строка ${row}`, 'Ошибка', 5);
    return;
  }

  const baseSheet = ss.getSheetByName(BASE_SHEET_NAME);
  const baseRow = findBaseRowByUuid_(baseSheet, itemId);
  if (!baseRow) {
    ss.toast(`Изделие ${itemId} не найдено в Базе`, 'Ошибка', 5);
    return;
  }

  writeBaseFields_(baseSheet, baseRow, { [clearFieldKey]: false });
  ss.toast('Отметка о готовности снята', 'Информация', 5);
}

const SHEET_CHECKBOX_SETTINGS = Object.freeze({
  'Проектирование': { firstColumn: 11, count: 1 },
  'Техконтроль': { firstColumn: 11, count: 2 },
  'Развертывание спецификации': { firstColumn: 11, count: 2 },
  'Технолог': { firstColumn: 11, count: 2 },
  'Проверка технологии': { firstColumn: 11, count: 2 },
  'Диспетчер- комплектовщик': { firstColumn: 11, count: 2 },
  'Комплектация': { firstColumn: 11, count: 2 },
  'Склад': { firstColumn: 11, count: 2 },
  'Склад производства': { firstColumn: 11, count: 2 },
  'Производство': { firstColumn: 11, count: 2 },
  'ОТК': { firstColumn: 11, count: 2 },
  'Приемка': { firstColumn: 11, count: 2 },
  'Готово': { firstColumn: 11, count: 2 }
});

/** Устанавливает checkbox-validation только в строке текущего изделия. */
function fillCheckboxes(ss, itemId) {
  if (!itemId) return;
  const baseSheet = ss.getSheetByName(BASE_SHEET_NAME);
  const baseRow = findBaseRowByUuid_(baseSheet, itemId);
  if (!baseRow) return;

  const status = String(readBaseFields_(baseSheet, baseRow, ['status']).status || '').trim();
  const targetSheet = ss.getSheetByName(status);
  const settings = getSettingsForSheet(status);
  if (!targetSheet || !settings) return;

  SpreadsheetApp.flush();
  const targetRow = findRowById(targetSheet, itemId);
  if (!targetRow) {
    Logger.log(`fillCheckboxes: UUID ${itemId} пока не найден на листе «${status}».`);
    return;
  }

  const checkboxRange = targetSheet.getRange(targetRow, settings.firstColumn, 1, settings.count);
  const rule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  checkboxRange.setDataValidation(rule);
  checkboxRange.setValue(false);
}

function getSettingsForSheet(targetSheetName) {
  return SHEET_CHECKBOX_SETTINGS[targetSheetName] || null;
}

function getLastRow(sheet, col) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 1) return 0;
  const values = sheet.getRange(1, col, lastRow, 1).getValues();
  for (let i = values.length - 1; i >= 0; i--) {
    if (String(values[i][0] ?? '').trim() !== '') return i + 1;
  }
  return 0;
}

/** Формирует объект новой строки развёрнутой спецификации. */
function createSpecRow(date, name, fullQty, parentName, futureIndex, parentUUIDForDestination, opCheckboxStatus, totalColumnCount, rowUUID, note = '') {
  const row = buildBaseRowTemplate();
  const now = new Date();
  row.date = date == null ? now : date;
  row.item_name = name;
  if (rowUUID !== null && typeof rowUUID !== 'undefined') row.uuid = rowUUID;

  if (parentName != null) {
    const cleanedParentName = String(parentName).replace(/<[^>]*>/g, '');
    const cleanedParentUUID = String(parentUUIDForDestination).replace(/<[^>]*>/g, '');
    row.destination = `${cleanedParentName} = ${fullQty} (${cleanedParentUUID})`;
  } else {
    row.destination = '';
  }

  row.quantity = fullQty;
  row.note = note;
  row.stage_design = now;
  row.stage_technical_control = now;
  row.stage_deployment = now;
  if (opCheckboxStatus === true) {
    row.stage_technologist = now;
    row.stage_technology_check = now;
  }
  return row;
}

function normalizeDestination(destinationString) {
  const guidRegex = /\s\([0-9a-fA-F]{8}(?:-[0-9a-fA-F]{4}){3}-[0-9a-fA-F]{12}\)/g;
  return String(destinationString || '').replace(guidRegex, '');
}

/**
 * Общая утилита для ручных/служебных листов. Не использовать для «База»:
 * формульные столбцы «База» теперь обслуживаются spill-якорями, а не копированием строк.
 */
function ensureFreeRows(sheet, templateRow, minFreeRows = 100) {
  if (!sheet) throw new Error('ensureFreeRows: sheet is required');
  if (!Number.isInteger(templateRow) || templateRow < 1) throw new Error(`ensureFreeRows: invalid templateRow ${templateRow}`);
  if (!Number.isInteger(minFreeRows) || minFreeRows < 0) throw new Error(`ensureFreeRows: invalid minFreeRows ${minFreeRows}`);

  let totalRows = sheet.getMaxRows();
  if (templateRow > totalRows) {
    sheet.insertRowsAfter(totalRows, templateRow - totalRows);
    totalRows = sheet.getMaxRows();
  }
  const freeRows = totalRows - templateRow + 1;
  if (freeRows >= minFreeRows) return { added: 0, freeRows, totalRows };

  const rowsToAdd = minFreeRows - freeRows;
  sheet.insertRowsAfter(totalRows, rowsToAdd);
  const lastCol = sheet.getLastColumn();
  if (lastCol > 0) {
    const sourceRange = sheet.getRange(templateRow, 1, 1, lastCol);
    const targetRange = sheet.getRange(totalRows + 1, 1, rowsToAdd, lastCol);
    sourceRange.copyTo(targetRange, { contentsOnly: false });
  }
  return { added: rowsToAdd, freeRows: freeRows + rowsToAdd, totalRows: totalRows + rowsToAdd };
}
