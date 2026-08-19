/**
 * Создаёт устанавливаемый onEdit-триггер, если такого триггера ещё нет.
 */
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
 * Формирует шаблон строки листа «База».
 * Формулы оставлены без изменения: оптимизация формульной модели таблицы будет
 * выполняться отдельным этапом.
 */
function buildBaseRowTemplate(futureIndex, sheetColumnCount) {
  const newRow = Array(sheetColumnCount).fill('');

  newRow[COL_DRAWING - 1] = `=IF(H${futureIndex}="";"";IF(E${futureIndex}="";"";IFERROR(IF(FILTER('Каталог ссылок'!B:B;'Каталог ссылок'!A:A=E${futureIndex})="";"Добавить ссылку в каталог";FILTER('Каталог ссылок'!B:B;'Каталог ссылок'!A:A=E${futureIndex}));"Добавить изделие в каталог ссылок")))`;
  newRow[COL_LINK - 1] = `=IF(H${futureIndex}="";"";IF(E${futureIndex}="";"";IFERROR(IF(FILTER('Каталог ссылок'!C:C;'Каталог ссылок'!A:A=E${futureIndex})="";"Добавить ссылку в каталог";FILTER('Каталог ссылок'!C:C;'Каталог ссылок'!A:A=E${futureIndex}));"Добавить изделие в каталог ссылок")))`;
  newRow[COL_UUID - 1] = Utilities.getUuid();
  newRow[COL_IN_STOCK - 1] = `=IFERROR(IF(E${futureIndex}<>"" ; FILTER('Остатки'!B:B; 'Остатки'!A:A=E${futureIndex}); ""); 0)`;
  newRow[COL_DEFICIT - 1] = `=IF((H${futureIndex}-I${futureIndex})<0; 0; H${futureIndex}-I${futureIndex})`;
  newRow[COL_READY_FOR_PRODUCTION - 1] = `=IFERROR(IF(L${futureIndex} = "Да";IF(OR(Z${futureIndex} = "Производство"; Z${futureIndex} = "Склад производства"); "В производстве";IF(OR(Z${futureIndex} = "ОТК"; Z${futureIndex} = "Приемка"; Z${futureIndex} = "Готово"); "Готово";IF(AA${futureIndex} = "В производство";IF(SUM(FILTER(J$2:J; ISNUMBER(SEARCH(E${futureIndex}; F$2:F)))) > 0; "Нет"; "Да");"Покупной")));"Нет");"Комплектующие не найдены")`;
  newRow[COL_IS_NEED - 1] = `=IF(J${futureIndex}>0; "Да"; "Нет")`;
  newRow[COL_STATUS - 1] = `=IFNA(INDEX('Настройки'!A$2:A$13; MATCH(FALSE; M${futureIndex}:W${futureIndex}; 0)); "Готово")`;
  newRow[COL_MODE - 1] = `=IF(E${futureIndex}="";"";IF(COUNTIF('Цена материалов'!A:A; E${futureIndex})=0;"В производство";"В снабжение"))`;
  newRow[COL_SDID - 1] = `=IFERROR(FILTER('SD ID'!A:A; 'SD ID'!B:B = E${futureIndex});"")`;
  newRow[COL_TRELLO_ID - 1] = '';
  newRow[COL_SD_ORDER_ID - 1] = '';
  newRow[COL_FACT_QTY - 1] = '';
  newRow[COL_TECHNOLOGIEST_NAME - 1] = '';
  newRow[COL_DESIGNER_NAME - 1] = '';
  newRow[COL_SUPPLY_ORDERED - 1] = `=IFERROR(SUM(FILTER('Снабжение'!C:C; 'Снабжение'!A:A = E${futureIndex})); "")`;
  newRow[COL_TIED_CAPITAL - 1] = `=IFERROR(IF(F${futureIndex}=""; FILTER('ГТ Себестоимость'!C:C; 'ГТ Себестоимость'!A:A = E${futureIndex}); 0); 0) * H${futureIndex}`;
  newRow[COL_QTY_SD - 1] = `=IFERROR(IF(E${futureIndex}<>""; FILTER('Остатки'!F:F; 'Остатки'!E:E = E${futureIndex}); ""); "")`;
  newRow[COL_ABC_CATEGORY - 1] = `=IFERROR(IF(F${futureIndex}=""; FILTER(ABC!B:B; ABC!A:A = E${futureIndex}); ""); "")`;
  newRow[COL_PRIORITY - 1] = '';
  newRow[COL_EQIPMENT - 1] = `=IFERROR(TEXTJOIN(", "; TRUE; FILTER('Диспетчер оснастки'!A:A; 'Диспетчер оснастки'!B:B = E${futureIndex})); "")`;
  newRow[COL_DEADLINE - 1] = '';

  for (let col = COL_DESIGN - 1; col <= COL_READY - 1; col++) {
    newRow[col] = false;
  }

  return newRow;
}

function columnLetterToNumber(letter) {
  const normalized = String(letter || '').trim().toUpperCase();
  if (!/^[A-Z]+$/.test(normalized)) return 0;

  let column = 0;
  for (let i = 0; i < normalized.length; i++) {
    column = column * 26 + normalized.charCodeAt(i) - 64;
  }
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
 * Ищет строку по UUID только в столбце D.
 * Раньше функция читала getDataRange(), то есть все ~40 столбцов «Базы» при
 * каждом поиске. TextFinder оставляет поиск на стороне Sheets и не переносит
 * весь лист в память Apps Script.
 */
function findRowById(sheet, itemId) {
  if (!sheet || itemId === null || typeof itemId === 'undefined') return null;

  const target = String(itemId).trim();
  if (!target) return null;

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;

  const match = sheet
    .getRange(2, COL_UUID, lastRow - 1, 1)
    .createTextFinder(target)
    .matchEntireCell(true)
    .findNext();

  return match ? match.getRow() : null;
}

function capitalizeFirstLetter(str) {
  const value = String(str || '').trim();
  return value ? value.charAt(0).toUpperCase() + value.slice(1) : '';
}

/**
 * Переносит ручное изменение с листа роли в соответствующий столбец «Базы».
 */
function handleNote(ss, sheet, range, row, col, val) {
  const value = String(val ?? '').trim();
  range.setValue('');

  const itemId = String(sheet.getRange(row, COL_UUID).getValue() || '').trim();
  if (!itemId) {
    ss.toast(`Идентификатор отсутствует, строка ${row}`, 'Ошибка', 5);
    return;
  }

  const baseSheet = ss.getSheetByName('База');
  const baseRow = findRowById(baseSheet, itemId);
  if (!baseRow) {
    ss.toast(`Изделие ${itemId} не найдено в Базе`, 'Ошибка', 5);
    return;
  }

  baseSheet.getRange(baseRow, col).setValue(value);
  ss.toast('Изменение перенесено в базу', 'Информация', 5);
}

/**
 * Возвращает изделие на предыдущий этап, снимая соответствующую отметку в «Базе».
 */
function handleReturn(ss, sheet, range, row, clearColumn) {
  range.setValue(false);

  const itemId = String(sheet.getRange(row, COL_UUID).getValue() || '').trim();
  if (!itemId) {
    ss.toast(`Идентификатор отсутствует, строка ${row}`, 'Ошибка', 5);
    return;
  }

  const baseSheet = ss.getSheetByName('База');
  const baseRow = findRowById(baseSheet, itemId);
  if (!baseRow) {
    ss.toast(`Изделие ${itemId} не найдено в Базе`, 'Ошибка', 5);
    return;
  }

  baseSheet.getRange(baseRow, clearColumn).setValue(false);
  ss.toast('Отметка о готовности снята', 'Информация', 5);
}

const SHEET_CHECKBOX_SETTINGS = Object.freeze({
  'Проектирование': { firstColumn: 11, count: 1 },
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

/**
 * Устанавливает checkbox-validation только в строке текущего изделия.
 * Раньше функция переписывала K:L для всего рабочего листа после каждого клика.
 */
function fillCheckboxes(ss, itemId) {
  if (!itemId) return;

  const baseSheet = ss.getSheetByName('База');
  const baseRow = findRowById(baseSheet, itemId);
  if (!baseRow) return;

  const status = String(baseSheet.getRange(baseRow, COL_STATUS).getValue() || '').trim();
  const targetSheet = ss.getSheetByName(status);
  const settings = getSettingsForSheet(status);
  if (!targetSheet || !settings) return;

  // Формула статуса и QUERY-лист должны успеть пересчитаться до поиска UUID.
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

/**
 * Определяет последнюю непустую строку указанного столбца.
 * Оставлено как общая утилита для функций, которым нужен именно этот критерий.
 */
function getLastRow(sheet, col) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 1) return 0;

  const values = sheet.getRange(1, col, lastRow, 1).getValues();
  for (let i = values.length - 1; i >= 0; i--) {
    if (String(values[i][0] ?? '').trim() !== '') return i + 1;
  }
  return 0;
}

/**
 * Формирует строку для записи развёрнутой спецификации в «База».
 */
function createSpecRow(date, name, fullQty, parentName, futureIndex, parentUUIDForDestination, opCheckboxStatus, totalColumnCount, rowUUID, note = '') {
  const row = buildBaseRowTemplate(futureIndex, totalColumnCount);
  const now = new Date();

  row[COL_DATE - 1] = date == null ? now : date;
  row[COL_ITEMNAME - 1] = name;

  if (rowUUID !== null && typeof rowUUID !== 'undefined') {
    row[COL_UUID - 1] = rowUUID;
  }

  if (parentName != null) {
    const cleanedParentName = String(parentName).replace(/<[^>]*>/g, '');
    const cleanedParentUUID = String(parentUUIDForDestination).replace(/<[^>]*>/g, '');
    row[COL_DESTINATION - 1] = `${cleanedParentName} = ${fullQty} (${cleanedParentUUID})`;
  } else {
    row[COL_DESTINATION - 1] = '';
  }

  row[COL_QUANTITY - 1] = fullQty;
  row[COL_NOTE - 1] = note;
  row[COL_DESIGN - 1] = now;
  row[COL_TECHNICAL_CONTROL - 1] = now;
  row[COL_DEPLOYMENT - 1] = now;

  if (opCheckboxStatus === true) {
    row[COL_TECHNOLOGIEST - 1] = now;
    row[COL_TECHNOLOGY_CHECK - 1] = now;
  }

  return row;
}

/** Удаляет UUID в скобках из текста назначения. */
function normalizeDestination(destinationString) {
  const guidRegex = /\s\([0-9a-fA-F]{8}(?:-[0-9a-fA-F]{4}){3}-[0-9a-fA-F]{12}\)/g;
  return String(destinationString || '').replace(guidRegex, '');
}

/**
 * Гарантирует наличие minFreeRows строк, начиная с templateRow, и при необходимости
 * копирует шаблон строки вместе с формулами, форматами и validation.
 */
function ensureFreeRows(sheet, templateRow, minFreeRows = 100) {
  if (!sheet) throw new Error('ensureFreeRows: sheet is required');
  if (!Number.isInteger(templateRow) || templateRow < 1) {
    throw new Error(`ensureFreeRows: templateRow must be positive integer, got ${templateRow}`);
  }
  if (!Number.isInteger(minFreeRows) || minFreeRows < 0) {
    throw new Error(`ensureFreeRows: minFreeRows must be non-negative integer, got ${minFreeRows}`);
  }

  let totalRows = sheet.getMaxRows();
  if (templateRow > totalRows) {
    sheet.insertRowsAfter(totalRows, templateRow - totalRows);
    totalRows = sheet.getMaxRows();
  }

  const freeRows = totalRows - templateRow + 1;
  if (freeRows >= minFreeRows) {
    return { added: 0, freeRows, totalRows };
  }

  const rowsToAdd = minFreeRows - freeRows;
  sheet.insertRowsAfter(totalRows, rowsToAdd);

  const lastCol = sheet.getLastColumn();
  if (lastCol > 0) {
    const sourceRange = sheet.getRange(templateRow, 1, 1, lastCol);
    const targetRange = sheet.getRange(totalRows + 1, 1, rowsToAdd, lastCol);
    sourceRange.copyTo(targetRange, { contentsOnly: false });
  }

  return {
    added: rowsToAdd,
    freeRows: freeRows + rowsToAdd,
    totalRows: totalRows + rowsToAdd
  };
}
