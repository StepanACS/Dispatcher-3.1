/**
 * Порядок старой части архива закреплён семантически и не зависит от порядка
 * столбцов «База». Первые 31 поле + цена сохраняют прежний контракт A:AF.
 */
const ARCHIVE_BASE_FIELD_ORDER = Object.freeze([
  'date', 'drawing_link', 'order_link', 'uuid', 'item_name', 'destination', 'note',
  'quantity', 'stock', 'deficit', 'need_status', 'need_flag',
  'stage_design', 'stage_technical_control', 'stage_deployment', 'stage_technologist',
  'stage_technology_check', 'stage_picker', 'stage_picking', 'stage_stock',
  'stage_production_stock', 'stage_production', 'stage_quality_check',
  'stage_acceptance', 'stage_ready', 'status', 'production_mode', 'salesdrive_id',
  'trello_id', 'salesdrive_order_id', 'actual_made'
]);

// Новая часть добавляется ПОСЛЕ прежней цены и не сдвигает существующие колонки.
const ARCHIVE_EXT_FIELD_ORDER = Object.freeze([
  'parent_uuid', 'root_uuid', 'order_quantity', 'minimum_batch',
  'supplier', 'supplier_request_date', 'manual_status'
]);
const ARCHIVE_EXT_HEADERS = Object.freeze([
  'Родитель ID', 'Первичное назначение ID', 'Количество в заказ', 'Минимальная партия',
  'Поставщик', 'Дата запроса', 'Статус сотрудника'
]);

function handleReady(ss, sheet, range, row) {
  const context = resolveActionContext_(ss, sheet, row);
  if (!context) {
    Logger.log(`handleReady: не удалось определить строку База для ${sheet.getName()}!${row}.`);
    return null;
  }

  const archiveSs = SpreadsheetApp.openById(ARCHIVE_SPREADSHEET_ID);
  const archiveSheet = archiveSs.getSheetByName('Архив');
  if (!archiveSheet) throw new Error('В архивном документе не найден лист «Архив».');
  resetRoleCheckbox_(context, range);

  const stageData = readBaseFields_(context.baseSheet, context.baseRow, ['stage_acceptance', 'stage_ready']);
  const updates = {};
  const now = new Date();
  if (!stageData.stage_acceptance) updates.stage_acceptance = now;
  if (!stageData.stage_ready) updates.stage_ready = now;
  if (Object.keys(updates).length) writeBaseFields_(context.baseSheet, context.baseRow, updates);

  // status и другие computed-поля должны успеть пересчитаться перед снимком архива.
  SpreadsheetApp.flush();

  const alreadyArchivedRow = findArchiveRowByUuid_(archiveSheet, context.itemId);
  if (!alreadyArchivedRow) {
    const allFields = [...ARCHIVE_BASE_FIELD_ORDER, ...ARCHIVE_EXT_FIELD_ORDER];
    const data = readBaseFields_(context.baseSheet, context.baseRow, allFields);
    const archiveValues = ARCHIVE_BASE_FIELD_ORDER.map(key => data[key]);

    // В тестовом Диспетчере сохраняем реальную архивную операцию, но делаем её
    // очевидной сотруднику, не меняя layout старых колонок.
    const noteIndex = ARCHIVE_BASE_FIELD_ORDER.indexOf('note');
    if (isTestDispatcher_(ss) && noteIndex >= 0) {
      archiveValues[noteIndex] = markExternalTestText_(ss, archiveValues[noteIndex]);
    }

    const price = fetchPrice_(context.itemName);
    const extraValues = ARCHIVE_EXT_FIELD_ORDER.map(key => data[key]);
    ensureArchiveExtension_(archiveSheet);
    const archiveRow = archiveSheet.getLastRow() + 1;
    archiveSheet.getRange(archiveRow, 1, 1, archiveValues.length + 1 + extraValues.length)
      .setValues([[...archiveValues, price, ...extraValues]]);
  }

  const trelloId = String(readBaseFields_(context.baseSheet, context.baseRow, ['trello_id']).trello_id || '').trim();
  const trelloClosed = closeTrelloCard(trelloId);
  if (trelloId && !trelloClosed) {
    ss.toast('Не удалось архивировать карточку Trello. Проверьте её вручную.', 'Внимание', 8);
  }

  // Архив подтверждён — теперь можно очистить первичную строку.
  clearBaseRawRow_(context.baseSheet, context.baseRow);
  return context.itemId;
}

function findArchiveRowByUuid_(archiveSheet, itemId) {
  const target = String(itemId || '').trim();
  if (!target) return null;
  const lastRow = archiveSheet.getLastRow();
  if (lastRow < 2) return null;
  const match = archiveSheet.getRange(2, 4, lastRow - 1, 1)
    .createTextFinder(target).matchEntireCell(true).findNext();
  return match ? match.getRow() : null;
}

function ensureArchiveExtension_(archiveSheet) {
  const legacyWidthWithPrice = ARCHIVE_BASE_FIELD_ORDER.length + 1;
  const requiredWidth = legacyWidthWithPrice + ARCHIVE_EXT_HEADERS.length;
  if (archiveSheet.getMaxColumns() < requiredWidth) {
    archiveSheet.insertColumnsAfter(archiveSheet.getMaxColumns(), requiredWidth - archiveSheet.getMaxColumns());
  }
  const range = archiveSheet.getRange(1, legacyWidthWithPrice + 1, 1, ARCHIVE_EXT_HEADERS.length);
  const existing = range.getDisplayValues()[0];
  if (existing.every(value => !String(value || '').trim())) range.setValues([ARCHIVE_EXT_HEADERS]);
}

function fetchPrice_(itemName) {
  const priceSs = SpreadsheetApp.openById(PRICE_DB_SPREADSHEET_ID);
  const dbSheet = priceSs.getSheetByName('БД артикулы');
  if (!dbSheet) throw new Error('Лист «БД артикулы» не найден.');
  const lastRow = dbSheet.getLastRow();
  if (lastRow < 1) return 0;
  const match = dbSheet.getRange(1, 1, lastRow, 1)
    .createTextFinder(String(itemName || '').trim())
    .matchCase(true).matchEntireCell(true).findNext();
  return match ? dbSheet.getRange(match.getRow(), 3).getValue() : 0;
}
