/**
 * Порядок полей архива закреплён семантически и больше не зависит от порядка
 * столбцов «База». Это сохраняет прежний контракт архива A:AE.
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

function handleReady(ss, sheet, range, row) {
  const baseSheet = ss.getSheetByName(BASE_SHEET_NAME);
  const archiveSs = SpreadsheetApp.openById(ARCHIVE_SPREADSHEET_ID);
  const archiveSheet = archiveSs.getSheetByName('Архив');
  if (!archiveSheet) throw new Error('В архивном документе не найден лист «Архив».');

  const roleData = sheet.getRange(row, ROLE_COL_UUID, 1, ROLE_COL_ITEMNAME - ROLE_COL_UUID + 1).getValues()[0];
  const itemId = String(roleData[0] || '').trim();
  const itemName = String(roleData[ROLE_COL_ITEMNAME - ROLE_COL_UUID] || '').trim();
  const baseRow = findBaseRowByUuid_(baseSheet, itemId);
  if (!baseRow) {
    Logger.log(`handleReady: UUID ${itemId} не найден в Базе.`);
    return null;
  }

  if (sheet.getSheetId() !== baseSheet.getSheetId()) range.setValue(false);
  const now = new Date();
  writeBaseFields_(baseSheet, baseRow, {
    stage_acceptance: now,
    stage_ready: now
  });
  SpreadsheetApp.flush();

  const data = readBaseFields_(baseSheet, baseRow, ARCHIVE_BASE_FIELD_ORDER);
  const archiveValues = ARCHIVE_BASE_FIELD_ORDER.map(key => data[key]);
  const trelloId = data.trello_id;
  const trelloClosed = closeTrelloCard(trelloId);
  if (trelloId && !trelloClosed) {
    ss.toast('Не удалось архивировать карточку Trello. Проверьте её вручную.', 'Внимание', 8);
  }

  const price = fetchPrice_(itemName);
  const archiveRow = archiveSheet.getLastRow() + 1;
  archiveSheet.getRange(archiveRow, 1, 1, archiveValues.length + 1).setValues([[...archiveValues, price]]);

  // Не удаляем физическую строку: строка 2 содержит spill-якоря вычисляемых полей.
  clearBaseRawRow_(baseSheet, baseRow);
  return itemId;
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

function fetchAndAppendPrice(itemName, archiveSheet, archiveLastRow, colFactQty) {
  archiveSheet.getRange(archiveLastRow + 1, colFactQty + 1).setValue(fetchPrice_(itemName));
}
