/**
 * Обрабатывает готовность изделия: ставит финальные отметки, закрывает Trello,
 * переносит строку в архив и удаляет её из «Базы».
 */
function handleReady(ss, sheet, range, row) {
  const baseSheet = ss.getSheetByName('База');
  const archiveSs = SpreadsheetApp.openById(ARCHIVE_SPREADSHEET_ID);
  const archiveSheet = archiveSs.getSheetByName('Архив');
  if (!archiveSheet) throw new Error('В архивном документе не найден лист «Архив».');

  const roleData = sheet.getRange(row, COL_UUID, 1, COL_ITEMNAME - COL_UUID + 1).getValues()[0];
  const itemId = String(roleData[0] || '').trim();
  const itemName = String(roleData[COL_ITEMNAME - COL_UUID] || '').trim();

  const baseRow = findRowById(baseSheet, itemId);
  if (!baseRow) {
    Logger.log(`handleReady: UUID ${itemId} не найден в Базе.`);
    return null;
  }

  const now = new Date();
  baseSheet
    .getRange(baseRow, COL_ACCEPTANCE, 1, COL_READY - COL_ACCEPTANCE + 1)
    .setValues([[now, now]]);

  const rowData = baseSheet.getRange(baseRow, 1, 1, COL_FACT_QTY).getValues()[0];
  closeTrelloCard(rowData[COL_TRELLO_ID - 1]);

  const price = fetchPrice_(itemName);
  const archiveRow = archiveSheet.getLastRow() + 1;
  archiveSheet
    .getRange(archiveRow, 1, 1, COL_FACT_QTY + 1)
    .setValues([[...rowData, price]]);

  baseSheet.deleteRow(baseRow);

  if (sheet.getSheetId() !== baseSheet.getSheetId()) {
    range.setValue(false);
  }

  return itemId;
}

/** Ищет цену server-side через TextFinder вместо загрузки всей БД A:C. */
function fetchPrice_(itemName) {
  const priceSs = SpreadsheetApp.openById(PRICE_DB_SPREADSHEET_ID);
  const dbSheet = priceSs.getSheetByName('БД артикулы');
  if (!dbSheet) throw new Error('Лист «БД артикулы» не найден.');

  const lastRow = dbSheet.getLastRow();
  if (lastRow < 1) return 0;

  const match = dbSheet
    .getRange(1, 1, lastRow, 1)
    .createTextFinder(String(itemName || '').trim())
    .matchCase(true)
    .matchEntireCell(true)
    .findNext();

  return match ? dbSheet.getRange(match.getRow(), 3).getValue() : 0;
}

/** Совместимая обёртка для старых ручных вызовов. */
function fetchAndAppendPrice(itemName, archiveSheet, archiveLastRow, colFactQty) {
  archiveSheet.getRange(archiveLastRow + 1, colFactQty + 1).setValue(fetchPrice_(itemName));
}
