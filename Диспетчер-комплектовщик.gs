/** Обрабатывает завершение этапа «Диспетчер- комплектовщик». */
function handleSupplyOrder(ss, sheet, range, row) {
  const baseSheet = ss.getSheetByName(BASE_SHEET_NAME);
  const roleData = sheet.getRange(row, ROLE_COL_UUID, 1, ROLE_COL_ITEMNAME - ROLE_COL_UUID + 1).getValues()[0];
  const itemId = String(roleData[0] || '').trim();
  const itemName = String(roleData[ROLE_COL_ITEMNAME - ROLE_COL_UUID] || '').trim();
  const baseRow = findBaseRowByUuid_(baseSheet, itemId);
  if (!baseRow) {
    Logger.log(`handleSupplyOrder: UUID ${itemId} не найден в Базе.`);
    return null;
  }

  const data = readBaseFields_(baseSheet, baseRow, ['destination', 'quantity', 'deficit', 'production_mode']);
  if (data.deficit > 0 && data.production_mode === 'В снабжение') {
    setSupplyOrder(itemName, data.quantity, String(data.destination || '').trim());
  }
  if (sheet.getSheetId() !== baseSheet.getSheetId()) range.setValue(false);
  checkboxOk(baseSheet.getRange(baseRow, baseCol_(ss, 'stage_picker')));
  return itemId;
}

function setSupplyOrder(name, qty, purpose) {
  const supplySpreadsheet = SpreadsheetApp.openById(SUPPLY_SPREADSHEET_ID);
  const supplySheet = supplySpreadsheet.getSheetByName('Сводный');
  if (!supplySheet) throw new Error('В документе снабжения не найден лист «Сводный».');
  const newOrderRow = getLastRowNumber(supplySheet, 2) + 1;
  supplySheet.getRange(newOrderRow, 1, 1, 8).setValues([[
    new Date(), 'Script - Диспетчер 3.1', qty, name, '', '', '', purpose
  ]]);
}

function getLastRowNumber(sheet, column) {
  const maxRows = sheet.getMaxRows();
  if (maxRows < 1) return 0;
  return sheet.getRange(maxRows, column).getNextDataCell(SpreadsheetApp.Direction.UP).getRow();
}

function getSpec(itemName, link) {
  const sheetId = extractSheetId(link);
  if (!sheetId) throw new Error(`Не удалось извлечь ID таблицы из ссылки для «${itemName}».`);
  const targetSS = SpreadsheetApp.openById(sheetId);
  const specSheet = targetSS.getSheetByName(itemName);
  if (!specSheet) throw new Error(`Лист «${itemName}» не найден в таблице спецификации.`);
  const firstDataRow = 5;
  const lastRow = specSheet.getLastRow();
  if (lastRow < firstDataRow) return [];
  const raw = specSheet.getRange(firstDataRow, 2, lastRow - firstDataRow + 1, 3).getValues();
  const totalIndex = raw.findIndex(r => String(r[0] || '').trim().toLowerCase() === 'итого');
  return totalIndex === -1 ? [] : raw.slice(0, totalIndex);
}

function loadSdIdMap_(ss) {
  const result = new Map();
  const sdSheet = ss.getSheetByName('SD ID');
  if (!sdSheet) return result;
  const lastRow = sdSheet.getLastRow();
  if (lastRow < 1) return result;
  const data = sdSheet.getRange(1, 1, lastRow, 2).getValues();
  for (const row of data) {
    const name = String(row[1] || '').trim();
    if (name) result.set(name, String(row[0] || '').trim());
  }
  return result;
}

function getSdId(ss, name, sdIdMap) {
  const map = sdIdMap || loadSdIdMap_(ss);
  return map.get(String(name || '').trim()) || null;
}
