function handleProductionStock(ss, sheet, range, row) {
  const baseSheet = ss.getSheetByName('База');
  const itemId = String(sheet.getRange(row, COL_UUID).getValue() || '').trim();

  let baseRow = findRowById(baseSheet, itemId);
  if (!baseRow) {
    Logger.log(`handleProductionStock: UUID ${itemId} не найден в Базе.`);
    return null;
  }

  const sdOrderId = baseSheet.getRange(baseRow, COL_SD_ORDER_ID).getValue();
  if (sdOrderId) changeSdStatus(sdOrderId, 23); // 23 — «Передано в производство»

  // Сворачивание спецификации может удалять строки выше текущей, поэтому старый
  // baseRow после этой операции использовать нельзя.
  handleUndeployment(ss, sheet, row);
  baseRow = findRowById(baseSheet, itemId);
  if (!baseRow) {
    Logger.log(`handleProductionStock: корневая строка ${itemId} не найдена после сворачивания.`);
    return null;
  }

  checkboxOk(baseSheet.getRange(baseRow, COL_PRODUCTION_STOCK));

  if (sheet.getSheetId() !== baseSheet.getSheetId()) {
    range.setValue(false);
  }

  return itemId;
}
