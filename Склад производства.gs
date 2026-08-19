function handleProductionStock(ss, sheet, range, row) {
  const baseSheet = ss.getSheetByName(BASE_SHEET_NAME);
  const itemId = String(sheet.getRange(row, ROLE_COL_UUID).getValue() || '').trim();
  let baseRow = findBaseRowByUuid_(baseSheet, itemId);
  if (!baseRow) {
    Logger.log(`handleProductionStock: UUID ${itemId} не найден в Базе.`);
    return null;
  }

  const sdOrderId = readBaseFields_(baseSheet, baseRow, ['salesdrive_order_id']).salesdrive_order_id;
  if (sdOrderId) changeSdStatus(sdOrderId, 23);
  if (sheet.getSheetId() !== baseSheet.getSheetId()) range.setValue(false);

  handleUndeployment(ss, sheet, row);
  baseRow = findBaseRowByUuid_(baseSheet, itemId);
  if (!baseRow) {
    Logger.log(`handleProductionStock: корневая строка ${itemId} не найдена после сворачивания.`);
    return null;
  }

  checkboxOk(baseSheet.getRange(baseRow, baseCol_(ss, 'stage_production_stock')));
  return itemId;
}
