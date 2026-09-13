/** Обрабатывает завершение этапа «Склад» и создаёт карточку Trello. */
function handleStock(ss, sheet, range, row) {
  const baseSheet = ss.getSheetByName(BASE_SHEET_NAME);
  const roleData = sheet.getRange(row, ROLE_COL_UUID, 1, ROLE_COL_ITEMNAME - ROLE_COL_UUID + 1).getValues()[0];
  const itemId = String(roleData[0] || '').trim();
  const itemName = String(roleData[ROLE_COL_ITEMNAME - ROLE_COL_UUID] || '').trim();
  const baseRow = findBaseRowByUuid_(baseSheet, itemId);
  if (!baseRow) {
    Logger.log(`handleStock: UUID ${itemId} не найден в Базе.`);
    return null;
  }

  const data = readBaseFields_(baseSheet, baseRow, [
    'destination', 'quantity', 'trello_id', 'stage_stock'
  ]);
  if (data.stage_stock) {
    if (sheet.getSheetId() !== baseSheet.getSheetId()) range.setValue(false);
    return itemId;
  }

  let cardId = String(data.trello_id || '').trim();
  if (!cardId) {
    const purpose = normalizeDestination(String(data.destination || '').trim());
    cardId = createTrelloCard(ss, itemId, itemName, data.quantity, purpose);
    writeBaseFields_(baseSheet, baseRow, { trello_id: cardId });
  }

  if (sheet.getSheetId() !== baseSheet.getSheetId()) range.setValue(false);
  checkboxOk(baseSheet.getRange(baseRow, baseCol_(ss, 'stage_stock')));
  return itemId;
}
