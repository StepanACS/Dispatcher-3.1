/** Обрабатывает завершение этапа «Склад» и создаёт карточку Trello. */
function handleStock(ss, sheet, range, row) {
  const baseSheet = ss.getSheetByName('База');
  const roleData = sheet.getRange(row, COL_UUID, 1, COL_ITEMNAME - COL_UUID + 1).getValues()[0];
  const itemId = String(roleData[0] || '').trim();
  const itemName = String(roleData[COL_ITEMNAME - COL_UUID] || '').trim();

  const baseRow = findRowById(baseSheet, itemId);
  if (!baseRow) {
    Logger.log(`handleStock: UUID ${itemId} не найден в Базе.`);
    return null;
  }

  const baseData = baseSheet
    .getRange(baseRow, COL_DESTINATION, 1, COL_QUANTITY - COL_DESTINATION + 1)
    .getValues()[0];

  const purpose = normalizeDestination(String(baseData[0] || '').trim());
  const qty = baseData[COL_QUANTITY - COL_DESTINATION];
  const cardId = createTrelloCard(itemName, qty, purpose);

  baseSheet.getRange(baseRow, COL_TRELLO_ID).setValue(cardId);

  // Сначала сбрасываем ручной чекбокс на QUERY-листе. После отметки в Базе
  // строка может сменить статус и исчезнуть, сдвинув следующую строку вверх.
  if (sheet.getSheetId() !== baseSheet.getSheetId()) {
    range.setValue(false);
  }

  checkboxOk(baseSheet.getRange(baseRow, COL_STOCK));
  return itemId;
}
