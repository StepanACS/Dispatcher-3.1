/** Обрабатывает завершение этапа «Склад» и создаёт карточку Trello. */
function handleStock(ss, sheet, range, row) {
  const context = resolveActionContext_(ss, sheet, row);
  if (!context) {
    Logger.log(`handleStock: не удалось определить строку База для ${sheet.getName()}!${row}.`);
    return null;
  }

  const data = readBaseFields_(context.baseSheet, context.baseRow, [
    'destination', 'quantity', 'trello_id', 'stage_stock'
  ]);
  if (data.stage_stock) {
    resetRoleCheckbox_(context, range);
    return context.itemId;
  }

  let cardId = String(data.trello_id || '').trim();
  if (!cardId) {
    const purpose = normalizeDestination(String(data.destination || '').trim());
    cardId = createTrelloCard(ss, context.itemId, context.itemName, data.quantity, purpose);
    if (!cardId) throw new Error('Trello не вернул ID созданной карточки.');
    writeBaseFields_(context.baseSheet, context.baseRow, { trello_id: cardId });
  }

  resetRoleCheckbox_(context, range);
  checkboxOk(context.baseSheet.getRange(context.baseRow, baseCol_(ss, 'stage_stock')));
  return context.itemId;
}
