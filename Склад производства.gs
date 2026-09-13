function handleProductionStock(ss, sheet, range, row) {
  let context = resolveActionContext_(ss, sheet, row);
  if (!context) {
    Logger.log(`handleProductionStock: не удалось определить строку База для ${sheet.getName()}!${row}.`);
    return null;
  }

  const data = readBaseFields_(context.baseSheet, context.baseRow, [
    'salesdrive_order_id', 'stage_production_stock'
  ]);
  if (data.stage_production_stock) {
    resetRoleCheckbox_(context, range);
    return context.itemId;
  }

  if (data.salesdrive_order_id) changeSdStatus(data.salesdrive_order_id, 23);
  resetRoleCheckbox_(context, range);

  removeSpecification(context.itemId, ss);
  const baseRowAfter = findBaseRowByUuid_(context.baseSheet, context.itemId);
  if (!baseRowAfter) {
    Logger.log(`handleProductionStock: корневая строка ${context.itemId} не найдена после сворачивания.`);
    return null;
  }

  checkboxOk(context.baseSheet.getRange(baseRowAfter, baseCol_(ss, 'stage_production_stock')));
  return context.itemId;
}
