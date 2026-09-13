function handlePickingOrder(ss, sheet, range, row) {
  const baseSheet = ss.getSheetByName(BASE_SHEET_NAME);
  const roleData = sheet.getRange(row, ROLE_COL_UUID, 1, ROLE_COL_ITEMNAME - ROLE_COL_UUID + 1).getValues()[0];
  const itemId = String(roleData[0] || '').trim();
  const itemName = String(roleData[ROLE_COL_ITEMNAME - ROLE_COL_UUID] || '').trim();
  const baseRow = findBaseRowByUuid_(baseSheet, itemId);
  if (!baseRow) {
    Logger.log(`handlePickingOrder: UUID ${itemId} не найден в Базе.`);
    return null;
  }

  const data = readBaseFields_(baseSheet, baseRow, [
    'order_link', 'order_quantity', 'deficit', 'production_mode',
    'salesdrive_order_id', 'stage_picking'
  ]);
  if (data.stage_picking) {
    if (sheet.getSheetId() !== baseSheet.getSheetId()) range.setValue(false);
    return itemId;
  }

  const orderQty = Number(data.order_quantity);
  if (!Number.isFinite(orderQty) || orderQty <= 0) {
    if (sheet.getSheetId() !== baseSheet.getSheetId()) range.setValue(false);
    ss.toast('Для комплектации не указано корректное «Количество в заказ».', 'Ошибка', 6);
    return null;
  }

  if (data.deficit > 0 && data.production_mode === 'В производство') {
    let orderId = String(data.salesdrive_order_id || '').trim();
    if (!orderId) {
      orderId = setSdOrder(ss, itemId, itemName, String(data.order_link || '').trim(), orderQty);
      writeBaseFields_(baseSheet, baseRow, { salesdrive_order_id: orderId });
    }
  }
  if (sheet.getSheetId() !== baseSheet.getSheetId()) range.setValue(false);
  checkboxOk(baseSheet.getRange(baseRow, baseCol_(ss, 'stage_picking')));
  return itemId;
}

function setSdOrder(ss, itemId, itemName, link, itemQty) {
  const comment = `${itemName} = ${itemQty}`;
  const spec = getSpec(itemName, link);
  const sdIdMap = loadSdIdMap_(ss);
  const products = [];
  for (const row of spec) {
    const name = String(row[0] || '').trim();
    const qty = parseFloat(row[2]);
    if (!name || isNaN(qty)) continue;
    products.push({
      id: getSdId(ss, name, sdIdMap), name, costPerItem: '',
      amount: qty * itemQty, discount: '', sku: name
    });
  }
  return sendPostRequest(ss, itemId, products, comment);
}
