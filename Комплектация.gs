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

  const data = readBaseFields_(baseSheet, baseRow, ['order_link', 'quantity', 'deficit', 'production_mode']);
  if (data.deficit > 0 && data.production_mode === 'В производство') {
    const orderId = setSdOrder(ss, itemName, String(data.order_link || '').trim(), data.quantity);
    writeBaseFields_(baseSheet, baseRow, { salesdrive_order_id: orderId });
  }
  if (sheet.getSheetId() !== baseSheet.getSheetId()) range.setValue(false);
  checkboxOk(baseSheet.getRange(baseRow, baseCol_(ss, 'stage_picking')));
  return itemId;
}

function setSdOrder(ss, itemName, link, itemQty) {
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
  return sendPostRequest(products, comment);
}
