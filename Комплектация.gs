function handlePickingOrder(ss, sheet, range, row) {
  const baseSheet = ss.getSheetByName('База');
  const roleData = sheet.getRange(row, COL_UUID, 1, COL_ITEMNAME - COL_UUID + 1).getValues()[0];
  const itemId = String(roleData[0] || '').trim();
  const itemName = String(roleData[COL_ITEMNAME - COL_UUID] || '').trim();

  const baseRow = findRowById(baseSheet, itemId);
  if (!baseRow) {
    Logger.log(`handlePickingOrder: UUID ${itemId} не найден в Базе.`);
    return null;
  }

  // Все необходимые значения из Базы читаются одним запросом.
  const baseData = baseSheet
    .getRange(baseRow, COL_LINK, 1, COL_MODE - COL_LINK + 1)
    .getValues()[0];

  const link = String(baseData[0] || '').trim();
  const qty = baseData[COL_QUANTITY - COL_LINK];
  const deficit = baseData[COL_DEFICIT - COL_LINK];
  const supplyWay = baseData[COL_MODE - COL_LINK];

  if (deficit > 0 && supplyWay === 'В производство') {
    const orderId = setSdOrder(ss, itemName, link, qty);
    baseSheet.getRange(baseRow, COL_SD_ORDER_ID).setValue(orderId);
  }

  if (sheet.getSheetId() !== baseSheet.getSheetId()) {
    range.setValue(false);
  }

  checkboxOk(baseSheet.getRange(baseRow, COL_PICKING));
  return itemId;
}

function setSdOrder(ss, itemName, link, itemQty) {
  const comment = `${itemName} = ${itemQty}`;
  const spec = getSpec(itemName, link);
  const sdIdMap = loadSdIdMap_(ss); // один раз на весь заказ
  const products = [];

  for (const row of spec) {
    const name = String(row[0] || '').trim();
    const qty = parseFloat(row[2]);
    if (!name || isNaN(qty)) continue;

    products.push({
      id: getSdId(ss, name, sdIdMap),
      name,
      costPerItem: '',
      amount: qty * itemQty,
      discount: '',
      sku: name
    });
  }

  return sendPostRequest(products, comment);
}
