function handlePickingOrder(ss, sheet, range, row) {
  const context = resolveActionContext_(ss, sheet, row);
  if (!context) {
    Logger.log(`handlePickingOrder: не удалось определить строку База для ${sheet.getName()}!${row}.`);
    return null;
  }

  const data = readBaseFields_(context.baseSheet, context.baseRow, [
    'order_link', 'order_quantity', 'deficit', 'production_mode',
    'salesdrive_order_id', 'stage_picking'
  ]);
  if (data.stage_picking) {
    resetRoleCheckbox_(context, range);
    return context.itemId;
  }

  const orderQty = Number(data.order_quantity);
  if (!Number.isFinite(orderQty) || orderQty <= 0) {
    resetRoleCheckbox_(context, range);
    ss.toast('Для комплектации не указано корректное «Количество в заказ».', 'Ошибка', 6);
    return null;
  }

  if (data.deficit > 0 && data.production_mode === 'В производство') {
    let orderId = String(data.salesdrive_order_id || '').trim();
    if (!orderId) {
      orderId = setSdOrder(
        ss,
        context.itemId,
        context.itemName,
        String(data.order_link || '').trim(),
        orderQty
      );
      if (!orderId) throw new Error('SalesDrive не вернул ID созданного заказа.');
      writeBaseFields_(context.baseSheet, context.baseRow, { salesdrive_order_id: orderId });
    }
  }
  resetRoleCheckbox_(context, range);
  checkboxOk(context.baseSheet.getRange(context.baseRow, baseCol_(ss, 'stage_picking')));
  return context.itemId;
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
