function handlePickingOrder(ss, sheet, range, row) {
  var baseSheet = ss.getSheetByName("База");

  // Читаем всю строку, где сработало событие, одним вызовом
  var rowData = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
  // Если константы определены с 1-индексацией, корректируем индексы:
  var itemId = rowData[COL_UUID - 1].toString().trim();
  var itemName = rowData[COL_ITEMNAME - 1].toString().trim();

  // Находим строку в "База" по itemId
  var baseRow = findRowById(baseSheet, itemId);
  if (!baseRow) {
    Logger.log("Строка с itemId " + itemId + " не найдена в 'База'.");
    return;
  }
  
  // Считываем всю строку из листа "База", чтобы получить все необходимые данные за один раз
  var baseRowData = baseSheet.getRange(baseRow, 1, 1, baseSheet.getLastColumn()).getValues()[0];
  var link = baseRowData[COL_LINK - 1].toString().trim();
  var qty = baseRowData[COL_QUANTITY - 1];
  var deficit = baseRowData[COL_DEFICIT - 1];
  var supplyWay = baseRowData[COL_MODE - 1];

  // Если есть дефицит и способ поставки "В производство", создаем заказ в SD
  if (deficit > 0 && supplyWay === "В производство") {
    var orderId = setSdOrder(ss, itemName, link, qty);
    baseSheet.getRange(baseRow, COL_SD_ORDER_ID).setValue(orderId);
  }

  // Если изменение произошло не на листе "База", сбрасываем исходный чекбокс
  if (sheet.getName() !== baseSheet.getName()){
    range.setValue(false);
  }
  
  // Ставим отметку о выполнении в столбце COL_PICKING на листе "База"
  checkboxOk(baseSheet.getRange(baseRow, COL_PICKING));
  return itemId;
}

function setSdOrder(ss, itemName, link, itemQty) { 
  var comment = itemName + " = " + itemQty;
  var products = [];

  var spec = getSpec(itemName, link);
  Logger.log(`Содержимое спецификации для "${itemName}".`);
  Logger.log(spec);
  for(var i = 0; i < spec.length; i++){
    var sdId = getSdId(ss, spec[i][0]); //определяем идентификатор для СД по имени
    var name = spec[i][0];
    var qty = spec[i][2];
    var amount = qty * itemQty;
    
    // Создаем массив товаров для отправки в СД
    var product = {
      "id": sdId,
      "name": name,
      "costPerItem": "",
      "amount": amount,
      "discount": "",
      "sku": name
    };
    products.push(product);
  }

Logger.log(`Массив товаров для записи в СД, "${itemName}".`);
Logger.log(products);

  // Отправляем запрос на добавление заявки в СД
  var orderId = sendPostRequest(products, comment);
  return orderId;
} 