function handleProductionStock(ss, sheet, range, row) {
  
  const baseSheet = ss.getSheetByName("База");
  const itemId = sheet.getRange(row, 4).getValue().toString().trim();
  const itemName = sheet.getRange(row, 5).getValue().toString().trim();
   
  Logger.log(`[${new Date().toLocaleString()}] Запуск обработки склада производства, строка ${row}, изделие: "${itemName}"`);
  
 // Поиск строки в "База" по itemId
  const baseRow = findRowById(baseSheet, itemId); 

  //Получаем идентификатор заказ из SalesDrive из столбца AB
  const sdOrderId = baseSheet.getRange(baseRow, COL_SD_ORDER_ID).getValue();
 
  changeSdStatus(sdOrderId, 23)//меняем статус заказа в СД, 23 - статус Передано в производство

  handleUndeployment(ss, sheet, row);

  // Сбрасываем исходный чекбокс, если изменение произошло вне "База"
  if (sheet != baseSheet){
    range.setValue(false);
  }
  checkboxOk(baseSheet.getRange(baseRow, COL_PRODUCTION_STOCK)); //Ставим отметку о выполнении на листе База
  return itemId;
}




