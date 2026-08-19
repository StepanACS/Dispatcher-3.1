function handleAcceptance(ss, sheet, range, row) {
    
  const itemId = sheet.getRange(row, 4).getValue().toString().trim();
  const itemName = sheet.getRange(row, 5).getValue().toString().trim();
  const factQty = sheet.getRange(row, 13).getValue();

  Logger.log(`[${new Date().toLocaleString()}] Запись фактического количества, строка ${row}, изделие: "${itemName}"`);
  
  // Сбрасываем исходный чекбокс, если изменение произошло на листе "База"
  if (sheet.getName == "База"){
    range.setValue(false);  //просто убираем отметку, это действие нельзя выполнять с листа база, потому что нужно указать фактическое количество
    ss.toast("Для выполенния приемки перейдите на соответствующий лист", "Ошибка", 5);
  }else{
    sheet.getRange(row, 11).setValue(false);  //Очищаем чекбокс
    sheet.getRange(row, 13).setValue(""); //Очищаем ячейку с фактическим количеством
    
    // Проверка, указано ли фактическое количество
    if (factQty === "" || factQty === null || factQty === undefined) {
    // Выводим сообщение об ошибке пользователю через toast
      ss.toast("Не указано фактическое количество или неверный формат записи", "Ошибка", 5);
      return; // Останавливаем выполнение функции
    }

    
    const baseSheet = ss.getSheetByName("База");
    const baseRow = findRowById(baseSheet, itemId); // Поиск строки в "База" по itemId
    baseSheet.getRange(baseRow,COL_FACT_QTY ).setValue(factQty); //Записываем фактическое количества в столбец AC в БД
    checkboxOk(baseSheet.getRange(baseRow, COL_ACCEPTANCE)); //Ставим отметку о выполнении на листе База в столбце Приемка
  
    return itemId;
  } 
}