function handleProduction(ss, sheet, range, row) {
  const baseSheet = ss.getSheetByName("База");
  const itemId = sheet.getRange(row, COL_UUID).getValue().toString().trim();
   
  // Сбрасываем исходный чекбокс, если изменение произошло вне "База"
    if (sheet !== baseSheet){
      range.setValue(false);
      sheet.getRange(row, COL_CHECKBOX).setValue(false);  //Очищаем чекбокс
    }

 // Поиск строки в "База" по itemId
  const baseRow = findRowById(baseSheet, itemId); 
  checkboxOk(baseSheet.getRange(baseRow, COL_PRODUCTION)); //Ставим отметку о выполнении на листе База в столбце Производство
  checkboxOk(baseSheet.getRange(baseRow, COL_QUALITY_CHECK)); //Ставим отметку о выполнении на листе База в столбце ОТК
  return itemId; 
}