/**
 * Обрабатывает изменения, связанные со складом, обновляя данные в листе "База"
 * и создавая/обновляя карточки в Trello.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss - Активная таблица Google Sheets.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - Лист, на котором произошло изменение 
 * @param {GoogleAppsScript.Spreadsheet.Range} range - Диапазон ячейки, которая была изменена.
 * @param {number} row - Номер строки, в которой произошло изменение.
 * @returns {string} UUID обработанного элемента.
 */
function handleStock(ss, sheet, range, row) {
    const baseSheet = ss.getSheetByName("База");
    const itemId = sheet.getRange(row, COL_UUID).getValue().toString().trim();
    const itemName = sheet.getRange(row, COL_ITEMNAME).getValue().toString().trim();
    
    Logger.log(`[${new Date().toLocaleString()}] Запуск обработки склада, строка ${row}, изделие: "${itemName}"`);
    
    // Поиск строки в "База" по itemId
    const baseRow = findRowById(baseSheet, itemId); 

    // Получаем назначение из столбца F (COL_DESTINATION)
    let purpose = baseSheet.getRange(baseRow, COL_DESTINATION).getValue().toString().trim();
    //Убираем UUID
    purpose = purpose.replace(/\s*\([0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}\)/g, '');
    

    // Получаем количество из столбца H
    const qty = baseSheet.getRange(baseRow, COL_QUANTITY).getValue();
        
    const cardId = createTrelloCard(itemName, qty, purpose); // Передаем очищенный 'purpose'
    baseSheet.getRange(baseRow, COL_TRELLO_ID).setValue(cardId); // Записываем идентификатор карточки в Трелло в БД

    // Сбрасываем исходный чекбокс, если изменение произошло вне "База"
    if (sheet != baseSheet){
        range.setValue(false);
    }
    checkboxOk(baseSheet.getRange(baseRow, COL_STOCK)); // Ставим отметку о выполнении на листе База
    return itemId;
}