/**
 * Проверяет и объединяет дублирующиеся строки в листе "База"
 * согласно определенным условиям и правилам слияния, сохраняя форматирование.
 * Строки обрабатываются до столбца columnCount.
 * Строки с любой отметкой (кроме булевого false) в столбце U (COL_PRODUCTION) исключаются из обработки.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} baseSheet - Объект листа "База".
 * @param {number} columnCount - Общее количество столбцов для обработки (фактически getLastColumn()).
 */
function checkAndMergeDuplicates(baseSheet, columnCount) {
    Logger.log("--- Начало checkAndMergeDuplicates ---");
    const sheet = baseSheet; // Используем переданный лист

    if (!sheet) {
        Logger.log("Лист 'База' не найден. (Проблема с передачей параметра baseSheet)");
        return;
    }

    const startRow = 2; // Предполагаем, что первая строка — заголовки
    let effectiveLastRow = sheet.getLastRow();
    const LAST_COLUMN_TO_PROCESS = columnCount; // Используем переданный columnCount

    if (effectiveLastRow < startRow) {
        Logger.log("На листе 'База' нет данных для обработки (только заголовки или пусто).");
        Logger.log("--- Завершено checkAndMergeDuplicates (нет данных) ---");
        return;
    }

    const dataRange = sheet.getRange(startRow, 1, effectiveLastRow - startRow + 1, LAST_COLUMN_TO_PROCESS);
    const values = dataRange.getValues();
    const formulas = dataRange.getFormulas();

    const uniqueItemsMap = new Map();
    const rowsToDelete = [];

    for (let i = 0; i < values.length; i++) {
        const currentRowValues = values[i];
        const currentRowFormulas = formulas[i];

        const productionMarker = currentRowValues[COL_PRODUCTION - 1];

        if (productionMarker !== false && productionMarker !== null && productionMarker !== "") {
            Logger.log(`Строка ${i + startRow} пропущена: обнаружена отметка в столбце "${COL_PRODUCTION}" (Отметка о выполнении задач производства).`);
            continue;
        }

        const itemName = currentRowValues[COL_ITEMNAME - 1];
        const destination = currentRowValues[COL_DESTINATION - 1];
        const quantity = currentRowValues[COL_QUANTITY - 1];
        const checkboxQ = currentRowValues[COL_PICKER - 1];

        if (itemName && checkboxQ === false && destination) {
            if (uniqueItemsMap.has(itemName)) {
                const existingEntry = uniqueItemsMap.get(itemName);
                const firstRowIndexInValues = existingEntry.firstRowIndexInValues;

                Logger.log(`Обнаружен дубликат для "${itemName}" (строка ${i + startRow}). Объединяем со строкой ${firstRowIndexInValues + startRow}.`);

                let existingDestination = values[firstRowIndexInValues][COL_DESTINATION - 1];
                if (typeof existingDestination !== 'string') existingDestination = String(existingDestination || '');
                if (!existingDestination.includes(destination)) {
                    existingDestination = existingDestination ? existingDestination + '\n' + destination : destination;
                    values[firstRowIndexInValues][COL_DESTINATION - 1] = existingDestination;
                    formulas[firstRowIndexInValues][COL_DESTINATION - 1] = '';
                    existingEntry.needsUpdate = true;
                }

                let existingQuantity = values[firstRowIndexInValues][COL_QUANTITY - 1];
                existingQuantity = typeof existingQuantity === 'number' ? existingQuantity : parseFloat(existingQuantity) || 0;
                let currentQuantity = typeof quantity === 'number' ? quantity : parseFloat(quantity) || 0;

                const newQuantity = existingQuantity + currentQuantity;
                if (newQuantity !== existingQuantity) {
                    values[firstRowIndexInValues][COL_QUANTITY - 1] = newQuantity;
                    formulas[firstRowIndexInValues][COL_QUANTITY - 1] = '';
                    existingEntry.needsUpdate = true;
                }

                rowsToDelete.push(i + startRow);

            } else {
                uniqueItemsMap.set(itemName, { firstRowIndexInValues: i, needsUpdate: false });
            }
        }
    }

    for (const [itemName, entry] of uniqueItemsMap.entries()) {
        if (entry.needsUpdate) {
            const rowInSheet = entry.firstRowIndexInValues + startRow;
            const rowValuesToUpdate = values[entry.firstRowIndexInValues];
            const rowFormulasToUpdate = formulas[entry.firstRowIndexInValues];

            const rowDataToWrite = [];
            for (let j = 0; j < LAST_COLUMN_TO_PROCESS; j++) {
                rowDataToWrite[j] = rowFormulasToUpdate[j] || rowValuesToUpdate[j];
            }

            Logger.log(`Обновление объединенной строки для "${itemName}" (строка ${rowInSheet}).`);
            sheet.getRange(rowInSheet, 1, 1, LAST_COLUMN_TO_PROCESS).setValues([rowDataToWrite]);
        }
    }

    rowsToDelete.sort((a, b) => b - a);

    if (rowsToDelete.length > 0) {
        Logger.log(`Начинаем удаление ${rowsToDelete.length} дубликатов.`);
        for (const rowIndex of rowsToDelete) {
            Logger.log(`Удаление строки: ${rowIndex}`);
            sheet.deleteRow(rowIndex);
        }
        Logger.log("Удаление дубликатов завершено.");
    } else {
        Logger.log("Дубликаты для удаления не найдены.");
    }

    Logger.log("--- Завершено checkAndMergeDuplicates ---");
}
