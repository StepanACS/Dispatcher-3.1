
/**
 * Функция-обработчик для листа "Запуск".
 * Вызывается, если редактирование произошло в нужной строке
 * (или если нужно обработать все строки, если 'rowNumber' не передан).
 * Применяется для добавления изделий из листа "Запуск" в "Базу".
 *
 * @param {number} [rowNumber] - Необязательный параметр. Если передан,
 * обрабатывается только указанная строка. Если не передан, обрабатываются все строки,
 * начиная со второй.
 */
  function handleLaunchEdit(ss, launchSheet, rowNumber) {
    Logger.log("--- Начало handleLaunchEdit ---");
    if (!launchSheet) {
        Logger.log("Лист 'Запуск' не найден.");
        return;
    }
    const baseSheet = ss.getSheetByName("База");
    const actualBaseColumnCount = baseSheet.getLastColumn();
    const startRowForBase = baseSheet.getLastRow() + 1; 
    
    const outputRowsData = []; // Этот массив заполняется строками, которые будут записаны в Базу
    const rowsToDeleteFromLaunch = []; // Строки, которые нужно удалить из Запуска
    
    let dataRange;
    const startProcessRow = 2; // Начинаем обработку листа "Запуск" со второй строки
    let numRowsToProcess;

    let lastColumn = launchSheet.getLastColumn();
    Logger.log(`handleLaunchEdit: Исходное значение launchSheet.getLastColumn() = ${lastColumn}`); 

    if (lastColumn === 0) {
        Logger.log("Лист 'Запуск' пустой или getLastColumn() вернул 0. Устанавливаем lastColumn в 1.");
        lastColumn = 1; 
    }
    Logger.log(`handleLaunchEdit: Итоговое значение lastColumn (после проверки) = ${lastColumn}`);

    if (rowNumber) {
        numRowsToProcess = 1;
        dataRange = launchSheet.getRange(rowNumber, 1, numRowsToProcess, lastColumn); 
    } else {
        if (launchSheet.getLastRow() < startProcessRow) {
            Logger.log("На листе 'Запуск' нет строк данных для обработки (getLastRow < startProcessRow).");
            ss.toast("Нет новых изделий для добавления", "Информация", 3);
            return;
        }
        numRowsToProcess = launchSheet.getLastRow() - startProcessRow + 1;
        
        if (numRowsToProcess <= 0) {
            Logger.log(`handleLaunchEdit: Вычислено numRowsToProcess <= 0 (${numRowsToProcess}). Нет строк для обработки.`);
            ss.toast("Нет новых изделий для добавления", "Информация", 3);
            return;
        }

        Logger.log(`handleLaunchEdit: numRowsToProcess (else branch) = ${numRowsToProcess}`);
        dataRange = launchSheet.getRange(startProcessRow, 1, numRowsToProcess, lastColumn); 
    }
    
    const data = dataRange.getValues();

    for (let i = 0; i < data.length; i++) {
        const row = data[i];
        // Проверяем, что есть данные в колонках A и C (название и количество)
        if (String(row[0]).trim() !== "" && String(row[2]).trim() !== "") {
            const actualRowInLaunchSheet = (rowNumber ? rowNumber : startProcessRow + i);
            // futureBaseRowIndex будет корректным, так как outputRowsData.length увеличивается
            // на каждой итерации, отражая будущую позицию в массиве.
            const futureBaseRowIndex = startRowForBase + outputRowsData.length; 

            // Передаем actualBaseColumnCount в buildBaseRow
            outputRowsData.push(buildBaseRow(row, futureBaseRowIndex, actualBaseColumnCount));
            rowsToDeleteFromLaunch.push(actualRowInLaunchSheet);
        }
    }
    
    if (outputRowsData.length === 0) {
        Logger.log("Нет строк для добавления в базу из листа 'Запуск'.");
        ss.toast("Нет новых изделий для добавления", "Информация", 3);
        return;
    }
    
    // --- КОРРЕКТНЫЙ ВЫЗОВ writeRowsToSheetForLaunch ---
    // startRowForBase уже содержит getLastRow() + 1
    writeRowsToSheetForLaunch(baseSheet, startRowForBase, outputRowsData, actualBaseColumnCount); 

    // Логика применения валидации чекбоксов
    if (outputRowsData.length > 0) {
        const firstNewRow = startRowForBase; // Первая новая строка - это startRowForBase
        const lastNewRow = startRowForBase + outputRowsData.length - 1; // Последняя новая строка

        const checkboxRange = baseSheet.getRange(
            firstNewRow, COL_DESIGN, 
            outputRowsData.length, // Количество строк для валидации
            COL_READY - COL_DESIGN + 1
        );

        const rule = SpreadsheetApp.newDataValidation()
            .requireCheckbox()
            .setAllowInvalid(false)
            .build();
        
        checkboxRange.setDataValidation(rule);
        Logger.log(`Применено форматирование чекбоксам к диапазону R${firstNewRow}C${COL_DESIGN}:R${lastNewRow}C${COL_READY}`);
    }
    
    // Сортируем строки для удаления в порядке убывания, чтобы избежать проблем со смещением индексов
    rowsToDeleteFromLaunch.sort((a, b) => b - a);

    Logger.log(`Будет удалено ${rowsToDeleteFromLaunch.length} строк из листа 'Запуск'.`);
    
    for (const rowIndex of rowsToDeleteFromLaunch) {
        // Проверка: убеждаемся, что rowIndex является действительным номером строки на листе
        if (rowIndex > 0 && rowIndex <= launchSheet.getLastRow()) { 
            try {
                launchSheet.deleteRow(rowIndex);
                Logger.log(`Удалена строка ${rowIndex} из листа 'Запуск'.`);
            } catch (e) {
                Logger.log(`Внимание: Не удалось удалить строку ${rowIndex} из листа 'Запуск'. Возможно, она уже отсутствует или произошла другая ошибка: ${e.message}`);
            }
        } else {
            Logger.log(`Внимание: Пропущено удаление строки ${rowIndex} из листа 'Запуск', так как она вышла за границы (rowIndex: ${rowIndex}, getLastRow(): ${launchSheet.getLastRow()}).`);
        }
    }
    Logger.log("Удаление строк из листа 'Запуск' завершено.");
    
    ss.toast("Изделие(я) добавлено(ы) в базу", "Информация", 5);
    //var templateRow = sheet.getLastRow();
                    //testEnsureFreeRows();
                    ensureFreeRows(sheet, 18, 50);
                    //ensureFreeRows(sheet, templateRow, 30)
    Logger.log("--- Завершено handleLaunchEdit ---");
}

/**
 * Функция buildBaseRow(rowData, futureIndex) формирует строку для записи в лист "База".
 *
 * @param {Array<any>} rowData - Массив данных одной строки с листа "Запуск".
 * Ожидаемый порядок: [itemName, destination, quantity, note]
 * @param {number} futureIndex - Будущий номер строки в листе "База", куда будет записана эта строка. Нужен для формул.
 * @returns {Array<any>} Сформированная строка-массив с данными и формулами.
 */
function buildBaseRow(rowData, futureIndex, totalColumnCount) {
    // Если buildBaseRowTemplate ожидает 1-индексированный номер строки, передайте futureIndex.
    // Если она ожидает 0-индексированный, то futureIndex - 1 было бы верно.
    // Исходя из стандартной практики GS API, обычно функции ожидают 1-индексированные номера.
    const newRow = buildBaseRowTemplate(futureIndex, totalColumnCount); // <-- ИСПРАВЛЕНИЕ: ПЕРЕДАЕМ futureIndex
    
    // Данные из листа "Запуск"
    const launchItemName = String(rowData[0]).trim();    // Запуск: A
    const launchDestination = String(rowData[1]).trim(); // Запуск: B
    const launchQuantity = rowData[2];                 // Запуск: C
    const launchNote = String(rowData[3]).trim();        // Запуск: D

    let formattedItemName = launchItemName;
    if (formattedItemName.length > 0) {
        formattedItemName = formattedItemName.charAt(0).toUpperCase() + formattedItemName.slice(1);
    }
    let formattedDestination = launchDestination;
    if (formattedDestination.length > 0) {
        formattedDestination = formattedDestination.charAt(0).toUpperCase() + formattedDestination.slice(1);
    }
    
    // Заполняем специфичные для buildBaseRow значения
    newRow[COL_DATE - 1] = new Date(); // Используем текущую дату
    newRow[COL_ITEMNAME - 1] = formattedItemName;
    newRow[COL_DESTINATION - 1] = formattedDestination;
    newRow[COL_NOTE - 1] = launchNote;
    newRow[COL_QUANTITY - 1] = launchQuantity;

    return newRow;
}

    
function writeRowsToSheetForLaunch(sheet, startRow, rows, columnCount) {
    // Добавим более информативный лог для отладки
    if (!rows || rows.length === 0) {
        Logger.log("writeRowsToSheetForLaunch: Нет строк для записи. Функция завершена.");
        return;
    }

     
    Logger.log(`writeRowsToSheetForLaunch: Запись ${rows.length} строк, начиная со строки ${startRow}, до столбца ${columnCount}.`);

    // Вычисляем, сколько строк нужно добавить на лист, если текущих недостаточно
    const maxRows = sheet.getMaxRows();
    const rowsNeeded = (startRow + rows.length - 1) - maxRows; 
    if (rowsNeeded > 0) {
        sheet.insertRowsAfter(maxRows, rowsNeeded); 
        Logger.log(`writeRowsToSheetForLaunch: Добавлено ${rowsNeeded} новых строк, т.к. текущих (${maxRows}) недостаточно.`);
    }
    
    try {
        const targetRange = sheet.getRange(startRow, 1, rows.length, columnCount);
        targetRange.setValues(rows);
        Logger.log("writeRowsToSheetForLaunch: Данные успешно записаны.");
    } catch (e) {
        Logger.log(`writeRowsToSheetForLaunch: Ошибка при записи данных: ${e.message}`);
        throw e; // Перебрасываем ошибку для дальнейшей обработки
    }
}
