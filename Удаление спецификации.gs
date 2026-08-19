function handleUndeployment(ss, sheet, row) {
  const itemId       = String(sheet.getRange(row, COL_UUID).getValue()).trim();
  removeSpecification(itemId, ss);
}

/**
 * Запускает процесс удаления спецификации.
 */
function startRemoveSpec() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const parentUuid = "22102587-c8a9-451f-87d4-6b5060e98a0a"; // Пример UUID для запуска
    removeSpecification(parentUuid, ss);
    ss.toast("Откат спецификации завершен.", "Информация", 5); // Уведомление после завершения
}


/**
 * Удаляет спецификацию для указанного изделия (по его UUID) и рекурсивно обрабатывает дочерние элементы.
 * Эта функция будет выполнять один проход по данным для сбора всех изменений.
 *
 * @param {string} initialParentUuid - Уникальный идентификатор корневого изделия, чью спецификацию нужно откатить.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss - Активная таблица.
 */
function removeSpecification(initialParentUuid, ss) {
    Logger.log(`--- Начало removeSpecification для корневого UUID: ${initialParentUuid} ---`);
    const baseSheet = ss.getSheetByName("База");
    if (!baseSheet) {
        Logger.log('Лист "База" не найден.');
        ss.toast("Ошибка: Лист 'База' не найден.", "Ошибка", 5);
        return;
    }

    const destinationColIndex = COL_DESTINATION - 1; 
    const itemnameColIndex = COL_ITEMNAME - 1; 
    const quantityColIndex = COL_QUANTITY - 1; 
    const uuidColIndex = COL_UUID - 1; 

    // 1. Считываем все данные один раз и создаем исходные мапы
    const baseData = baseSheet.getDataRange().getValues();
    
    // Map: UUID -> { index: rowNum, rowData: Array }
    const initialBaseDataMapByUuid = new Map(); 
    // Map: ChildUUID -> Set<ParentUUID> - МОДЕЛЬ ТЕКУЩИХ СВЯЗЕЙ В ПАМЯТИ
    const currentRelationships = new Map(); 
    // Map: ParentUUID -> Set<ChildUUID> - Для быстрого поиска прямых детей
    const parentToChildrenMap = new Map();

    for (let i = 1; i < baseData.length; i++) {
        const childUuid = baseData[i][uuidColIndex]?.toString().trim();
        if (childUuid) {
            initialBaseDataMapByUuid.set(childUuid, { index: i, rowData: baseData[i] });
            
            const destinationCellContent = baseData[i][destinationColIndex]?.toString();
            if (destinationCellContent) {
                const parsedDestinations = parseDestinationCell(destinationCellContent);
                for (const dest of parsedDestinations) {
                    const parentUuid = dest.uuid;
                    // Строим currentRelationships (ChildUUID -> Set<ParentUUID>)
                    if (!currentRelationships.has(childUuid)) {
                        currentRelationships.set(childUuid, new Set());
                    }
                    currentRelationships.get(childUuid).add(parentUuid); 

                    // Строим parentToChildrenMap (ParentUUID -> Set<ChildUUID>)
                    if (!parentToChildrenMap.has(parentUuid)) {
                        parentToChildrenMap.set(parentUuid, new Set());
                    }
                    parentToChildrenMap.get(parentUuid).add(childUuid);
                }
            }
        }
    }

    // Очередь для UUID, которые мы рассматриваем как "удаляемые" или "освобождающиеся"
    const processingQueue = [initialParentUuid];
    // Set для отслеживания UUID, которые уже были добавлены в processingQueue
    const queuedForProcessing = new Set([initialParentUuid]);
    
    // Set для UUID строк, которые будут окончательно удалены из таблицы
    const rowsToBeRemovedUuids = new Set(); 
    
    // Map: ChildUUID -> { newDestination: string, newQty: number, initialRowData: array }
    const pendingUpdates = new Map();

    // Переменная для отслеживания строк, которые будут удалены (с их номерами строк)
    const actualRowsToDelete = []; 

    // Добавляем корневой элемент в список на удаление по умолчанию.
    // Его статус будет пересмотрен в конце, исходя из его назначений.
    rowsToBeRemovedUuids.add(initialParentUuid);
    Logger.log(`Корень ${initialParentUuid} изначально добавлен в список потенциальных удалений.`);


    while (processingQueue.length > 0) {
        const currentUuid = processingQueue.shift(); // Извлекаем текущий UUID для обработки

        Logger.log(`removeSpecification (обработка очереди): Обрабатываем UUID: ${currentUuid}`);

        // Итерируем по всем элементам, которые ССЫЛАЮТСЯ НА currentUuid (они являются его детьми)
        // И используем parentToChildrenMap для эффективного поиска прямых детей
        const childrenOfCurrentUuid = parentToChildrenMap.get(currentUuid) || new Set();

        for (const childUuid of childrenOfCurrentUuid) {
            // Убеждаемся, что мы работаем с актуальной копией связей
            const childParentUuidsSet = currentRelationships.get(childUuid);

            if (!childParentUuidsSet) {
                // Если у элемента нет связей (уже был обработан/удален), пропускаем
                continue;
            }

            // Удаляем ссылку на currentUuid из списка родителей этого дочернего элемента
            if (childParentUuidsSet.has(currentUuid)) {
                Logger.log(`  Элемент "${initialBaseDataMapByUuid.get(childUuid)?.rowData[itemnameColIndex]}" (UUID: ${childUuid}) ссылается на ${currentUuid}. Удаляем эту ссылку.`);
                childParentUuidsSet.delete(currentUuid); 
            } else {
                // Это может произойти, если связь уже была удалена ранее другим путем
                continue; 
            }

            // Получаем исходную строку для этого дочернего элемента
            const childOriginalData = initialBaseDataMapByUuid.get(childUuid)?.rowData;
            if (!childOriginalData) {
                Logger.log(`  Внимание: Исходные данные для дочернего UUID ${childUuid} не найдены. Пропускаем.`);
                continue;
            }

            // Пересоздаем список назначений, используя только оставшиеся родители
            const updatedDestinations = [];
            let totalQtyForChild = 0;

            // Важно: перебираем ТОЛЬКО оставшиеся parentUuid в childParentUuidsSet
            for (const remainingParentUuid of childParentUuidsSet) {
                 const originalDestinationContent = initialBaseDataMapByUuid.get(childUuid)?.rowData[destinationColIndex]?.toString();
                 if (originalDestinationContent) {
                     const originalParsed = parseDestinationCell(originalDestinationContent);
                     const foundOriginalDest = originalParsed.find(d => d.uuid === remainingParentUuid);
                     if (foundOriginalDest) {
                         updatedDestinations.push(foundOriginalDest);
                         totalQtyForChild += foundOriginalDest.quantity;
                     } else {
                         Logger.log(`  Предупреждение: Не найдено оригинальное назначение для UUID ${remainingParentUuid} в ячейке ${childUuid}.`);
                     }
                 }
            }
            
            const currentChildName = childOriginalData[itemnameColIndex];

            if (childParentUuidsSet.size === 0) {
                // Если у дочернего элемента не осталось родителей, он становится бесхозным
                rowsToBeRemovedUuids.add(childUuid);
                Logger.log(`    UUID ${childUuid} ("${currentChildName}") стал бесхозным. Добавлен в список на удаление.`);
                
                // Если этот элемент сам является родителем для кого-то,
                // и он стал бесхозным, то его "дети" тоже потенциально бесхозны.
                // Добавляем его в очередь, чтобы он был обработан как currentUuid
                // для поиска своих собственных дочерних элементов.
                if (!queuedForProcessing.has(childUuid)) {
                    processingQueue.push(childUuid);
                    queuedForProcessing.add(childUuid);
                    Logger.log(`    UUID ${childUuid} ("${currentChildName}") добавлен в очередь для рекурсивной обработки его детей.`);
                }
                // Удаляем из pendingUpdates, если там был, т.к. теперь он удаляется полностью
                pendingUpdates.delete(childUuid);

            } else {
                // Если остались родители, обновляем его назначения и количество
                const newDestinationContent = formatDestinationCell(updatedDestinations);
                pendingUpdates.set(childUuid, {
                    newDestination: newDestinationContent,
                    newQty: totalQtyForChild,
                    initialRowData: childOriginalData 
                });
                Logger.log(`    UUID ${childUuid} ("${currentChildName}") помечен на обновление. Новое назначение: "${newDestinationContent}", Новое количество: ${totalQtyForChild}.`);
                // Убеждаемся, что этот элемент не помечен на удаление, если он вдруг там оказался
                rowsToBeRemovedUuids.delete(childUuid);
            }
        } // Конец for (const childUuid of childrenOfCurrentUuid)
    } // Конец while (processingQueue.length > 0)

    // *** ФИНАЛЬНАЯ ПРОВЕРКА КОРНЕВОГО ЭЛЕМЕНТА (initialParentUuid) ***
    // Корневой элемент должен быть удален, ТОЛЬКО ЕСЛИ он стал полностью бесхозным
    // (т.е. не имеет ни одного родителя в currentRelationships)
    // и не является "сборочной единицей" (т.е. у него было исходное назначение).

    const initialParentRow = initialBaseDataMapByUuid.get(initialParentUuid);
    const initialParentName = initialParentRow?.rowData[itemnameColIndex] || "Неизвестный элемент";
    const initialDestinationContent = initialParentRow?.rowData[destinationColIndex]?.toString().trim();

    // Проверяем, есть ли у корневого элемента какие-либо родители в текущей модели отношений
    const hasAnyRemainingParents = currentRelationships.has(initialParentUuid) && currentRelationships.get(initialParentUuid).size > 0;
    
    if (hasAnyRemainingParents) {
        // Если у корневого элемента остались родители, он не удаляется.
        rowsToBeRemovedUuids.delete(initialParentUuid); 
        Logger.log(`Корневой элемент (UUID: ${initialParentUuid} - "${initialParentName}") имеет оставшиеся назначения и НЕ будет удален.`);
    } else if (!initialDestinationContent) { 
        // Если корневой элемент не имеет родителей (что типично для "сборочной единицы")
        // И у него пустое поле "Назначение" (является "сборочной единицей"), он не удаляется.
        rowsToBeRemovedUuids.delete(initialParentUuid); 
        Logger.log(`Корневой элемент (UUID: ${initialParentUuid} - "${initialParentName}") не имеет назначений (является "сборочной единицей") и НЕ будет удален.`);
    } else {
        // Если у корневого элемента были назначения, и они все были удалены (т.е. он стал бесхозным),
        // тогда он должен быть удален.
        rowsToBeRemovedUuids.add(initialParentUuid); // Добавляем его обратно, если он был удален ранее
        Logger.log(`Корневой элемент (UUID: ${initialParentUuid} - "${initialParentName}") стал бесхозным (все его назначения были удалены) и БУДЕТ удален.`);
    }

    // Собираем actualRowsToDelete на основе rowsToBeRemovedUuids
    for (const uuid of rowsToBeRemovedUuids) {
        const rowInfo = initialBaseDataMapByUuid.get(uuid);
        if (rowInfo) {
            actualRowsToDelete.push({
                rowNum: rowInfo.index + 1, // +1 потому что индексы массива начинаются с 0, а строки в таблице с 1
                name: rowInfo.rowData[itemnameColIndex],
                uuid: uuid
            });
        }
    }

    // Сортируем строки на удаление по убыванию номера строки
    actualRowsToDelete.sort((a, b) => b.rowNum - a.rowNum);

    // Выполняем все удаления
    if (actualRowsToDelete.length > 0) {
        Logger.log(`Будет удалено ${actualRowsToDelete.length} строк(а).`);
        for (const rowInfo of actualRowsToDelete) {
            Logger.log(`    Удаляем строку: ${rowInfo.rowNum} ("${rowInfo.name}", UUID: ${rowInfo.uuid})`);
            try {
                baseSheet.deleteRow(rowInfo.rowNum);
            } catch (e) {
                Logger.log(`ОШИБКА: Не удалось удалить строку ${rowInfo.rowNum} ("${rowInfo.name}", UUID: ${rowInfo.uuid}). Возможно, уже удалена или смещена: ${e.message}`);
            }
        }
    } else {
        Logger.log("Нет строк для полного удаления.");
    }

    // 2. Выполняем все обновления
    // Перечитаем данные после удалений, чтобы получить актуальные индексы
    const currentBaseData = baseSheet.getDataRange().getValues(); 
    const currentBaseDataMapByUuid = new Map(); 
    for (let i = 1; i < currentBaseData.length; i++) {
        const uuid = currentBaseData[i][uuidColIndex]?.toString().trim();
        if (uuid) {
            currentBaseDataMapByUuid.set(uuid, { index: i, rowData: currentBaseData[i] });
        }
    }

    if (pendingUpdates.size > 0) {
        Logger.log(`Будет обновлено ${pendingUpdates.size} строк(а).`);
        
        for (const [uuidToUpdate, updateInfo] of pendingUpdates.entries()) {
            const currentEntry = currentBaseDataMapByUuid.get(uuidToUpdate);
            if (currentEntry) {
                const actualRowNum = currentEntry.index + 1; 
                Logger.log(`    Обновляем строку: ${actualRowNum} ("${updateInfo.initialRowData[itemnameColIndex]}"). Новое назначение: "${updateInfo.newDestination}", Новое количество: ${updateInfo.newQty}.`);
                
                try {
                    baseSheet.getRange(actualRowNum, destinationColIndex + 1).setValue(updateInfo.newDestination);
                    baseSheet.getRange(actualRowNum, quantityColIndex + 1).setValue(updateInfo.newQty);
                } catch (e) {
                    Logger.log(`ОШИБКА: Не удалось обновить ячейки для строки ${actualRowNum} (UUID: ${uuidToUpdate}): ${e.message}`);
                }
            } else {
                Logger.log(`    Пропускаем обновление строки (UUID: ${uuidToUpdate}) ("${updateInfo.initialRowData[itemnameColIndex]}"), так как она была удалена.`);
            }
        }
        Logger.log(`Выполнено обновление ${pendingUpdates.size} строк.`); 
    } else {
        Logger.log("Нет строк для обновления.");
    }
    
    Logger.log(`--- Откат спецификации для корневого UUID: ${initialParentUuid} завершён. ---`);
    ss.toast("Откат спецификации завершен.", "Информация", 5);
}





/**
 * Разбирает содержимое ячейки столбца F (Назначение) на массив объектов.
 * Пример строки: "Р155.02 = 10 (b7ee7162-37a7-49a6-b36c-15f5845b6852)\nР01.001 = 5 (abc-def-ghi)"
 *
 * @param {string} cellValue - Строка из ячейки столбца F.
 * @returns {Array<{name: string, quantity: number, uuid: string}>} Массив разобранных назначений.
 */
function parseDestinationCell(cellValue) {
    if (!cellValue) {
        return [];
    }
    const destinations = [];
    const lines = cellValue.split(/\r?\n/); // Разбиваем по переносам строк

    // Изменено: (\d+(?:[.,]\d+)?) для захвата дробных чисел
    const regex = /(.+)\s*=\s*(\d+(?:[.,]\d+)?)\s*\(([-\w]{36})\)/; 

    for (const line of lines) {
        const trimmedLine = line.trim();
        const match = trimmedLine.match(regex);
        if (match) {
            destinations.push({
                name: match[1].trim(),
                quantity: parseFloat(match[2].replace(',', '.')), // Используем parseFloat и заменяем запятую на точку, если она есть
                uuid: match[3].trim()
            });
        }
    }
    return destinations;
}


/**
 * Форматирует массив объектов назначений обратно в строку для записи в ячейку F.
 *
 * @param {Array<{name: string, quantity: number, uuid: string}>} destinations - Массив назначений.
 * @returns {string} Отформатированная строка.
 */
function formatDestinationCell(destinations) {
    if (!destinations || destinations.length === 0) return "";
    return destinations.map(d => `${d.name} = ${d.quantity} (${d.uuid})`).join("\n");
}