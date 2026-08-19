/**
 * Вспомогательная функция для получения экземпляра Firestore.
 * Настраивает подключение к базе данных Firestore, используя предоставленные учетные данные.
 *
 * @returns {GoogleAppsScript.Firestore.Firestore} Экземпляр клиента Firestore, готовый к использованию.
 */
function getFirestoreDb() {
  // Исправлен порядок аргументов согласно вашему тесту
  return FirestoreApp.getFirestore(CLIENT_EMAIL, PRIVATE_KEY, PROJECT_ID);
}

/*******************************************************************************
 * Модуль 1: Диспетчер развертывания спецификации
 ******************************************************************************/

/**
 * Обрабатывает развертывание спецификации изделия на лист "База" или подтверждает,
 * что спецификация уже развернута.
 *
 * @param {string} itemUUID Уникальный идентификатор изделия.
 * @param {string} itemName Имя изделия.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss Активная таблица.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} baseSheet Лист "База".
 * @param {number} rowFromDeploymentSheet Номер строки на листе "Развертывание спецификации", откуда был вызван триггер.
 * @param {number} initialQuantity Исходное количество изделия.
 * @returns {string|null} itemUUID при успешном выполнении или null в случае ошибки.
 */
function handleDeployment(itemUUID, itemName, ss, baseSheet, rowFromDeploymentSheet, initialQuantity) {
    Logger.log(`--- Начало handleDeployment для строки (на листе "Развертывание спецификации"): ${rowFromDeploymentSheet}, UUID: ${itemUUID} ---`);

    // --- Валидация входных данных ---
    if (!itemName || typeof itemName !== 'string' || itemName.trim() === '') {
        Logger.log(`handleDeployment: Имя изделия не определено или некорректно. itemName: "${itemName}"`);
        ss.toast("Ошибка: Не могу развернуть спецификацию, имя изделия не определено.", "Ошибка", 5);
        return null;
    }
    itemName = itemName.trim();

    if (!itemUUID || typeof itemUUID !== 'string' || itemUUID.trim() === '') {
        Logger.log(`handleDeployment: UUID изделия не определен или некорректен. itemUUID: "${itemUUID}"`);
        ss.toast("Ошибка: Не могу развернуть спецификацию, UUID изделия не определен.", "Ошибка", 5);
        return null;
    }

    // Определяем номер строки на листе "База", соответствующей текущему itemUUID,
    // которую нужно исключить из проверки "уже развернута".
    const initialRowOnBaseSheetForMainItem = findRowById(baseSheet, itemUUID);

    // --- 🔍 Использование новой функции для проверки развернутости, исключая исходную строку ---
    const deploymentStatus = isSpecificationDeployed(baseSheet, itemUUID, initialRowOnBaseSheetForMainItem || 0);

    const isAlreadyDeployed = deploymentStatus.isDeployed;

    if (isAlreadyDeployed) {
        Logger.log(`handleDeployment: Спецификация "${itemName}" уже развернута. Пропускаем повтор.`);
        ss.toast(`Спецификация "${itemName}" уже развернута.`, "Информация", 5);

        // Если уже развернуто, ставим отметку в чекбоксе на листе "База"
        if (initialRowOnBaseSheetForMainItem !== null) {
            Logger.log(`handleDeployment: Установка чекбокса для уже развернутой спецификации в строке ${initialRowOnBaseSheetForMainItem}, столбец ${COL_DEPLOYMENT} на листе "База".`);
            checkboxOk(baseSheet.getRange(initialRowOnBaseSheetForMainItem, COL_DEPLOYMENT));
        } else {
            Logger.log(`handleDeployment: ВНИМАНИЕ: Спецификация "${itemName}" помечена как развернутая, но основное изделие с UUID "${itemUUID}" не найдено на листе "База" для установки чекбокса.`);
        }
        return itemUUID;
    }

    // --- Основная логика развертывания, если спецификация НЕ развернута ---
    const productSpec = getOrUploadProductSpecFromFirestore(itemName, ss);
    if (!productSpec) {
        Logger.log(`handleDeployment: Не удалось получить данные спецификации для "${itemName}".`);
        return null;
    }

    const firestoreSpecData = productSpec.data;

    Logger.log(`Начинаем запись спецификации для "${itemName}" на лист "База" из Firestore.`);

    let deploymentSuccess = false;
    try {
        writeSpecToBaseSheetFromFirestore(itemUUID, itemName, firestoreSpecData, ss, baseSheet, initialQuantity);
        Logger.log(`handleDeployment: writeSpecToBaseSheetFromFirestore успешно завершена для "${itemName}".`);
        deploymentSuccess = true;
    } catch (e) {
        Logger.log(`handleDeployment: Общая ошибка при записи спецификации на лист "База": ${e.message}`);
        ss.toast(`Произошла ошибка при развертывании спецификации: ${e.message}`, "Ошибка", 5);
    }

    // --- Установка чекбокса и постобработка после успешного развертывания ---
    if (deploymentSuccess) {
        const finalTargetRowOnBaseSheet = findRowById(baseSheet, itemUUID);

        if (finalTargetRowOnBaseSheet !== null) {
            Logger.log(`handleDeployment: Установка чекбокса после успешного развертывания в строке ${finalTargetRowOnBaseSheet}, столбец ${COL_DEPLOYMENT} на листе "База".`);
            checkboxOk(baseSheet.getRange(finalTargetRowOnBaseSheet, COL_DEPLOYMENT));
            

            // Получаем количество столбцов перед вызовом checkAndMergeDuplicates
            const actualBaseColumnCount = baseSheet.getLastColumn();
            Logger.log(`handleDeployment: Вызов checkAndMergeDuplicates с actualBaseColumnCount: ${actualBaseColumnCount}`);
            checkAndMergeDuplicates(baseSheet, actualBaseColumnCount);

            ss.toast(`Спецификация для "${itemName}" успешно обновлена.`, "Успех", 5);
            Logger.log(`--- Завершено handleDeployment для "${itemName}" ---`);
            return itemUUID;
        } else {
            Logger.log(`handleDeployment: КРИТИЧЕСКАЯ ОШИБКА: Не удалось найти строку основного изделия "${itemUUID}" на листе "База" после успешного развертывания для установки чекбокса.`);
            return null;
        }
    } else {
        return null;
    }
}

/**
 * Проверяет, развернута ли спецификация для данного изделия на листе "База",
 * исключая указанную строку из проверки.
 *
 * Спецификация считается развернутой, если ее UUID найден в столбце COL_UUID
 * или в ссылках на родительский UUID в столбце COL_DESTINATION
 * в ЛЮБОЙ ДРУГОЙ строке, кроме excludeRow.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} baseSheet Лист "База".
 * @param {string} itemUUID Уникальный идентификатор изделия.
 * @param {number} excludeRow Номер строки (1-индексированный) на листе "База", которую нужно исключить из поиска.
 * @returns {{isDeployed: boolean, targetRowOnBaseSheet: number|null}} Объект с флагом развернутости и номером строки главного изделия.
 */
function isSpecificationDeployed(baseSheet, itemUUID, excludeRow) {
    Logger.log(`isSpecificationDeployed: Начинаю проверку для UUID: "${itemUUID}" на листе "${baseSheet.getName()}", исключая строку ${excludeRow}.`);

    const data = baseSheet.getDataRange().getValues();
    const trimmedItemUUID = itemUUID.toString().trim();

    let foundMainItem = false;
    let targetRow = null;
    let foundInDestination = false;

    for (let i = 1; i < data.length; i++) {
        const currentRow = i + 1;

        if (currentRow === excludeRow) {
            continue;
        }

        const uuidInCell = data[i][COL_UUID - 1];
        const destinationTextInCell = data[i][COL_DESTINATION - 1];

        const trimmedUuidInCell = uuidInCell?.toString().trim();

        if (trimmedUuidInCell === trimmedItemUUID) {
            targetRow = currentRow;
            foundMainItem = true;
            Logger.log(`isSpecificationDeployed: Главное изделие с UUID "${itemUUID}" найдено вне excludeRow в строке ${targetRow}.`);
            break;
        }

        const parentUUIDMatch = (destinationTextInCell || '').toString().trim().match(/\(([^)]+)\)/);
        const parentUUIDInDestination = parentUUIDMatch ? parentUUIDMatch[1] : null;

        if (parentUUIDInDestination === trimmedItemUUID) {
            foundInDestination = true;
            Logger.log(`isSpecificationDeployed: Найден дочерний элемент, ссылающийся на UUID "${itemUUID}" вне excludeRow в строке ${currentRow}.`);
            break;
        }
    }

    const isDeployed = foundMainItem || foundInDestination;

    Logger.log(`isSpecificationDeployed: Завершено для UUID "${itemUUID}". Развернута: ${isDeployed}, Строка главного изделия (вне excludeRow): ${targetRow}.`);
    return { isDeployed: isDeployed, targetRowOnBaseSheet: targetRow };
}

/**
 * Ищет спецификацию продукта в Firestore по его имени.
 * Если спецификация найдена, она возвращается.
 * Если не найдена, функция пытается загрузить её из соответствующей Google Таблицы
 * (используя "Каталог ссылок") и затем прочитать из Firestore после загрузки.
 * Предоставляет обратную связь пользователю через toast-сообщения.
 *
 * @param {string} itemName - Название изделия, спецификацию которого нужно найти или загрузить.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss - Объект активной таблицы для вывода toast-сообщений.
 * @returns {{docId: string, data: Object}|null} Объект с ID документа Firestore и его данными, или null в случае неудачи.
 */
function getOrUploadProductSpecFromFirestore(itemName, ss) {
    Logger.log(`getOrUploadProductSpecFromFirestore: Поиск/загрузка спецификации для "${itemName}".`);

    try {
        const firestore = getFirestoreDb();

        let queryResults = firestore.query(PRODUCTS_COLLECTION)
                                        .Where('name', '==', itemName)
                                        .Limit(1)
                                        .Execute();

        if (queryResults && queryResults.length > 0) {
            const doc = queryResults[0];
            Logger.log(`Спецификация для "${itemName}" найдена в Firestore. ID: ${doc.path.split('/').pop()}`);
            return { docId: doc.path.split('/').pop(), data: doc.obj };
        } else {
            Logger.log(`Спецификация для "${itemName}" не найдена в Firestore. Запускаем загрузку из Google Sheets...`);
            ss.toast(`Изделие "${itemName}" отсутствует в базе данных. Загружаю...`, "Информация", 5);

            const uploadedDocId = uploadSpecToFirestoreFromSheet(itemName);

            if (uploadedDocId) {
                const loadedDoc = getFirestoreDocument(PRODUCTS_COLLECTION, uploadedDocId);
                if (loadedDoc && loadedDoc.obj) {
                    Logger.log(`Спецификация для "${itemName}" успешно загружена и прочитана. ID: ${uploadedDocId}.`);
                    ss.toast(`Изделие "${itemName}" загружено в базу данных.`, "Информация", 5);
                    return { docId: uploadedDocId, data: loadedDoc.obj };
                } else {
                    Logger.log(`Ошибка: Не удалось прочитать только что загруженный документ для "${itemName}" с ID: ${uploadedDocId}.`);
                    ss.toast(`Ошибка: Не удалось прочитать загруженную спецификацию для "${itemName}".`, "Ошибка", 5);
                    return null;
                }
            } else {
                Logger.log(`Ошибка: Не удалось загрузить спецификацию для "${itemName}" из Google Sheets в Firestore.`);
                ss.toast(`Ошибка: Не удалось загрузить спецификацию для "${itemName}" в базу данных. Проверьте "Каталог ссылок".`, "Ошибка", 5);
                return null;
            }
        }
    } catch (e) {
        Logger.log(`Критическая ошибка при поиске/загрузке/чтении спецификации для "${itemName}" в Firestore: ${e.message}`);
        ss.toast(`Критическая ошибка при работе с базой данных для "${itemName}".`, "Ошибка", 5);
        return null;
    }
}

/*******************************************************************************
 * Модуль 2: Управление документами в Firestore (products collection)
 ******************************************************************************/

/**
 * Ищет документ компонента/изделия в коллекции PRODUCTS_COLLECTION в Firestore по его имени.
 * Если документ найден, возвращает его ID.
 * Если документ не найден, создает новый документ для данного компонента, инициализируя
 * его базовыми полями ('name', 'dateCreated', 'lastSpecUpdated', 'specification'),
 * и возвращает ID нового документа.
 *
 * Предоставляет опциональную возможность принудительно сбросить поля 'lastSpecUpdated'
 * и 'specification' существующего документа на дефолтные значения, если он найден.
 * Это полезно для сценариев, где требуется полная перезапись спецификации, а не
 * её сохранение при наличии в базе.
 *
 * @param {string} componentName - Название компонента/изделия, которое нужно найти или создать.
 * @param {boolean} [overwriteIfFound=false] - Флаг, указывающий, нужно ли перезаписать
 * поля 'lastSpecUpdated' и 'specification' на пустые значения, если документ уже существует.
 * По умолчанию `false` (сохранять существующие данные).
 * @returns {string|null} ID документа компонента в Firestore, или null в случае ошибки.
 */
function findOrCreateComponentDocId(componentName, overwriteIfFound = false) {
    if (!componentName) {
        Logger.log('Пустое имя компонента передано в findOrCreateComponentDocId.');
        return null;
    }

    try {
        let docId = searchFirestoreDocument(PRODUCTS_COLLECTION, "name", componentName);

        if (docId) {
            Logger.log(`Найден компонент "${componentName}" с ID: ${docId}`);

            if (overwriteIfFound) {
                Logger.log(`Перезаписываем спецификацию для компонента "${componentName}" с ID: ${docId}.`);
                const resetData = {
                    lastSpecUpdated: new Date().toISOString(),
                    specification: []
                };
                const updateSuccess = updateFirestoreDocument(PRODUCTS_COLLECTION, docId, resetData, true);
                if (!updateSuccess) {
                    Logger.log(`Ошибка: Не удалось перезаписать спецификацию для "${componentName}" с ID: ${docId}.`);
                    return null;
                }
            }
            return docId;
        } else {
            const now = new Date().toISOString();
            const newDocData = {
                name: componentName,
                dateCreated: now,
                lastSpecUpdated: null,
                specification: []
            };

            const newDocId = writeFirestoreDocument(PRODUCTS_COLLECTION, newDocData);
            if (newDocId) {
                Logger.log(`Создан новый компонент "${componentName}" с ID: ${newDocId}`);
                return newDocId;
            } else {
                Logger.log(`Ошибка: Не удалось создать новый компонент "${componentName}".`);
                return null;
            }
        }
    } catch (e) {
        Logger.log(`Общая ошибка в findOrCreateComponentDocId для "${componentName}": ${e.message}`);
        Logger.log(`Подробности ошибки: ${e.stack}`);
        return null;
    }
}

/**
 * Рекурсивно обрабатывает спецификацию компонента/изделия, считывая её из Google Sheets
 * и сохраняя (или обновляя) в Firestore.
 * Функция проходит по вложенным компонентам, вызывая себя для каждого из них,
 * тем самым формируя полную иерархию спецификации в Firestore.
 * Если ссылка на спецификацию не найдена, или возникает ошибка при чтении таблицы,
 * компонент все равно обновляется в Firestore с актуальной датой 'lastSpecUpdated',
 * но без вложенной спецификации.
 *
 * @param {string} currentCompName - Название текущего компонента/изделия, для которого обрабатывается спецификация.
 * @param {number} depth - Текущая глубина рекурсии (для предотвращения бесконечных циклов).
 * @param {Map<string, string>} catalogMap - Карта, содержащая соответствия названий изделий
 * ссылкам на их Google Таблицы спецификаций (полученная из "Каталога ссылок").
 * @returns {string|null} ID документа текущего компонента в Firestore, или null в случае критической ошибки.
 */
function processAndGetFirestoreComponentId(currentCompName, depth, catalogMap) {
    if (depth > MAX_SPEC_DEPTH) {
        Logger.log(`Превышена максимальная глубина рекурсии для "${currentCompName}".`);
        return null;
    }

    Logger.log(`Обработка компонента: ${currentCompName} (глубина ${depth})`);
    const componentDocId = findOrCreateComponentDocId(currentCompName);
    if (!componentDocId) {
        Logger.log(`Не удалось получить/создать ID для "${currentCompName}".`);
        return null;
    }

    const componentSpecLink = catalogMap.get(currentCompName);
    if (!componentSpecLink) {
        Logger.log(`Ссылка на спецификацию для "${currentCompName}" не найдена. Обновляем lastSpecUpdated.`);
        const updateSuccess = updateFirestoreDocument(PRODUCTS_COLLECTION, componentDocId, {
            name: currentCompName,
            lastSpecUpdated: new Date().toISOString()
        }, true);
        if (!updateSuccess) {
            Logger.log(`Ошибка: Не удалось обновить lastSpecUpdated для "${currentCompName}".`);
        }
        return componentDocId;
    }

    const sheetId = extractSheetId(componentSpecLink);
    if (!sheetId) {
        Logger.log(`Не удалось извлечь ID таблицы из ссылки для "${currentCompName}". Возвращаем ID: ${componentDocId}`);
        return componentDocId;
    }

    let sheetDoc;
    try {
        sheetDoc = SpreadsheetApp.openById(sheetId);
    } catch (e) {
        Logger.log(`Ошибка открытия Google Таблицы для "${currentCompName}" (ID: ${sheetId}): ${e.message}`);
        Logger.log(`Подробности ошибки: ${e.stack}`);
        return componentDocId;
    }

    const sheet = sheetDoc.getSheetByName(currentCompName);
    if (!sheet) {
        Logger.log(`Лист "${currentCompName}" не найден в таблице ID: ${sheetId}. Возвращаем ID: ${componentDocId}`);
        return componentDocId;
    }

    const rawData = sheet.getRange(5, 2, sheet.getLastRow() - 4, 3).getValues();
    const itogoIndex = rawData.findIndex(row => (row[0] || '').toString().toLowerCase() === 'итого');
    if (itogoIndex === -1) {
        Logger.log(`Строка "итого" не найдена в спецификации "${currentCompName}". Возвращаем ID: ${componentDocId}`);
        return componentDocId;
    }

    const spec = [];
    for (const row of rawData.slice(0, itogoIndex)) {
        const name = row[0]?.toString().trim();
        const qty = parseFloat(row[2]);
        if (!name || isNaN(qty)) {
            Logger.log(`Пропущен некорректный элемент спецификации: Имя="${name}", Кол-во="${row[2]}"`);
            continue;
        }

        const childId = processAndGetFirestoreComponentId(name, depth + 1, catalogMap);
        if (childId) {
            spec.push({
                componentName: name,
                quantityPerParent: qty,
                componentDocId: childId
            });
        }
    }

    Logger.log(`Обновляем спецификацию для "${PRODUCTS_COLLECTION}/${componentDocId}"...`);
    const updateSpecSuccess = updateFirestoreDocument(PRODUCTS_COLLECTION, componentDocId, {
        name: currentCompName,
        specification: spec,
        lastSpecUpdated: new Date().toISOString()
    }, true);
    if (updateSpecSuccess) {
        Logger.log(`Спецификация для "${PRODUCTS_COLLECTION}/${componentDocId}" обновлена.`);
    } else {
        Logger.log(`Ошибка: Не удалось обновить спецификацию для "${PRODUCTS_COLLECTION}/${componentDocId}".`);
    }

    return componentDocId;
}

/**
 * Получает карту ссылок на спецификации из листа "Каталог ссылок".
 *
 * @returns {Map<string, string>} Карта, где ключ - название изделия, значение - ссылка на его спецификацию.
 */
function getCatalogMap() {
    const catalogMap = new Map();
    try {
        const ssCatalog = SpreadsheetApp.openById("1JTr3AFm5Q1ycipTsrC5qY-SGKWY3v_w_2Jgfhq2xHgU");
        const catalogSheet = ssCatalog.getSheetByName("Каталог ссылок");
        if (!catalogSheet) {
            Logger.log("Лист 'Каталог ссылок' не найден.");
            return catalogMap;
        }

        const data = catalogSheet.getDataRange().getValues();
        for (let i = 1; i < data.length; i++) {
            const name = data[i][CATALOG_COL_ITEMNAME - 1]?.toString().trim();
            const link = data[i][CATALOG_COL_LINK - 1]?.toString().trim();
            if (name && link) catalogMap.set(name, link);
        }
    } catch (e) {
        Logger.log(`Ошибка в getCatalogMap: ${e.message}`);
    }
    return catalogMap;
}



/**
 * Запускает процесс загрузки полной иерархической спецификации изделия из Google Sheets
 * в коллекцию Firestore 'PRODUCTS_COLLECTION'.
 * Функция считывает ссылки на спецификации из "Каталога ссылок" и инициирует рекурсивную
 * обработку, начиная с основного изделия.
 *
 * @param {string} productName - Название основного изделия, для которого загружается спецификация.
 * @returns {string|null} ID документа основного изделия в Firestore, или null в случае ошибки.
 */
function uploadSpecToFirestoreFromSheet(productName) {
    Logger.log(`--- Начало uploadSpecToFirestoreFromSheet для изделия: ${productName} ---`);
    if (!productName) {
        Logger.log('Название продукта не указано.');
        return null;
    }

    // Чтение "Каталога ссылок"
    const catalogMap = getCatalogMap(); // Используем новую, общую функцию
    if (catalogMap.size === 0) {
        Logger.log('Каталог ссылок пуст или не удалось его загрузить.');
        return null;
    }

    // --- Основное изделие (начало рекурсии) ---
    const mainDocId = findOrCreateComponentDocId(productName);
    if (!mainDocId) {
        Logger.log(`Не удалось получить/создать ID для основного изделия "${productName}".`);
        return null;
    }

    const finalMainDocId = processAndGetFirestoreComponentId(productName, 1, catalogMap);

    if (finalMainDocId) {
        Logger.log(`Спецификация изделия "${productName}" успешно загружена. ID: ${finalMainDocId}`);
    } else {
        Logger.log(`Ошибка: Не удалось загрузить спецификацию основного изделия "${productName}".`);
    }

    Logger.log(`--- Завершение uploadSpecToFirestoreFromSheet для изделия: ${productName} ---`);
    return finalMainDocId;
}



/*******************************************************************************
 * Модуль 3: Запись на лист "База" из Firestore
 ******************************************************************************/
/**
 * Записывает развернутую спецификацию изделия из Firestore на лист "База" в Google Таблицах.
 * Функция обновляет существующую заглавную строку основного изделия (если найдена)
 * или добавляет новую, а затем вставляет новую развернутую структуру спецификации,
 * динамически добавляя необходимое количество строк на лист "База" после заглавной строки.
 * Она также учитывает количество на остатках и покупные/непокупные статусы для расчета дефицита и проставления чекбосов.
 *
 * @param {string} mainItemId - UUID основного изделия на листе "База", который будет использоваться
 * для идентификации и обновления/добавления заглавной строки, а также в качестве родительского UUID для первого уровня вложенности.
 * @param {string} mainItemName - Название основного изделия.
 * @param {Object} firestoreSpecData - Полный объект данных спецификации изделия, полученный из Firestore,
 * включающий поле 'specification' (массив вложенных компонентов).
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss - Активная таблица Google Sheets.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} baseSheet - Объект листа "База".
 * @param {number} initialQuantity - Исходное количество основного изделия, с которого начинается развертывание.
 * Используется для расчета общего количества дочерних компонентов.
 */
function writeSpecToBaseSheetFromFirestore(mainItemId, mainItemName, firestoreSpecData, ss, baseSheet, initialQuantity) {
    const actualLastBaseColumn = baseSheet.getLastColumn();
    
    // 1. Начальная проверка входных данных
    if (!mainItemId || !mainItemName || !firestoreSpecData || !ss || !baseSheet || typeof initialQuantity === 'undefined') {
        Logger.log("Ошибка: Недостаточно данных для записи спецификации на лист 'База'.");
        ss.toast("Ошибка: Недостаточно данных для записи спецификации (отсутствует initialQuantity).", "Ошибка", 5);
        return;
    }

    // 2. Определение строки для вставки и очистка существующей заглавной строки
    const existingRow = findRowById(baseSheet, mainItemId); // findRowById возвращает либо число (номер строки), либо null
    let targetStartRowForInsertion; // Строка, С КОТОРОЙ начнется ЗАПИСЬ новой спецификации (включая заголовок)
    
    // Новая переменная для хранения примечания
    let existingNote = ''; 

    if (existingRow) { // Если заглавная строка для mainItemId уже существует
        // --- ИЗМЕНЕНИЕ 1: Сначала считываем примечание из существующей строки ---
        try {
            // Предполагаем, что COL_NOTE - это 1-индексированный номер столбца G
            const noteRange = baseSheet.getRange(existingRow, COL_NOTE); 
            existingNote = String(noteRange.getValue() || '').trim();
            Logger.log(`[writeSpecToBaseSheetFromFirestore] Найденное примечание из строки ${existingRow}: "${existingNote}"`);
        } catch (e) {
            Logger.log(`[writeSpecToBaseSheetFromFirestore] Ошибка при чтении примечания из строки ${existingRow}: ${e.message}`);
            existingNote = ''; // Сбросить примечание в случае ошибки
        }

        baseSheet.deleteRow(existingRow); // Удаляем старую заглавную строку
        Logger.log(`[writeSpecToBaseSheetFromFirestore] Удалена старая заглавная строка для "${mainItemName}" (UUID: ${mainItemId}) из строки ${existingRow}.`);
        targetStartRowForInsertion = existingRow; // Новая запись начнется с той же строки
    } else {
        targetStartRowForInsertion = baseSheet.getLastRow() + 1; // Добавляем в конец
        Logger.log(`[writeSpecToBaseSheetFromFirestore] Заглавная строка для "${mainItemName}" не найдена. Новая спецификация будет добавлена с строки ${targetStartRowForInsertion}.`);
    }

    // 3. Формирование массива строк для вставки (вызываем новую функцию!)
    // Передаем ей targetStartRowForInsertion как начальный номер строки для формул
    const allRowsToInsert = generateExpandedSpecificationRows(
        firestoreSpecData,
        initialQuantity,
        mainItemId,
        mainItemName,
        ss,
        actualLastBaseColumn,
        targetStartRowForInsertion,
        existingNote // <-- ИЗМЕНЕНИЕ 2: Передаем существующее примечание
    );

    // 4. Вызываем writeRowsToSheet, которая теперь сама вставляет строки и очищает фон
    writeRowsToSheet(allRowsToInsert, targetStartRowForInsertion, baseSheet, actualLastBaseColumn);

    Logger.log(`--- Завершено writeSpecToBaseSheetFromFirestore для ${mainItemName} ---`);
}

/**
 * Генерирует и возвращает массив строк, представляющих полную развернутую спецификацию
 * изделия из Firestore, включая расчеты дефицита, статусов и внутренние ссылки формул.
 * Эта функция не взаимодействует с листом Google Таблиц напрямую.
 *
 * @param {Object} firestoreSpecData - Полный объект данных спецификации изделия из Firestore.
 * @param {number} initialQuantity - Исходное количество основного изделия.
 * @param {string} mainItemId - UUID основного изделия (для первого уровня вложенности).
 * @param {string} mainItemName - Название основного изделия.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss - Активная таблица Google Sheets (для доступа к другим служебным функциям).
 * @param {number} totalColumnCount - Общее количество столбцов, используемых для записи (для createSpecRow).
 * @param {number} startRowForFormulas - Номер строки, с которой начнется запись этой спецификации в таблицу.
 * Используется для формирования внутренних ссылок в формулах.
 * @param {string} [headerNote=''] - Примечание для заглавной строки, если оно было считано из исходной строки.
 * @returns {Array<Array<any>>} Массив массивов, где каждый внутренний массив - это строка для записи в Google Sheet.
 */
function generateExpandedSpecificationRows(firestoreSpecData, initialQuantity, mainItemId, mainItemName, ss, totalColumnCount, startRowForFormulas, headerNote = '') { // <-- ИЗМЕНЕНИЕ 3: Добавляем параметр headerNote
    const allRowsToGenerate = [];
    let currentRowOffset = 0; // Смещение относительно startRowForFormulas

    // 1. Заглавная строка основного изделия //date, name, fullQty, parentName, futureIndex, parentUUIDForDestination, opCheckboxStatus, totalColumnCount, rowUUID, note = ''
    const reinsertedHeaderRow = createSpecRow(
        new Date(),
        mainItemName,
        initialQuantity,
        null, // Parent Name for main item is null
        startRowForFormulas + currentRowOffset, // Передаем актуальный номер строки для формул
        mainItemId,
        false, // Поменял на False по пожеланию Иры, чтобы не ставились лишние отметки
        totalColumnCount,
        mainItemId,
        headerNote // <-- ИЗМЕНЕНИЕ 4: Передаем примечание в createSpecRow
    );
    allRowsToGenerate.push(reinsertedHeaderRow);
    currentRowOffset++; // Увеличиваем смещение после добавления заглавной строки

    /**
     * Рекурсивно строит строки спецификации для вложенных компонентов.
     * Эта функция является внутренней вспомогательной функцией для generateExpandedSpecificationRows.
     *
     * @param {Array<Object>} components - Массив компонентов текущего уровня вложенности.
     * @param {number} currentParentItemTotalQty - Общее количество родительского изделия на корневом уровне.
     * @param {number} depth - Текущая глубина вложенности (для MAX_SPEC_DEPTH).
     * @param {string} currentParentItemUUID - UUID родительского элемента.
     * @param {string} currentParentItemName - Название родительского элемента.
     */
    function buildSpecRowsRecursive(components, currentParentItemTotalQty, depth, currentParentItemUUID, currentParentItemName) {
        if (depth > MAX_SPEC_DEPTH) { return; } // Предполагается, что MAX_SPEC_DEPTH определен глобально
        if (!Array.isArray(components) || components.length === 0) { return; }

        for (const component of components) {
            if (!component || typeof component !== 'object' || !component.componentName || typeof component.quantityPerParent === 'undefined' || !component.componentDocId) { continue; }

            const childName = component.componentName;
            const quantityPerParent = parseFloat(component.quantityPerParent);
            const childDocId = component.componentDocId;

            if (!childName || isNaN(quantityPerParent)) { continue; }

            const totalQuantityOfChildAtRootLevel = currentParentItemTotalQty * quantityPerParent;

            const deficit = calculateDeficit(childName, totalQuantityOfChildAtRootLevel, ss);
            const { isPurchased, opCheckboxStatus } = determinePurchasedAndCheckboxStatus(childName, deficit, ss);

            let nestedSpecComponents = null;

            try {
                let childDoc = getFirestoreDocument(PRODUCTS_COLLECTION, childDocId); // Предполагается, PRODUCTS_COLLECTION определен
                if (!childDoc || !childDoc.obj || !Array.isArray(childDoc.obj.specification) || childDoc.obj.specification.length === 0) {
                    const catalogMap = getCatalogMap();
                    processAndGetFirestoreComponentId(childName, depth + 1, catalogMap);
                    childDoc = getFirestoreDocument(PRODUCTS_COLLECTION, childDocId);
                }
                if (childDoc && childDoc.obj && Array.isArray(childDoc.obj.specification)) {
                    nestedSpecComponents = childDoc.obj.specification;
                }
            } catch (e) {
                Logger.log(`Ошибка при получении вложенной спецификации для ${childName} (ID: ${childDocId}): ${e.message}`);
            }

            const newRow = createSpecRow(
                null, // Date for child component is null
                childName,
                totalQuantityOfChildAtRootLevel,
                currentParentItemName,
                startRowForFormulas + currentRowOffset, // Передаем актуальный номер строки для формул
                currentParentItemUUID,
                opCheckboxStatus,
                totalColumnCount,
                null // UUID для компонента будет сгенерирован createSpecRow
            );

            allRowsToGenerate.push(newRow);
            currentRowOffset++; // Увеличиваем смещение для следующей строки

            if (Array.isArray(nestedSpecComponents) && nestedSpecComponents.length > 0 && deficit > 0 && !isPurchased) {
                const newlyGeneratedChildUUID = newRow[COL_UUID - 1]; // Получаем UUID только что сгенерированной строки
                buildSpecRowsRecursive(nestedSpecComponents, totalQuantityOfChildAtRootLevel, depth + 1, newlyGeneratedChildUUID, childName);
            }
        }
    }

    // Запускаем рекурсивный сбор данных
    buildSpecRowsRecursive(
        Array.isArray(firestoreSpecData.specification) ? firestoreSpecData.specification : [],
        initialQuantity,
        1, // Начальная глубина 1
        mainItemId,
        mainItemName
    );

    return allRowsToGenerate;
}


/**
 * Записывает массив данных на указанный лист Google Таблиц в заданном диапазоне.
 * Функция вставляет новые строки, если необходимо, или перезаписывает существующие.
 * Также очищает фоновое форматирование для вставленных строк.
 *
 * @param {Array<Array<any>>} rowsToInsert - Массив массивов, представляющий строки для записи.
 * @param {number} startRow - Номер строки (1-индексированный), с которой начинается запись.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - Объект листа, на который производится запись.
 * @param {number} columnCount - Количество столбцов, которые будут затронуты записью.
 */
function writeRowsToSheet(rowsToInsert, startRow, sheet, columnCount) {
    // Начальная проверка объекта sheet - оставляем, она важна для предотвращения критических ошибок.
    if (!sheet || typeof sheet.getName !== 'function' || typeof sheet.getLastRow !== 'function' || typeof sheet.getMaxRows !== 'function') {
        Logger.log(`[writeRowsToSheet] КРИТИЧЕСКАЯ ОШИБКА: Переданный объект 'sheet' не является действительным листом SpreadsheetApp.`);
        try {
            Logger.log(`[writeRowsToSheet] Некорректный объект sheet: ${JSON.stringify(sheet)}`);
        } catch (e) {
            Logger.log(`[writeRowsToSheet] Не удалось сериализовать некорректный объект sheet: ${e.message}`);
        }
        throw new Error('Критическая ошибка: В функцию writeRowsToSheet передан некорректный объект листа.');
    }

    const numRows = rowsToInsert.length;

    if (!Array.isArray(rowsToInsert) || numRows === 0) {
        Logger.log("[writeRowsToSheet] Нет строк для записи. Массив пуст.");
        return;
    }

    try {
        const lastPopulatedRow = sheet.getLastRow();
        const maxRowsInSheet = sheet.getMaxRows();
        const neededEndRow = startRow + numRows - 1;

        // Определяем, нужно ли вставлять строки
        // Случай 1: Вставляем в середину или в начало (если лист не пуст)
        if (startRow <= lastPopulatedRow && lastPopulatedRow > 0) {
            // Если startRow находится в пределах существующих строк листа,
            // вставляем numRows строк НАЧИНАЯ с startRow.
            sheet.insertRows(startRow, numRows);
            Logger.log(`[writeRowsToSheet] Вставлено ${numRows} пустых строк, начиная со строки ${startRow}, для новой записи (в середине).`);
        } 
        // Случай 2: Добавляем в конец листа или в пустой лист
        else if (startRow > lastPopulatedRow || (startRow === 1 && lastPopulatedRow === 0)) {
            // Если startRow находится за пределами последней заполненной строки
            // ИЛИ лист пуст (startRow = 1, lastPopulatedRow = 0)
            // И нам нужны дополнительные строки до того, как мы сможем записать
            const rowsToAddBelowMax = neededEndRow - maxRowsInSheet;
            if (rowsToAddBelowMax > 0) {
                // Если диапазон записи выходит за текущий максимальный размер листа,
                // добавляем столько строк, сколько необходимо.
                sheet.insertRowsAfter(maxRowsInSheet, rowsToAddBelowMax);
                Logger.log(`[writeRowsToSheet] Добавлено ${rowsToAddBelowMax} строк в конец листа для расширения до ${neededEndRow} (было ${maxRowsInSheet}).`);
            }
            // Здесь НЕТ вызова insertRows(startRow, numRows),
            // потому что мы уже либо расширили лист, либо startRow соответствует концу.
            // setValues будет работать непосредственно на существующем диапазоне.
            Logger.log(`[writeRowsToSheet] Лист расширен или данные добавляются в конец/пустой лист. Вставка строк insertRows не требуется.`);
        }
        
        // Очистка фонового форматирования для только что вставленного (или перезаписываемого) диапазона
        const rangeToFormat = sheet.getRange(startRow, 1, numRows, columnCount);
        rangeToFormat.setBackground(null); // Устанавливаем фоновый цвет в "нет заливки"
        Logger.log(`[writeRowsToSheet] Очищена фоновая заливка для ${numRows} строк, начиная со строки ${startRow}.`);

        // Запись данных в диапазон
        Logger.log(`[writeRowsToSheet] Запись ${numRows} строк в лист '${sheet.getName()}' в диапазон R${startRow}C1:R${startRow + numRows - 1}C${columnCount}`);
        rangeToFormat.setValues(rowsToInsert); // Записываем данные в очищенный диапазон
        Logger.log("[writeRowsToSheet] Запись завершена.");
    } catch (e) {
        Logger.log(`[writeRowsToSheet] ОШИБКА ПРИ ЗАПИСИ ДАННЫХ НА ЛИСТ:`);
        Logger.log(`[writeRowsToSheet] Сообщение: ${e.message}`);
        Logger.log(`[writeRowsToSheet] Имя ошибки: ${e.name}`);
        if (e.stack) {
            Logger.log(`[writeRowsToSheet] Стек вызовов: ${e.stack}`);
        }
        throw new Error(`Ошибка при записи данных на лист: ${e.message}`);
    }
}

/**
 * Рассчитывает дефицит для указанного компонента на основе данных об остатках на листе "Остатки".
 *
 * @param {string} itemName Название компонента, для которого нужно рассчитать дефицит.
 * @param {number} requiredQuantity Общее количество компонента, необходимое на корневом уровне.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss Активная таблица Google Sheets (для доступа к листам).
 * @returns {number} Рассчитанный дефицит. Если лист "Остатки" не найден или данные некорректны, возвращает requiredQuantity (т.е., весь требуемый объем в дефиците).
 */
function calculateDeficit(itemName, requiredQuantity, ss) {
    const stockSheet = ss.getSheetByName("Остатки");

    if (!stockSheet) {
        Logger.log(`Ошибка: Лист 'Остатки' не найден в calculateDeficit для ${itemName}. Предполагаем полный дефицит.`);
        return requiredQuantity;
    }

    const stockData = stockSheet.getDataRange().getValues();
    const stockMap = new Map();

    // Normalize item names when populating the stockMap
    for (let i = 1; i < stockData.length; i++) {
        if (stockData[i][0]) { // Ensure item name exists
            const normalizedItemName = stockData[i][0].toString().trim().toLowerCase(); // Normalize: trim and toLowerCase
            stockMap.set(normalizedItemName, parseFloat(stockData[i][1]) || 0);
        }
    }

    // Normalize the itemName being searched
    const normalizedSearchItemName = itemName.toString().trim().toLowerCase(); // Normalize the input itemName

    const stockQty = stockMap.has(normalizedSearchItemName) ? stockMap.get(normalizedSearchItemName) : 0;
    const deficit = requiredQuantity - stockQty;

    Logger.log(`[calculateDeficit] Изделие: ${itemName}, Требуется: ${requiredQuantity}, На складе: ${stockQty}, Дефицит: ${deficit}`);
    return deficit;
}


/**
 * Определяет статус "покупного" изделия и флаг для установки чекбокса
 * на основе его наличия в списке покупных материалов (с листа "Цена материалов")
 * и дефицита.
 *
 * @param {string} itemName Название изделия/компонента.
 * @param {number} deficit Рассчитанный дефицит (общее необходимое количество - количество на складе).
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss Активная таблица Google Sheets (необходима для доступа к листам).
 * @returns {{isPurchased: boolean, opCheckboxStatus: boolean}} Объект, содержащий статус "покупного" и статус чекбокса.
 */
function determinePurchasedAndCheckboxStatus(itemName, deficit, ss) {
    const purchasedSheet = ss.getSheetByName("Цена материалов");

    if (!purchasedSheet) {
        Logger.log("Ошибка: Лист 'Цена материалов' не найден в determinePurchasedAndCheckboxStatus. Предполагаем, что изделие не покупное.");
        return { isPurchased: false, opCheckboxStatus: deficit <= 0 };
    }

    const purchasedData = purchasedSheet.getDataRange().getValues();
    const purchasedSet = new Set();
    for (let i = 1; i < purchasedData.length; i++) {
        if (purchasedData[i][0]) { // Ensure item name exists
            purchasedSet.add(purchasedData[i][0].toString().trim());
        }
    }

    const isPurchased = purchasedSet.has(itemName);

    let opCheckboxStatus = false;
    // Чекбокс устанавливается, если изделие является покупным ИЛИ дефицита нет
    if (isPurchased || deficit <= 0) {
        opCheckboxStatus = true;
        Logger.log(`[determineStatus] Изделие ${itemName} является покупным (${isPurchased}) или имеется в наличии (дефицит: ${deficit}). Чекбокс будет установлен.`);
    } else {
        Logger.log(`[determineStatus] Изделие ${itemName} не является покупным и есть дефицит (${deficit}). Чекбокс НЕ будет установлен.`);
    }

    return {
        isPurchased: isPurchased,
        opCheckboxStatus: opCheckboxStatus
    };
}

/**
 * Вспомогательная функция для извлечения UUID из строки назначения в формате "РХХ = YYY (UUID)".
 *
 * @param {string} destinationString Строка из столбца COL_DESTINATION.
 * @returns {string|null} Извлеченный UUID или null, если не найден.
 */
function extractUuidFromDestination(destinationString) {
    if (!destinationString) return null;
    const match = destinationString.toString().match(/\(([0-9a-fA-F]{8}(?:-[0-9a-fA-F]{4}){3}-[0-9a-fA-F]{12})\)/);
    return match ? match[1] : null;
}


/**
 * Вспомогательная функция для извлечения информации о родительских компонентах
 * из строки столбца 'Destination' на листе "База".
 * Эта строка может содержать одно или несколько упоминаний о родительских изделиях
 * и их количестве, а также UUID родителя, если он доступен.
 *
 * @param {string} destinationString - Строка из столбца COL_DESTINATION (например, "Р01 = 2 (UUID), Р02 = 1 (UUID)").
 * @returns {Array<{parentName: string, parentQty: number, parentUuid: string|null}>} Массив объектов,
 * каждый из которых содержит имя родителя (parentName), количество (parentQty)
 * и UUID родителя (parentUuid), или null, если UUID отсутствует.
 */
function parseDestinationString(destinationString) {
    const parents = [];
    if (!destinationString) return parents;

    // Разбиваем строку на отдельные пары "Имя = Количество (UUID)"
    const parts = destinationString.split(',').map(part => part.trim());

    for (const part of parts) {
        // Regex для "Имя = Кол-во (UUID)"
        const matchWithUuid = part.match(/(.+) = (\d+)\s*\(([-\w]{25,})\)/);
        if (matchWithUuid && matchWithUuid.length === 4) {
            const parentName = matchWithUuid[1].trim();
            const parentQty = parseInt(matchWithUuid[2], 10);
            const parentUuid = matchWithUuid[3].trim();
            if (parentName && !isNaN(parentQty) && parentUuid) {
                parents.push({ parentName, parentQty, parentUuid });
            }
        } else {
            // Если формат старый "Имя = Кол-во" без UUID
            const oldMatch = part.match(/(.+) = (\d+)/);
            if (oldMatch && oldMatch.length === 3) {
                const parentName = oldMatch[1].trim();
                const parentQty = parseInt(oldMatch[2], 10);
                if (parentName && !isNaN(parentQty)) {
                    parents.push({ parentName, parentQty, parentUuid: null }); // UUID будет null
                }
            }
        }
    }
    return parents;
}