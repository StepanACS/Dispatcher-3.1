/**
 * Ищет документ в коллекции Firestore по заданному полю и значению.
 *
 * @param {string} collectionName Имя коллекции, в которой производится поиск (например, "products").
 * @param {string} fieldName Имя поля, по которому производится поиск (например, "name", "uuid").
 * @param {any} searchValue Значение, которое нужно найти в поле (например, "Шпиндель XYZ").
 * @returns {string|null} ID найденного документа Firestore или null, если документ не найден.
 */
function searchFirestoreDocument(collectionName, fieldName, searchValue) {
  const firestore = getFirestoreDb();
    if (!collectionName || !fieldName || searchValue === undefined || searchValue === null) {
        Logger.log(`Ошибка: Некорректные параметры для searchFirestoreDocument. collectionName: ${collectionName}, fieldName: ${fieldName}, searchValue: ${searchValue}`);
        return null;
    }

    try {
        // Выполняем запрос к Firestore с синтаксисом из вашего образца
        const queryResult = firestore.query(collectionName)
            .Where(fieldName, "==", searchValue) // Возвращено к 'Where' с большой буквы
            .Limit(1)
            .Execute(); // Возвращено к 'Execute()'

        if (queryResult && queryResult.length > 0) {
            // Если документ найден, возвращаем его ID
            const doc = queryResult[0];
            const docId = doc.path.split('/').pop();
            Logger.log(`searchFirestoreDocument: Найден документ в коллекции "${collectionName}" по полю "${fieldName}" со значением "${searchValue}". ID: ${docId}`);
            return docId;
        } else {
            Logger.log(`searchFirestoreDocument: Документ в коллекции "${collectionName}" по полю "${fieldName}" со значением "${searchValue}" не найден.`);
            return null;
        }
    } catch (e) {
        Logger.log(`searchFirestoreDocument: Ошибка при поиске в коллекции "${collectionName}" по полю "${fieldName}" со значением "${searchValue}": ${e.message}`);
        return null;
    }
}


/**
 * Создает новый документ в коллекции Firestore с заданными данными.
 *
 * @param {string} collectionName Имя коллекции, в которую нужно записать документ (например, "products").
 * @param {Object} documentData Объект, содержащий данные для нового документа (например, { name: "Новый компонент", dateCreated: "...", specification: [] }).
 * @param {string|null} [docId=null] Опциональный ID для нового документа. Если null, Firestore сгенерирует ID автоматически.
 * @returns {string|null} ID созданного документа Firestore или null в случае ошибки.
 */
function writeFirestoreDocument(collectionName, documentData, docId = null) {
  const firestore = getFirestoreDb();
    if (!collectionName || !documentData || typeof documentData !== 'object') {
        Logger.log(`Ошибка: Некорректные параметры для writeFirestoreDocument. collectionName: ${collectionName}, documentData: ${JSON.stringify(documentData)}`);
        return null;
    }

    try {
        let createdDoc;
        if (docId) {
            // Если указан ID, создаем документ с этим ID
            createdDoc = firestore.createDocument(collectionName + '/' + docId, documentData);
            Logger.log(`writeFirestoreDocument: Создан документ в коллекции "${collectionName}" с указанным ID "${docId}".`);
        } else {
            // Если ID не указан, Firestore сгенерирует его автоматически
            createdDoc = firestore.createDocument(collectionName, documentData);
            Logger.log(`writeFirestoreDocument: Создан новый документ в коллекции "${collectionName}" с автоматически сгенерированным ID.`);
        }
        
        const newDocId = createdDoc.path.split('/').pop();
        return newDocId;
    } catch (e) {
        Logger.log(`writeFirestoreDocument: Ошибка при записи документа в коллекцию "${collectionName}": ${e.message}`);
        return null;
    }
}

/**
 * Обновляет существующий документ в Firestore.
 *
 * @param {string} collectionName Имя коллекции, в которой находится документ (например, "products").
 * @param {string} docId ID документа, который нужно обновить.
 * @param {Object} updateData Объект, содержащий поля и их новые значения для обновления.
 * @param {boolean} [merge=true] Флаг, указывающий, нужно ли объединять данные (true) или полностью перезаписывать документ (false). По умолчанию true.
 * @returns {boolean} true, если обновление прошло успешно, false в случае ошибки.
 */
function updateFirestoreDocument(collectionName, docId, updateData, merge = true) {
    if (!collectionName || !docId || !updateData || typeof updateData !== 'object') {
        Logger.log(`Ошибка: Некорректные параметры для updateFirestoreDocument. collectionName: ${collectionName}, docId: ${docId}, updateData: ${JSON.stringify(updateData)}`);
        return false;
    }

    try {
        const firestore = getFirestoreDb(); // Инициализация Firestore внутри функции
        const fullDocPath = `${collectionName}/${docId}`;
        
        firestore.updateDocument(fullDocPath, updateData, merge);
        Logger.log(`updateFirestoreDocument: Документ "${fullDocPath}" успешно обновлен. Merge: ${merge}`);
        return true;
    } catch (e) {
        Logger.log(`updateFirestoreDocument: Ошибка при обновлении документа "${collectionName}/${docId}": ${e.message}`);
        return false;
    }
}


/**
 * Получает данные документа из Firestore по его коллекции и ID.
 *
 * @param {string} collectionName Имя коллекции, в которой находится документ.
 * @param {string} docId ID документа, который нужно получить.
 * @returns {Object|null} Объект документа Firestore (например, {obj: {...}}) или null, если документ не найден или произошла ошибка.
 */
function getFirestoreDocument(collectionName, docId) {
    if (!collectionName || !docId) {
        Logger.log(`Ошибка: Некорректные параметры для getFirestoreDocument. collectionName: ${collectionName}, docId: ${docId}`);
        return null;
    }

    try {
        const firestore = getFirestoreDb(); // Инициализация Firestore
        const fullDocPath = `${collectionName}/${docId}`;
        const doc = firestore.getDocument(fullDocPath);

        if (doc && doc.obj) {
            Logger.log(`getFirestoreDocument: Документ "${fullDocPath}" успешно получен.`);
            return doc;
        } else {
            Logger.log(`getFirestoreDocument: Документ "${fullDocPath}" не найден или пуст.`);
            return null;
        }
    } catch (e) {
        Logger.log(`getFirestoreDocument: Ошибка при получении документа "${collectionName}/${docId}": ${e.message}`);
        return null;
    }
}










