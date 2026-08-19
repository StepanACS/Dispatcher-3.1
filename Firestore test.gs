function getFirestoreDb() {
  return FirestoreApp.getFirestore(CLIENT_EMAIL, PRIVATE_KEY, PROJECT_ID);
}

/**
 * Тестирует работоспособность подключения к базе данных Firestore.
 * Использует методы, которые реально доступны в этой библиотеке (например, createDocument).
 */
function testFirestoreLib() {
  Logger.log('Начинаем тест подключения к Firestore...');

  try {
    Logger.log('Тип FirestoreApp: ' + typeof FirestoreApp);
    Logger.log('Объект FirestoreApp: ' + FirestoreApp);

    const db = getFirestoreDb(); // Получаем экземпляр Firestore

    const testCollectionName = 'test_collection_from_gas'; // Имя коллекции для теста
    const testDocId = 'test_doc_' + new Date().getTime(); // Уникальный ID документа
    const fullDocPath = testCollectionName + '/' + testDocId; // Полный путь к документу

    const testData = {
      "message": "Hello from Apps Script, updated!",
      "timestamp": new Date()
    };

    // --- Создание документа с указанным ID (как в документации) ---
    Logger.log(`Попытка создать тестовый документ "${fullDocPath}"...`);
    db.createDocument(fullDocPath, testData);
    Logger.log(`Успешно создан тестовый документ: ${fullDocPath}`);

    // --- Обновление документа (тест updateDocument) ---
    const updateData = {
      "updated_field": "This document was updated",
      "timestamp_updated": new Date()
    };
    Logger.log(`Попытка обновить тестовый документ "${fullDocPath}"...`);
    db.updateDocument(fullDocPath, updateData, true); // Обновляем с маской, чтобы сохранить старые поля
    Logger.log(`Успешно обновлен тестовый документ: ${fullDocPath}`);


    // --- Чтение документа (как в документации) ---
    Logger.log(`Попытка прочитать тестовый документ "${fullDocPath}"...`);
    const retrievedDoc = db.getDocument(fullDocPath); // Используем полный путь
    Logger.log('Содержимое тестового документа: ' + JSON.stringify(retrievedDoc.obj)); // Доступ к данным через .obj
    Logger.log('Время создания: ' + retrievedDoc.created);
    Logger.log('Время обновления: ' + retrievedDoc.updated);


    // --- Удаление документа (как в документации) ---
    Logger.log(`Удаляем тестовый документ "${fullDocPath}"...`);
    db.deleteDocument(fullDocPath);
    Logger.log(`Тестовый документ "${fullDocPath}" успешно удален.`);


    Logger.log('Тест подключения к Firestore успешно завершен!');

  } catch (e) {
    Logger.log('ОШИБКА при тестировании Firestore: ' + e.message);
    Logger.log('Подробности ошибки: ' + e.stack);
  }
}