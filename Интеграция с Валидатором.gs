/**
 * Обновляет поле 'specification' и 'lastSpecUpdated' для существующего компонента в Firestore.
 * Переформатирует входные данные из validator в структуру, которую ожидает диспетчер.
 *
 * @param {string} componentName Название компонента/изделия.
 * @param {Array<Object>} specificationData Массив строк спецификации из validator.
 * @returns {string|null} ID документа компонента Firestore, или null в случае ошибки.
 */
function updateProductSpecification(componentName, specificationData) {
    if (!componentName || !Array.isArray(specificationData)) {
        Logger.log(`[dispatcher.updateProductSpecification] Недостаточно данных для "${componentName}"`);
        return null;
    }

    const now = new Date().toISOString();

    const reformattedSpec = [];

    for (const row of specificationData) {
        const name = row.name?.toString().trim();
        const qty = parseFloat(row.quantity);
        if (!name || isNaN(qty)) {
            Logger.log(`[dispatcher.updateProductSpecification] Пропущена строка без имени или количества: ${JSON.stringify(row)}`);
            continue;
        }

        const childId = findOrCreateComponentDocId(name); // НЕ рекурсивно!
        if (!childId) {
            Logger.log(`[dispatcher.updateProductSpecification] Не удалось получить docId для компонента "${name}"`);
            continue;
        }

        reformattedSpec.push({
            componentName: name,
            quantityPerParent: qty,
            componentDocId: childId
        });
    }

    try {
        let docId = searchFirestoreDocument(PRODUCTS_COLLECTION, "name", componentName);

        if (docId) {
            Logger.log(`[dispatcher.updateProductSpecification] Обновляем спецификацию для "${componentName}" (ID: ${docId})`);
            const updateSuccess = updateFirestoreDocument(PRODUCTS_COLLECTION, docId, {
                lastSpecUpdated: now,
                specification: reformattedSpec
            }, true);
            return updateSuccess ? docId : null;
        } else {
            Logger.log(`[dispatcher.updateProductSpecification] Создаем новый документ для "${componentName}"`);
            return writeFirestoreDocument(PRODUCTS_COLLECTION, {
                name: componentName,
                dateCreated: now,
                lastSpecUpdated: now,
                specification: reformattedSpec
            });
        }
    } catch (e) {
        Logger.log(`[dispatcher.updateProductSpecification] Ошибка при обновлении "${componentName}": ${e.message}`);
        return null;
    }
}
