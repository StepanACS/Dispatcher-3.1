let _firestoreDbCache = null;
let _catalogMapCache = null;

/**
 * Возвращает Firestore-клиент. Конфигурация и клиент создаются лениво и только
 * один раз за execution.
 */
function getFirestoreDb() {
  if (_firestoreDbCache) return _firestoreDbCache;
  const config = getFirestoreConfig_();
  _firestoreDbCache = FirestoreApp.getFirestore(
    config.clientEmail,
    config.privateKey,
    config.projectId
  );
  return _firestoreDbCache;
}

/**
 * Разворачивает спецификацию изделия на лист «База» либо подтверждает уже
 * существующее развёртывание.
 */
function handleDeployment(itemUUID, itemName, ss, baseSheet, rowFromDeploymentSheet, initialQuantity) {
  itemName = String(itemName || '').trim();
  itemUUID = String(itemUUID || '').trim();

  if (!itemName) {
    ss.toast('Ошибка: не определено имя изделия.', 'Ошибка', 5);
    return null;
  }
  if (!itemUUID) {
    ss.toast('Ошибка: не определён UUID изделия.', 'Ошибка', 5);
    return null;
  }

  const initialRow = findRowById(baseSheet, itemUUID);
  const deploymentStatus = isSpecificationDeployed(baseSheet, itemUUID, initialRow || 0);

  if (deploymentStatus.isDeployed) {
    if (initialRow) checkboxOk(baseSheet.getRange(initialRow, COL_DEPLOYMENT));
    ss.toast(`Спецификация «${itemName}» уже развернута.`, 'Информация', 5);
    return itemUUID;
  }

  const productSpec = getOrUploadProductSpecFromFirestore(itemName, ss);
  if (!productSpec) return null;

  try {
    writeSpecToBaseSheetFromFirestore(
      itemUUID,
      itemName,
      productSpec.data,
      ss,
      baseSheet,
      initialQuantity
    );
  } catch (error) {
    Logger.log(`handleDeployment: ${error.stack || error.message}`);
    ss.toast(`Ошибка при развертывании спецификации: ${error.message}`, 'Ошибка', 5);
    return null;
  }

  const finalRow = findRowById(baseSheet, itemUUID);
  if (!finalRow) {
    Logger.log(`handleDeployment: после записи не найден UUID ${itemUUID}.`);
    return null;
  }

  checkboxOk(baseSheet.getRange(finalRow, COL_DEPLOYMENT));
  checkAndMergeDuplicates(baseSheet, baseSheet.getLastColumn());
  ss.toast(`Спецификация для «${itemName}» успешно обновлена.`, 'Успех', 5);
  return itemUUID;
}

/**
 * Проверяет наличие того же UUID или дочерней ссылки на UUID за пределами
 * исходной строки. Для проверки читаются только D:F, а не весь диапазон «Базы».
 */
function isSpecificationDeployed(baseSheet, itemUUID, excludeRow) {
  const lastRow = baseSheet.getLastRow();
  if (lastRow < 2) return { isDeployed: false, targetRowOnBaseSheet: null };

  const width = COL_DESTINATION - COL_UUID + 1;
  const data = baseSheet.getRange(2, COL_UUID, lastRow - 1, width).getValues();
  const targetUuid = String(itemUUID).trim();
  const uuidOffset = 0;
  const destinationOffset = COL_DESTINATION - COL_UUID;

  for (let i = 0; i < data.length; i++) {
    const currentRow = i + 2;
    if (currentRow === excludeRow) continue;

    const rowUuid = String(data[i][uuidOffset] || '').trim();
    if (rowUuid === targetUuid) {
      return { isDeployed: true, targetRowOnBaseSheet: currentRow };
    }

    const destination = String(data[i][destinationOffset] || '');
    if (extractUuidFromDestination(destination) === targetUuid) {
      return { isDeployed: true, targetRowOnBaseSheet: null };
    }
  }

  return { isDeployed: false, targetRowOnBaseSheet: null };
}

/** Получает спецификацию из Firestore, а при отсутствии загружает её из Sheets. */
function getOrUploadProductSpecFromFirestore(itemName, ss) {
  try {
    const queryResults = getFirestoreDb()
      .query(PRODUCTS_COLLECTION)
      .Where('name', '==', itemName)
      .Limit(1)
      .Execute();

    if (queryResults && queryResults.length > 0) {
      const doc = queryResults[0];
      return { docId: doc.path.split('/').pop(), data: doc.obj };
    }

    ss.toast(`Изделие «${itemName}» отсутствует в базе данных. Загружаю...`, 'Информация', 5);
    const uploadedDocId = uploadSpecToFirestoreFromSheet(itemName);
    if (!uploadedDocId) return null;

    const loadedDoc = getFirestoreDocument(PRODUCTS_COLLECTION, uploadedDocId);
    if (loadedDoc && loadedDoc.obj) {
      ss.toast(`Изделие «${itemName}» загружено в базу данных.`, 'Информация', 5);
      return { docId: uploadedDocId, data: loadedDoc.obj };
    }

    return null;
  } catch (error) {
    Logger.log(`getOrUploadProductSpecFromFirestore: ${error.stack || error.message}`);
    ss.toast(`Ошибка при работе с Firestore для «${itemName}».`, 'Ошибка', 5);
    return null;
  }
}

/** Ищет компонент в Firestore и при отсутствии создаёт его. */
function findOrCreateComponentDocId(componentName, overwriteIfFound = false) {
  if (!componentName) return null;

  try {
    const docId = searchFirestoreDocument(PRODUCTS_COLLECTION, 'name', componentName);
    if (docId) {
      if (overwriteIfFound) {
        const ok = updateFirestoreDocument(PRODUCTS_COLLECTION, docId, {
          lastSpecUpdated: new Date().toISOString(),
          specification: []
        }, true);
        if (!ok) return null;
      }
      return docId;
    }

    const now = new Date().toISOString();
    return writeFirestoreDocument(PRODUCTS_COLLECTION, {
      name: componentName,
      dateCreated: now,
      lastSpecUpdated: null,
      specification: []
    });
  } catch (error) {
    Logger.log(`findOrCreateComponentDocId(${componentName}): ${error.stack || error.message}`);
    return null;
  }
}

/**
 * Рекурсивно читает спецификации из связанных Google Sheets и синхронизирует
 * иерархию компонентов с Firestore. catalogMap передаётся по рекурсии и не
 * перечитывается для каждого компонента.
 */
function processAndGetFirestoreComponentId(currentCompName, depth, catalogMap) {
  if (depth > MAX_SPEC_DEPTH) return null;

  const componentDocId = findOrCreateComponentDocId(currentCompName);
  if (!componentDocId) return null;

  const componentSpecLink = catalogMap.get(currentCompName);
  if (!componentSpecLink) {
    updateFirestoreDocument(PRODUCTS_COLLECTION, componentDocId, {
      name: currentCompName,
      lastSpecUpdated: new Date().toISOString()
    }, true);
    return componentDocId;
  }

  const sheetId = extractSheetId(componentSpecLink);
  if (!sheetId) return componentDocId;

  let sheet;
  try {
    const sheetDoc = SpreadsheetApp.openById(sheetId);
    sheet = sheetDoc.getSheetByName(currentCompName);
  } catch (error) {
    Logger.log(`Не удалось открыть спецификацию «${currentCompName}»: ${error.message}`);
    return componentDocId;
  }

  if (!sheet) return componentDocId;
  const lastRow = sheet.getLastRow();
  if (lastRow < 5) return componentDocId;

  const rawData = sheet.getRange(5, 2, lastRow - 4, 3).getValues();
  const totalIndex = rawData.findIndex(row =>
    String(row[0] || '').trim().toLowerCase() === 'итого'
  );
  if (totalIndex === -1) return componentDocId;

  const spec = [];
  for (const row of rawData.slice(0, totalIndex)) {
    const name = String(row[0] || '').trim();
    const qty = parseFloat(row[2]);
    if (!name || isNaN(qty)) continue;

    const childId = processAndGetFirestoreComponentId(name, depth + 1, catalogMap);
    if (childId) {
      spec.push({
        componentName: name,
        quantityPerParent: qty,
        componentDocId: childId
      });
    }
  }

  updateFirestoreDocument(PRODUCTS_COLLECTION, componentDocId, {
    name: currentCompName,
    specification: spec,
    lastSpecUpdated: new Date().toISOString()
  }, true);

  return componentDocId;
}

/**
 * Возвращает карту «изделие → ссылка на спецификацию». Карта читается один раз
 * за execution и затем повторно используется всем деревом развертывания.
 */
function getCatalogMap() {
  if (_catalogMapCache) return _catalogMapCache;

  const catalogMap = new Map();
  try {
    const ssCatalog = SpreadsheetApp.openById(MAIN_CATALOG_SHEET_ID);
    const catalogSheet = ssCatalog.getSheetByName('Каталог ссылок');
    if (!catalogSheet) return catalogMap;

    const lastRow = catalogSheet.getLastRow();
    if (lastRow >= 2) {
      const width = Math.max(CATALOG_COL_ITEMNAME, CATALOG_COL_LINK);
      const data = catalogSheet.getRange(2, 1, lastRow - 1, width).getValues();
      for (const row of data) {
        const name = String(row[CATALOG_COL_ITEMNAME - 1] || '').trim();
        const link = String(row[CATALOG_COL_LINK - 1] || '').trim();
        if (name && link) catalogMap.set(name, link);
      }
    }
  } catch (error) {
    Logger.log(`getCatalogMap: ${error.message}`);
  }

  _catalogMapCache = catalogMap;
  return _catalogMapCache;
}

/** Загружает полную иерархию спецификации изделия в Firestore. */
function uploadSpecToFirestoreFromSheet(productName) {
  if (!productName) return null;
  const catalogMap = getCatalogMap();
  if (catalogMap.size === 0) return null;
  return processAndGetFirestoreComponentId(productName, 1, catalogMap);
}

/** Записывает развёрнутую спецификацию в «База». */
function writeSpecToBaseSheetFromFirestore(mainItemId, mainItemName, firestoreSpecData, ss, baseSheet, initialQuantity) {
  if (!mainItemId || !mainItemName || !firestoreSpecData || !ss || !baseSheet ||
      typeof initialQuantity === 'undefined') {
    ss.toast('Ошибка: недостаточно данных для записи спецификации.', 'Ошибка', 5);
    return;
  }

  const columnCount = baseSheet.getLastColumn();
  const existingRow = findRowById(baseSheet, mainItemId);
  let targetStartRow;
  let existingNote = '';

  if (existingRow) {
    existingNote = String(baseSheet.getRange(existingRow, COL_NOTE).getValue() || '').trim();
    baseSheet.deleteRow(existingRow);
    targetStartRow = existingRow;
  } else {
    targetStartRow = baseSheet.getLastRow() + 1;
  }

  const rows = generateExpandedSpecificationRows(
    firestoreSpecData,
    initialQuantity,
    mainItemId,
    mainItemName,
    ss,
    columnCount,
    targetStartRow,
    existingNote
  );

  writeRowsToSheet(rows, targetStartRow, baseSheet, columnCount);
}

/**
 * Строит массив строк спецификации полностью в памяти. Справочники «Остатки»,
 * «Цена материалов» и «Каталог ссылок» загружаются один раз до рекурсии.
 */
function generateExpandedSpecificationRows(firestoreSpecData, initialQuantity, mainItemId, mainItemName, ss, totalColumnCount, startRowForFormulas, headerNote = '') {
  const rows = [];
  let currentRowOffset = 0;

  const stockMap = buildStockMap_(ss);
  const purchasedSet = buildPurchasedSet_(ss);
  const catalogMap = getCatalogMap();

  const headerRow = createSpecRow(
    new Date(),
    mainItemName,
    initialQuantity,
    null,
    startRowForFormulas,
    mainItemId,
    false,
    totalColumnCount,
    mainItemId,
    headerNote
  );
  rows.push(headerRow);
  currentRowOffset++;

  function buildSpecRowsRecursive(components, parentTotalQty, depth, parentUuid, parentName) {
    if (depth > MAX_SPEC_DEPTH || !Array.isArray(components) || components.length === 0) return;

    for (const component of components) {
      if (!component || typeof component !== 'object') continue;

      const childName = String(component.componentName || '').trim();
      const quantityPerParent = parseFloat(component.quantityPerParent);
      const childDocId = component.componentDocId;
      if (!childName || isNaN(quantityPerParent) || !childDocId) continue;

      const totalQuantity = parentTotalQty * quantityPerParent;
      const deficit = calculateDeficit(childName, totalQuantity, ss, stockMap);
      const status = determinePurchasedAndCheckboxStatus(childName, deficit, ss, purchasedSet);

      let childDoc = getFirestoreDocument(PRODUCTS_COLLECTION, childDocId);
      if (!childDoc || !childDoc.obj || !Array.isArray(childDoc.obj.specification) ||
          childDoc.obj.specification.length === 0) {
        processAndGetFirestoreComponentId(childName, depth + 1, catalogMap);
        childDoc = getFirestoreDocument(PRODUCTS_COLLECTION, childDocId);
      }

      const nestedSpec = childDoc && childDoc.obj && Array.isArray(childDoc.obj.specification)
        ? childDoc.obj.specification
        : [];

      const newRow = createSpecRow(
        null,
        childName,
        totalQuantity,
        parentName,
        startRowForFormulas + currentRowOffset,
        parentUuid,
        status.opCheckboxStatus,
        totalColumnCount,
        null
      );

      rows.push(newRow);
      currentRowOffset++;

      if (nestedSpec.length > 0 && deficit > 0 && !status.isPurchased) {
        buildSpecRowsRecursive(
          nestedSpec,
          totalQuantity,
          depth + 1,
          newRow[COL_UUID - 1],
          childName
        );
      }
    }
  }

  buildSpecRowsRecursive(
    Array.isArray(firestoreSpecData.specification) ? firestoreSpecData.specification : [],
    initialQuantity,
    1,
    mainItemId,
    mainItemName
  );

  return rows;
}

/** Записывает подготовленные строки одним setValues(). */
function writeRowsToSheet(rowsToInsert, startRow, sheet, columnCount) {
  if (!sheet || !Array.isArray(rowsToInsert) || rowsToInsert.length === 0) return;

  const numRows = rowsToInsert.length;
  const lastPopulatedRow = sheet.getLastRow();
  const maxRows = sheet.getMaxRows();
  const neededEndRow = startRow + numRows - 1;

  if (startRow <= lastPopulatedRow && lastPopulatedRow > 0) {
    sheet.insertRows(startRow, numRows);
  } else if (neededEndRow > maxRows) {
    sheet.insertRowsAfter(maxRows, neededEndRow - maxRows);
  }

  const target = sheet.getRange(startRow, 1, numRows, columnCount);
  target.setBackground(null);
  target.setValues(rowsToInsert);
}

/** Загружает остатки один раз в Map для всего развёртывания. */
function buildStockMap_(ss) {
  const result = new Map();
  const sheet = ss.getSheetByName('Остатки');
  if (!sheet) return result;

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return result;

  const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  for (const row of data) {
    const name = String(row[0] || '').trim().toLowerCase();
    if (name) result.set(name, parseFloat(row[1]) || 0);
  }
  return result;
}

/** Загружает список покупных материалов один раз в Set для всего развёртывания. */
function buildPurchasedSet_(ss) {
  const result = new Set();
  const sheet = ss.getSheetByName('Цена материалов');
  if (!sheet) return result;

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return result;

  const data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (const row of data) {
    const name = String(row[0] || '').trim();
    if (name) result.add(name);
  }
  return result;
}

/** Рассчитывает дефицит; готовый stockMap можно передать из рекурсивного процесса. */
function calculateDeficit(itemName, requiredQuantity, ss, stockMap) {
  const map = stockMap || buildStockMap_(ss);
  const key = String(itemName || '').trim().toLowerCase();
  const stockQty = map.has(key) ? map.get(key) : 0;
  return requiredQuantity - stockQty;
}

/** Определяет покупной статус; готовый Set можно передать из рекурсивного процесса. */
function determinePurchasedAndCheckboxStatus(itemName, deficit, ss, purchasedSet) {
  const set = purchasedSet || buildPurchasedSet_(ss);
  const isPurchased = set.has(String(itemName || '').trim());
  return {
    isPurchased,
    opCheckboxStatus: isPurchased || deficit <= 0
  };
}

function extractUuidFromDestination(destinationString) {
  if (!destinationString) return null;
  const match = String(destinationString).match(/\(([0-9a-fA-F]{8}(?:-[0-9a-fA-F]{4}){3}-[0-9a-fA-F]{12})\)/);
  return match ? match[1] : null;
}

function parseDestinationString(destinationString) {
  const parents = [];
  if (!destinationString) return parents;

  const parts = String(destinationString).split(',').map(part => part.trim());
  for (const part of parts) {
    const matchWithUuid = part.match(/(.+) = (\d+)\s*\(([-\w]{25,})\)/);
    if (matchWithUuid && matchWithUuid.length === 4) {
      const parentName = matchWithUuid[1].trim();
      const parentQty = parseInt(matchWithUuid[2], 10);
      const parentUuid = matchWithUuid[3].trim();
      if (parentName && !isNaN(parentQty) && parentUuid) {
        parents.push({ parentName, parentQty, parentUuid });
      }
      continue;
    }

    const oldMatch = part.match(/(.+) = (\d+)/);
    if (oldMatch && oldMatch.length === 3) {
      const parentName = oldMatch[1].trim();
      const parentQty = parseInt(oldMatch[2], 10);
      if (parentName && !isNaN(parentQty)) {
        parents.push({ parentName, parentQty, parentUuid: null });
      }
    }
  }
  return parents;
}
