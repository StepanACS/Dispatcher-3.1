let _firestoreDbCache = null;
let _catalogMapCache = new Map();
let _firestoreConfiguredCache = null;
let _oneLevelSpecCache = new Map();
let _firestoreExpansionCache = new Map();

function isFirestoreConfigured_() {
  if (_firestoreConfiguredCache !== null) return _firestoreConfiguredCache;
  const props = PropertiesService.getScriptProperties();
  _firestoreConfiguredCache = Boolean(
    props.getProperty('client_email') && props.getProperty('private_key') && props.getProperty('project_id')
  );
  return _firestoreConfiguredCache;
}

function getFirestoreDb() {
  if (_firestoreDbCache) return _firestoreDbCache;
  if (!isFirestoreConfigured_()) {
    throw new Error('Firestore не настроен: отсутствуют Script Properties client_email/private_key/project_id.');
  }
  const config = getFirestoreConfig_();
  _firestoreDbCache = FirestoreApp.getFirestore(config.clientEmail, config.privateKey, config.projectId);
  return _firestoreDbCache;
}

function handleDeployment(itemUUID, itemName, ss, baseSheet, rowFromDeploymentSheet, initialQuantity) {
  itemName = String(itemName || '').trim();
  itemUUID = String(itemUUID || '').trim();
  if (!itemName) { ss.toast('Ошибка: не определено имя изделия.', 'Ошибка', 5); return null; }
  if (!itemUUID) { ss.toast('Ошибка: не определён UUID изделия.', 'Ошибка', 5); return null; }

  setDispatcherRuntimeSpreadsheet_(ss);
  const initialRow = findBaseRowByUuid_(baseSheet, itemUUID);
  const deploymentStatus = isSpecificationDeployed(baseSheet, itemUUID);
  if (deploymentStatus.isDeployed) {
    if (initialRow) checkboxOk(baseSheet.getRange(initialRow, baseCol_(ss, 'stage_deployment')));
    ss.toast(`Спецификация «${itemName}» уже развернута.`, 'Информация', 5);
    return itemUUID;
  }

  const productSpec = getOrUploadProductSpecFromFirestore(itemName, ss);
  if (!productSpec) return null;

  try {
    writeSpecToBaseSheetFromFirestore(itemUUID, itemName, productSpec.data, ss, baseSheet, initialQuantity);
  } catch (error) {
    Logger.log(`handleDeployment: ${error.stack || error.message}`);
    ss.toast(`Ошибка при развертывании спецификации: ${error.message}`, 'Ошибка', 8);
    return null;
  }

  const finalRow = findBaseRowByUuid_(baseSheet, itemUUID);
  if (!finalRow) {
    Logger.log(`handleDeployment: после записи не найден UUID ${itemUUID}.`);
    return null;
  }

  checkboxOk(baseSheet.getRange(finalRow, baseCol_(ss, 'stage_deployment')));
  ss.toast(`Спецификация для «${itemName}» успешно обновлена.`, 'Успех', 5);
  return itemUUID;
}

/**
 * Variant B: спецификация считается развёрнутой, если у UUID есть хотя бы один
 * непосредственный ребёнок в База.parent_uuid. Compatibility mirror больше не нужен.
 */
function isSpecificationDeployed(baseSheet, itemUUID) {
  const target = String(itemUUID || '').trim();
  if (!target) return { isDeployed: false, targetRowOnBaseSheet: null };
  const ss = baseSheet.getParent();
  const parentCol = baseCol_(ss, 'parent_uuid');
  const lastRow = baseLastDataRow_(baseSheet);
  if (lastRow < 2) return { isDeployed: false, targetRowOnBaseSheet: null };

  const match = baseSheet.getRange(2, parentCol, lastRow - 1, 1)
    .createTextFinder(target).matchEntireCell(true).findNext();
  return {
    isDeployed: Boolean(match),
    targetRowOnBaseSheet: match ? match.getRow() : null
  };
}

function getOrUploadProductSpecFromFirestore(itemName, ss) {
  if (!isFirestoreConfigured_()) {
    Logger.log(`Firestore не настроен. «${itemName}» будет прочитано напрямую из Google Sheets.`);
    return getProductSpecDirectFromSheets_(itemName, ss);
  }
  try {
    const queryResults = getFirestoreDb().query(PRODUCTS_COLLECTION)
      .Where('name', '==', itemName).Limit(1).Execute();
    if (queryResults && queryResults.length > 0) {
      const doc = queryResults[0];
      return { docId: doc.path.split('/').pop(), data: doc.obj };
    }

    ss.toast(`Изделие «${itemName}» отсутствует в базе данных. Загружаю...`, 'Информация', 5);
    const uploadedDocId = uploadSpecToFirestoreFromSheet(itemName, ss);
    if (uploadedDocId) {
      const loadedDoc = getFirestoreDocument(PRODUCTS_COLLECTION, uploadedDocId);
      if (loadedDoc && loadedDoc.obj) return { docId: uploadedDocId, data: loadedDoc.obj };
    }
  } catch (error) {
    Logger.log(`Firestore недоступен для «${itemName}»: ${error.stack || error.message}`);
  }
  Logger.log(`Переход на прямое чтение спецификации «${itemName}» из Google Sheets.`);
  return getProductSpecDirectFromSheets_(itemName, ss);
}

function loadOneLevelSpecFromSheets_(itemName, catalogMap) {
  const normalizedName = String(itemName || '').trim();
  if (!normalizedName) return null;
  if (_oneLevelSpecCache.has(normalizedName)) return _oneLevelSpecCache.get(normalizedName);

  const link = catalogMap.get(normalizedName);
  if (!link) {
    const empty = { name: normalizedName, specification: [] };
    _oneLevelSpecCache.set(normalizedName, empty);
    return empty;
  }
  const sheetId = extractSheetId(link);
  if (!sheetId) return null;

  try {
    const sourceSs = SpreadsheetApp.openById(sheetId);
    const sourceSheet = sourceSs.getSheetByName(normalizedName);
    if (!sourceSheet) return null;
    const lastRow = sourceSheet.getLastRow();
    if (lastRow < 5) {
      const empty = { name: normalizedName, specification: [] };
      _oneLevelSpecCache.set(normalizedName, empty);
      return empty;
    }
    const rawData = sourceSheet.getRange(5, 2, lastRow - 4, 3).getValues();
    const totalIndex = rawData.findIndex(row => String(row[0] || '').trim().toLowerCase() === 'итого');
    if (totalIndex === -1) {
      const empty = { name: normalizedName, specification: [] };
      _oneLevelSpecCache.set(normalizedName, empty);
      return empty;
    }

    const specification = [];
    for (const row of rawData.slice(0, totalIndex)) {
      const componentName = String(row[0] || '').trim();
      const quantityPerParent = parseFloat(row[2]);
      if (!componentName || isNaN(quantityPerParent)) continue;
      specification.push({ componentName, quantityPerParent, componentDocId: null });
    }
    const result = { name: normalizedName, specification };
    _oneLevelSpecCache.set(normalizedName, result);
    return result;
  } catch (error) {
    Logger.log(`loadOneLevelSpecFromSheets_(«${normalizedName}»): ${error.stack || error.message}`);
    return null;
  }
}

function getProductSpecDirectFromSheets_(itemName, ss) {
  const catalogMap = getCatalogMap(ss);
  if (catalogMap.size === 0) { ss.toast('Не удалось прочитать «Каталог ссылок».', 'Ошибка', 5); return null; }
  const data = loadOneLevelSpecFromSheets_(itemName, catalogMap);
  if (!data) { ss.toast(`Не удалось прочитать спецификацию «${itemName}».`, 'Ошибка', 5); return null; }
  return { docId: null, data };
}

function findOrCreateComponentDocId(componentName, overwriteIfFound = false) {
  if (!componentName || !isFirestoreConfigured_()) return null;
  try {
    const docId = searchFirestoreDocument(PRODUCTS_COLLECTION, 'name', componentName);
    if (docId) {
      if (overwriteIfFound) {
        const ok = updateFirestoreDocument(PRODUCTS_COLLECTION, docId, {
          lastSpecUpdated: new Date().toISOString(), specification: []
        }, true);
        if (!ok) return null;
      }
      return docId;
    }
    const now = new Date().toISOString();
    return writeFirestoreDocument(PRODUCTS_COLLECTION, {
      name: componentName, dateCreated: now, lastSpecUpdated: null, specification: []
    });
  } catch (error) {
    Logger.log(`findOrCreateComponentDocId(${componentName}): ${error.stack || error.message}`);
    return null;
  }
}

function processAndGetFirestoreComponentId(currentCompName, depth, catalogMap, path) {
  if (depth > MAX_SPEC_DEPTH || !isFirestoreConfigured_()) return null;
  const normalizedName = String(currentCompName || '').trim();
  const currentPath = path instanceof Set ? new Set(path) : new Set();
  const cycleKey = normalizedName.toLowerCase();
  if (currentPath.has(cycleKey)) {
    throw new Error(`Циклическая спецификация: ${[...currentPath, cycleKey].join(' → ')}`);
  }
  currentPath.add(cycleKey);

  const componentDocId = findOrCreateComponentDocId(normalizedName);
  if (!componentDocId) return null;
  const source = loadOneLevelSpecFromSheets_(normalizedName, catalogMap);
  if (!source) return componentDocId;

  const spec = [];
  for (const component of source.specification) {
    const childId = processAndGetFirestoreComponentId(
      component.componentName, depth + 1, catalogMap, currentPath
    );
    if (childId) spec.push({
      componentName: component.componentName,
      quantityPerParent: component.quantityPerParent,
      componentDocId: childId
    });
  }
  updateFirestoreDocument(PRODUCTS_COLLECTION, componentDocId, {
    name: normalizedName, specification: spec, lastSpecUpdated: new Date().toISOString()
  }, true);
  _firestoreExpansionCache.delete(componentDocId);
  return componentDocId;
}

function getCatalogMap(ss) {
  const cacheKey = ss ? ss.getId() : MAIN_CATALOG_SHEET_ID;
  if (_catalogMapCache.has(cacheKey)) return _catalogMapCache.get(cacheKey);
  const catalogMap = new Map();
  try {
    let catalogSs = ss || SpreadsheetApp.openById(MAIN_CATALOG_SHEET_ID);
    let catalogSheet = catalogSs.getSheetByName('Каталог ссылок');
    if (!catalogSheet && (!ss || ss.getId() !== MAIN_CATALOG_SHEET_ID)) {
      catalogSs = SpreadsheetApp.openById(MAIN_CATALOG_SHEET_ID);
      catalogSheet = catalogSs.getSheetByName('Каталог ссылок');
    }
    if (catalogSheet) {
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
    }
  } catch (error) {
    Logger.log(`getCatalogMap: ${error.stack || error.message}`);
  }
  _catalogMapCache.set(cacheKey, catalogMap);
  return catalogMap;
}

function uploadSpecToFirestoreFromSheet(productName, ss) {
  if (!productName || !isFirestoreConfigured_()) return null;
  const catalogMap = getCatalogMap(ss);
  if (catalogMap.size === 0) return null;
  return processAndGetFirestoreComponentId(productName, 1, catalogMap, new Set());
}

function writeSpecToBaseSheetFromFirestore(mainItemId, mainItemName, firestoreSpecData, ss, baseSheet, initialQuantity) {
  if (!mainItemId || !mainItemName || !firestoreSpecData || !ss || !baseSheet || typeof initialQuantity === 'undefined') {
    ss.toast('Ошибка: недостаточно данных для записи спецификации.', 'Ошибка', 5);
    return;
  }

  const existingRow = findBaseRowByUuid_(baseSheet, mainItemId);
  let existingNote = '';
  if (existingRow) {
    existingNote = String(readBaseFields_(baseSheet, existingRow, ['note']).note || '').trim();
    clearBaseRawRow_(baseSheet, existingRow);
  }

  const targetStartRow = Math.max(baseLastDataRow_(baseSheet) + 1, 3);
  const rows = generateExpandedSpecificationRows(
    firestoreSpecData, initialQuantity, mainItemId, mainItemName, ss,
    baseSheet.getLastColumn(), targetStartRow, existingNote
  );
  writeRowsToSheet(rows, targetStartRow, baseSheet, baseSheet.getLastColumn());
}

function getFirestoreExpansionDocument_(docId) {
  const key = String(docId || '').trim();
  if (!key) return null;
  if (_firestoreExpansionCache.has(key)) return _firestoreExpansionCache.get(key);
  const doc = getFirestoreDocument(PRODUCTS_COLLECTION, key);
  _firestoreExpansionCache.set(key, doc || null);
  return doc || null;
}

/**
 * Разворачивает Variant B с двумя важными инвариантами:
 * 1) складской остаток расходуется между повторными применениями компонента;
 * 2) вложенный BOM строится только на фактический дефицит, а не на всю потребность.
 */
function generateExpandedSpecificationRows(firestoreSpecData, initialQuantity, mainItemId, mainItemName, ss, totalColumnCount, startRowForFormulas, headerNote = '') {
  const rows = [];
  const remainingStock = new Map(buildStockMap_(ss));
  const purchasedSet = buildPurchasedSet_(ss);
  const catalogMap = getCatalogMap(ss);

  rows.push(createSpecRow(
    new Date(), mainItemName, initialQuantity, null, startRowForFormulas,
    mainItemId, false, totalColumnCount, mainItemId, headerNote
  ));

  function getNestedSpec_(component, childName, depth, shouldExpand) {
    if (!shouldExpand || depth >= MAX_SPEC_DEPTH) return [];
    if (Array.isArray(component.specification)) return component.specification;
    const childDocId = component.componentDocId;
    if (childDocId && isFirestoreConfigured_()) {
      try {
        let childDoc = getFirestoreExpansionDocument_(childDocId);
        if (!childDoc || !childDoc.obj || !Array.isArray(childDoc.obj.specification) || childDoc.obj.specification.length === 0) {
          processAndGetFirestoreComponentId(childName, depth + 1, catalogMap, new Set());
          _firestoreExpansionCache.delete(childDocId);
          childDoc = getFirestoreExpansionDocument_(childDocId);
        }
        if (childDoc && childDoc.obj && Array.isArray(childDoc.obj.specification)) return childDoc.obj.specification;
      } catch (error) {
        Logger.log(`Не удалось получить вложенную спецификацию «${childName}» из Firestore: ${error.message}`);
      }
    }
    const direct = loadOneLevelSpecFromSheets_(childName, catalogMap);
    return direct && Array.isArray(direct.specification) ? direct.specification : [];
  }

  function consumeStock_(itemName, requiredQuantity) {
    const key = String(itemName || '').trim().toLowerCase();
    const required = Math.max(Number(requiredQuantity) || 0, 0);
    const available = Math.max(Number(remainingStock.get(key)) || 0, 0);
    const used = Math.min(required, available);
    remainingStock.set(key, available - used);
    return required - used;
  }

  function buildSpecRowsRecursive(components, parentQtyToMake, depth, parentUuid, parentName, path) {
    if (depth > MAX_SPEC_DEPTH || !Array.isArray(components) || !components.length) return;
    for (const component of components) {
      if (!component || typeof component !== 'object') continue;
      const childName = String(component.componentName || '').trim();
      const quantityPerParent = parseFloat(component.quantityPerParent);
      if (!childName || isNaN(quantityPerParent)) continue;

      const totalQuantity = parentQtyToMake * quantityPerParent;
      const shortage = consumeStock_(childName, totalQuantity);
      const status = determinePurchasedAndCheckboxStatus(childName, shortage, ss, purchasedSet);
      const shouldExpand = shortage > 0 && !status.isPurchased;
      const childKey = childName.toLowerCase();
      if (shouldExpand && path.has(childKey)) {
        throw new Error(`Циклическая спецификация: ${[...path, childKey].join(' → ')}`);
      }

      const nestedSpec = getNestedSpec_(component, childName, depth, shouldExpand);
      const newRow = createSpecRow(
        null, childName, totalQuantity, parentName, 0, parentUuid,
        status.opCheckboxStatus, totalColumnCount, null
      );
      rows.push(newRow);

      if (nestedSpec.length > 0 && shouldExpand) {
        const childPath = new Set(path);
        childPath.add(childKey);
        buildSpecRowsRecursive(nestedSpec, shortage, depth + 1, newRow.uuid, childName, childPath);
      }
    }
  }

  const rootPath = new Set([String(mainItemName || '').trim().toLowerCase()]);
  buildSpecRowsRecursive(
    Array.isArray(firestoreSpecData.specification) ? firestoreSpecData.specification : [],
    Number(initialQuantity) || 0, 1, mainItemId, mainItemName, rootPath
  );
  return rows;
}

function writeRowsToSheet(rowsToInsert, startRow, sheet, columnCount) {
  if (!sheet || !Array.isArray(rowsToInsert) || !rowsToInsert.length) return;
  writeBaseRows_(sheet, startRow, rowsToInsert);
  applyBaseStageValidation_(sheet.getParent(), sheet, startRow, rowsToInsert.length);
}

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

function buildPurchasedSet_(ss) {
  const result = new Set();
  const sheet = ss.getSheetByName('Цена материалов');
  if (!sheet) return result;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return result;
  const data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (const row of data) {
    const name = String(row[0] || '').trim().toLowerCase();
    if (name) result.add(name);
  }
  return result;
}

function calculateDeficit(itemName, requiredQuantity, ss, stockMap) {
  const map = stockMap || buildStockMap_(ss);
  const key = String(itemName || '').trim().toLowerCase();
  const stockQty = map.has(key) ? Number(map.get(key)) || 0 : 0;
  return Math.max((Number(requiredQuantity) || 0) - stockQty, 0);
}

function determinePurchasedAndCheckboxStatus(itemName, deficit, ss, purchasedSet) {
  const set = purchasedSet || buildPurchasedSet_(ss);
  const isPurchased = set.has(String(itemName || '').trim().toLowerCase());
  return { isPurchased, opCheckboxStatus: isPurchased || deficit <= 0 };
}
