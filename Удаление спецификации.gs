/** Сворачивает спецификацию изделия, указанного в текущей строке. */
function handleUndeployment(ss, sheet, row) {
  const itemId = String(sheet.getRange(row, COL_UUID).getValue() || '').trim();
  if (!itemId) return;
  removeSpecification(itemId, ss);
}

/**
 * Удаляет развёрнутую ветку спецификации по UUID корневого изделия.
 *
 * Алгоритм сначала строит граф связей полностью в памяти, затем выполняет
 * структурные удаления блоками и обновляет только те строки, у которых остались
 * другие родители. Для анализа читаются только D:H вместо всей 40-колоночной Базы.
 */
function removeSpecification(initialParentUuid, ss) {
  const baseSheet = ss.getSheetByName('База');
  if (!baseSheet) {
    ss.toast('Ошибка: лист «База» не найден.', 'Ошибка', 5);
    return;
  }

  const rootUuid = String(initialParentUuid || '').trim();
  if (!rootUuid) return;

  const lastRow = baseSheet.getLastRow();
  if (lastRow < 2) return;

  // D:H: UUID, имя, назначение, примечание, количество.
  const data = baseSheet
    .getRange(2, COL_UUID, lastRow - 1, COL_QUANTITY - COL_UUID + 1)
    .getValues();

  const UUID_OFFSET = 0;
  const NAME_OFFSET = COL_ITEMNAME - COL_UUID;
  const DESTINATION_OFFSET = COL_DESTINATION - COL_UUID;
  const QUANTITY_OFFSET = COL_QUANTITY - COL_UUID;

  // UUID -> исходная информация о строке.
  const rowsByUuid = new Map();
  // ChildUUID -> Set<ParentUUID>.
  const currentRelationships = new Map();
  // ParentUUID -> Set<ChildUUID>.
  const parentToChildren = new Map();

  data.forEach((rowData, index) => {
    const uuid = String(rowData[UUID_OFFSET] || '').trim();
    if (!uuid) return;

    const destination = String(rowData[DESTINATION_OFFSET] || '').trim();
    const parsedDestinations = parseDestinationCell(destination);
    const entry = {
      rowNum: index + 2,
      name: String(rowData[NAME_OFFSET] || '').trim(),
      destination,
      quantity: rowData[QUANTITY_OFFSET],
      parsedDestinations
    };
    rowsByUuid.set(uuid, entry);

    for (const destinationInfo of parsedDestinations) {
      const parentUuid = destinationInfo.uuid;

      if (!currentRelationships.has(uuid)) {
        currentRelationships.set(uuid, new Set());
      }
      currentRelationships.get(uuid).add(parentUuid);

      if (!parentToChildren.has(parentUuid)) {
        parentToChildren.set(parentUuid, new Set());
      }
      parentToChildren.get(parentUuid).add(uuid);
    }
  });

  if (!rowsByUuid.has(rootUuid)) {
    Logger.log(`removeSpecification: UUID ${rootUuid} не найден в Базе.`);
    return;
  }

  const processingQueue = [rootUuid];
  const queued = new Set([rootUuid]);
  const uuidsToDelete = new Set([rootUuid]);
  const pendingUpdates = new Map();

  while (processingQueue.length > 0) {
    const currentUuid = processingQueue.shift();
    const children = parentToChildren.get(currentUuid) || new Set();

    for (const childUuid of children) {
      const parents = currentRelationships.get(childUuid);
      if (!parents || !parents.has(currentUuid)) continue;

      parents.delete(currentUuid);
      const childEntry = rowsByUuid.get(childUuid);
      if (!childEntry) continue;

      if (parents.size === 0) {
        uuidsToDelete.add(childUuid);
        pendingUpdates.delete(childUuid);

        if (!queued.has(childUuid)) {
          processingQueue.push(childUuid);
          queued.add(childUuid);
        }
        continue;
      }

      const remainingDestinations = childEntry.parsedDestinations
        .filter(destination => parents.has(destination.uuid));
      const newQty = remainingDestinations
        .reduce((sum, destination) => sum + destination.quantity, 0);

      pendingUpdates.set(childUuid, {
        newDestination: formatDestinationCell(remainingDestinations),
        newQty
      });
      uuidsToDelete.delete(childUuid);
    }
  }

  // Корневая строка сборочной единицы с пустым назначением остаётся в Базе.
  const rootEntry = rowsByUuid.get(rootUuid);
  const rootParents = currentRelationships.get(rootUuid);
  const rootHasParents = rootParents && rootParents.size > 0;

  if (rootHasParents || !rootEntry.destination) {
    uuidsToDelete.delete(rootUuid);
  } else {
    uuidsToDelete.add(rootUuid);
  }

  const rowsToDelete = [];
  for (const uuid of uuidsToDelete) {
    const entry = rowsByUuid.get(uuid);
    if (entry) rowsToDelete.push(entry.rowNum);
  }

  // Одна структурная операция на каждый непрерывный блок вместо deleteRow() на
  // каждую строку.
  deleteRowsInBlocks_(baseSheet, rowsToDelete);

  if (pendingUpdates.size > 0) {
    // После удаления номера строк изменились. Перечитываем только UUID-столбец D.
    const currentLastRow = baseSheet.getLastRow();
    const currentRowsByUuid = new Map();

    if (currentLastRow >= 2) {
      const uuidValues = baseSheet
        .getRange(2, COL_UUID, currentLastRow - 1, 1)
        .getValues();

      uuidValues.forEach((rowData, index) => {
        const uuid = String(rowData[0] || '').trim();
        if (uuid) currentRowsByUuid.set(uuid, index + 2);
      });
    }

    for (const [uuid, update] of pendingUpdates.entries()) {
      const rowNum = currentRowsByUuid.get(uuid);
      if (!rowNum) continue;

      baseSheet.getRange(rowNum, COL_DESTINATION).setValue(update.newDestination);
      baseSheet.getRange(rowNum, COL_QUANTITY).setValue(update.newQty);
    }
  }

  ss.toast('Откат спецификации завершен.', 'Информация', 5);
}

/**
 * Разбирает многострочное назначение вида
 * «Р155.02 = 10 (b7ee7162-37a7-49a6-b36c-15f5845b6852)».
 */
function parseDestinationCell(cellValue) {
  if (!cellValue) return [];

  const destinations = [];
  const lines = String(cellValue).split(/\r?\n/);
  const regex = /(.+)\s*=\s*(\d+(?:[.,]\d+)?)\s*\(([-\w]{36})\)/;

  for (const line of lines) {
    const match = line.trim().match(regex);
    if (!match) continue;

    destinations.push({
      name: match[1].trim(),
      quantity: parseFloat(match[2].replace(',', '.')),
      uuid: match[3].trim()
    });
  }

  return destinations;
}

/** Форматирует массив назначений обратно в многострочную ячейку. */
function formatDestinationCell(destinations) {
  if (!destinations || destinations.length === 0) return '';
  return destinations
    .map(destination => `${destination.name} = ${destination.quantity} (${destination.uuid})`)
    .join('\n');
}
