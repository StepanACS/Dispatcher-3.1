/** Сворачивает спецификацию изделия, указанного в текущей строке. */
function handleUndeployment(ss, sheet, row) {
  const baseSheet = ss.getSheetByName(BASE_SHEET_NAME);
  let itemId = '';
  if (sheet.getSheetId() === baseSheet.getSheetId()) {
    itemId = String(readBaseFields_(baseSheet, row, ['uuid']).uuid || '').trim();
  } else {
    itemId = String(sheet.getRange(row, ROLE_COL_UUID).getValue() || '').trim();
  }
  if (!itemId) return;
  removeSpecification(itemId, ss);
}

/**
 * Сворачивает ветку по нормализованному графу «Связи спецификации».
 * Физические строки «База» не удаляются: лишние записи очищаются, а записи с
 * другими родителями получают пересчитанные destination/quantity.
 */
function removeSpecification(initialParentUuid, ss) {
  const baseSheet = ss.getSheetByName(BASE_SHEET_NAME);
  const relationSheet = ss.getSheetByName(SPEC_RELATIONS_SHEET_NAME);
  if (!baseSheet || !relationSheet) {
    ss.toast('Ошибка: отсутствует «База» или «Связи спецификации».', 'Ошибка', 5);
    return;
  }

  const rootUuid = String(initialParentUuid || '').trim();
  if (!rootUuid) return;
  SpreadsheetApp.flush();

  const lastBaseRow = baseLastDataRow_(baseSheet);
  if (lastBaseRow < 2) return;

  const ssArch = getBaseArchitecture_(ss);
  const baseKeys = ['uuid', 'item_name', 'destination', 'quantity'];
  const cols = baseKeys.map(key => ({ key, col: baseCol_(ss, key) }));
  const minCol = Math.min(...cols.map(x => x.col));
  const maxCol = Math.max(...cols.map(x => x.col));
  const baseData = baseSheet.getRange(2, minCol, lastBaseRow - 1, maxCol - minCol + 1).getValues();
  const idx = key => cols.find(x => x.key === key).col - minCol;

  const rowsByUuid = new Map();
  baseData.forEach((r, i) => {
    const uuid = String(r[idx('uuid')] || '').trim();
    if (!uuid) return;
    rowsByUuid.set(uuid, {
      rowNum: i + 2,
      name: String(r[idx('item_name')] || '').trim(),
      destination: String(r[idx('destination')] || '').trim(),
      quantity: r[idx('quantity')]
    });
  });
  if (!rowsByUuid.has(rootUuid)) {
    Logger.log(`removeSpecification: UUID ${rootUuid} не найден в Базе.`);
    return;
  }

  // A:E = child_uuid, parent_uuid, child_name, parent_name, quantity.
  const relLastRow = relationSheet.getLastRow();
  const relData = relLastRow >= 2 ? relationSheet.getRange(2, 1, relLastRow - 1, 5).getValues() : [];
  const childToRelations = new Map();
  const parentToChildren = new Map();

  for (const rel of relData) {
    const childUuid = String(rel[0] || '').trim();
    const parentUuid = String(rel[1] || '').trim();
    if (!childUuid || !parentUuid) continue;
    const relation = {
      childUuid,
      parentUuid,
      parentName: String(rel[3] || '').trim(),
      quantity: Number(rel[4]) || 0
    };
    if (!childToRelations.has(childUuid)) childToRelations.set(childUuid, []);
    childToRelations.get(childUuid).push(relation);
    if (!parentToChildren.has(parentUuid)) parentToChildren.set(parentUuid, new Set());
    parentToChildren.get(parentUuid).add(childUuid);
  }

  const queue = [rootUuid];
  const queued = new Set(queue);
  const uuidsToClear = new Set([rootUuid]);
  const pendingUpdates = new Map();
  const removedParentEdges = new Map();

  while (queue.length) {
    const currentUuid = queue.shift();
    const children = parentToChildren.get(currentUuid) || new Set();

    for (const childUuid of children) {
      const allRelations = childToRelations.get(childUuid) || [];
      if (!allRelations.some(rel => rel.parentUuid === currentUuid)) continue;
      if (!removedParentEdges.has(childUuid)) removedParentEdges.set(childUuid, new Set());
      removedParentEdges.get(childUuid).add(currentUuid);

      const removed = removedParentEdges.get(childUuid);
      const remaining = allRelations.filter(rel => !removed.has(rel.parentUuid));
      if (!remaining.length) {
        uuidsToClear.add(childUuid);
        pendingUpdates.delete(childUuid);
        if (!queued.has(childUuid)) {
          queued.add(childUuid);
          queue.push(childUuid);
        }
      } else {
        uuidsToClear.delete(childUuid);
        pendingUpdates.set(childUuid, {
          newDestination: remaining
            .map(rel => `${rel.parentName} = ${rel.quantity} (${rel.parentUuid})`)
            .join('\n'),
          newQty: remaining.reduce((sum, rel) => sum + rel.quantity, 0)
        });
      }
    }
  }

  // Верхнеуровневая сборка с пустым назначением остаётся как самостоятельная строка.
  const rootEntry = rowsByUuid.get(rootUuid);
  const rootRelations = childToRelations.get(rootUuid) || [];
  if (rootRelations.length || !rootEntry.destination) uuidsToClear.delete(rootUuid);

  const rowsToClear = [];
  for (const uuid of uuidsToClear) {
    const entry = rowsByUuid.get(uuid);
    if (entry) rowsToClear.push(entry.rowNum);
  }
  clearBaseRawRows_(baseSheet, rowsToClear);

  for (const [uuid, update] of pendingUpdates) {
    const entry = rowsByUuid.get(uuid);
    if (!entry || uuidsToClear.has(uuid)) continue;
    writeBaseFields_(baseSheet, entry.rowNum, {
      destination: update.newDestination,
      quantity: update.newQty
    });
  }

  ss.toast('Откат спецификации завершен.', 'Информация', 5);
}

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

function formatDestinationCell(destinations) {
  if (!destinations || !destinations.length) return '';
  return destinations.map(d => `${d.name} = ${d.quantity} (${d.uuid})`).join('\n');
}
