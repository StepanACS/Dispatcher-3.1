/** Сворачивает спецификацию изделия, указанного в текущей строке. */
function handleUndeployment(ss, sheet, row) {
  const context = resolveActionContext_(ss, sheet, row);
  if (!context) return;
  removeSpecification(context.itemId, ss);
}

/**
 * Variant B: граф спецификации читается непосредственно из «База.parent_uuid».
 * Выбранная строка сохраняется как узел сборки, а все её потомки очищаются.
 */
function removeSpecification(initialParentUuid, ss) {
  const baseSheet = ss.getSheetByName(BASE_SHEET_NAME);
  if (!baseSheet) {
    ss.toast('Ошибка: отсутствует лист «База».', 'Ошибка', 5);
    return;
  }

  const targetUuid = String(initialParentUuid || '').trim();
  if (!targetUuid) return;

  const lastRow = baseLastDataRow_(baseSheet);
  if (lastRow < 2) return;
  const rowCount = lastRow - 1;
  const uuidCol = baseCol_(ss, 'uuid');
  const parentCol = baseCol_(ss, 'parent_uuid');
  const uuids = baseSheet.getRange(2, uuidCol, rowCount, 1).getDisplayValues().flat();
  const parents = baseSheet.getRange(2, parentCol, rowCount, 1).getDisplayValues().flat();

  const rowByUuid = new Map();
  const childrenByParent = new Map();
  for (let i = 0; i < uuids.length; i++) {
    const uuid = String(uuids[i] || '').trim();
    const parent = String(parents[i] || '').trim();
    if (!uuid) continue;
    rowByUuid.set(uuid, i + 2);
    if (parent) {
      if (!childrenByParent.has(parent)) childrenByParent.set(parent, []);
      childrenByParent.get(parent).push(uuid);
    }
  }

  if (!rowByUuid.has(targetUuid)) {
    Logger.log(`removeSpecification: UUID ${targetUuid} не найден в Базе.`);
    return;
  }

  const descendants = new Set();
  const queue = [...(childrenByParent.get(targetUuid) || [])];
  for (let head = 0; head < queue.length; head++) {
    const uuid = queue[head];
    if (!uuid || descendants.has(uuid)) continue;
    descendants.add(uuid);
    for (const child of childrenByParent.get(uuid) || []) queue.push(child);
  }

  const rowsToClear = [...descendants]
    .map(uuid => rowByUuid.get(uuid))
    .filter(Boolean);
  clearBaseRawRows_(baseSheet, rowsToClear);
  ss.toast('Откат спецификации завершен.', 'Информация', 5);
}
