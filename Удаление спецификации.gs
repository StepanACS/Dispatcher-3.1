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
 * Вариант Б: граф спецификации читается непосредственно из «База.parent_uuid».
 * Выбранная строка сохраняется как узел сборки, а все её потомки очищаются.
 * Благодаря правилу «одна строка = один родитель» никакого перераспределения
 * quantity между несколькими назначениями здесь больше нет.
 */
function removeSpecification(initialParentUuid, ss) {
  const baseSheet = ss.getSheetByName(BASE_SHEET_NAME);
  if (!baseSheet) {
    ss.toast('Ошибка: отсутствует лист «База».', 'Ошибка', 5);
    return;
  }

  const targetUuid = String(initialParentUuid || '').trim();
  if (!targetUuid) return;
  SpreadsheetApp.flush();

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
  while (queue.length) {
    const uuid = queue.shift();
    if (!uuid || descendants.has(uuid)) continue;
    descendants.add(uuid);
    for (const child of childrenByParent.get(uuid) || []) queue.push(child);
  }

  const rowsToClear = [...descendants]
    .map(uuid => rowByUuid.get(uuid))
    .filter(Boolean);
  clearBaseRawRows_(baseSheet, rowsToClear);

  // Поддерживаем старый лист только как compatibility mirror, пока проверка
  // повторного развёртывания ещё использует его. Расчёты от него не зависят.
  if (ss.getSheetByName(SPEC_RELATIONS_SHEET_NAME) && descendants.size) {
    replaceSpecRelations_(ss, readSpecRelations_(ss).filter(rel =>
      !descendants.has(rel.childUuid)
    ));
  }

  SpreadsheetApp.flush();
  ss.toast('Откат спецификации завершен.', 'Информация', 5);
}

/** Legacy parser: только для диагностики старых текстовых Назначений. */
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
