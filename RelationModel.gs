/**
 * Временное compatibility mirror спецификаций.
 *
 * После перехода на Variant B источником истины для иерархии являются поля
 * «База.parent_uuid» и «База.root_uuid». Лист «Связи спецификации» сохраняется
 * временно, чтобы старые вспомогательные пути и диагностика могли работать до
 * завершения smoke-test. Формулы «База» и основная логика сворачивания от него
 * больше не зависят.
 *
 * A child_uuid | B parent_uuid | C child_name | D parent_name | E quantity
 */
const SPEC_RELATION_COLUMNS = Object.freeze({
  CHILD_UUID: 1,
  PARENT_UUID: 2,
  CHILD_NAME: 3,
  PARENT_NAME: 4,
  QUANTITY: 5,
  WIDTH: 5
});

function getSpecRelationSheet_(ss) {
  const sheet = ss.getSheetByName(SPEC_RELATIONS_SHEET_NAME);
  if (!sheet) throw new Error(`Не найден служебный лист «${SPEC_RELATIONS_SHEET_NAME}».`);
  return sheet;
}

function normalizeSpecRelation_(relation) {
  if (!relation) return null;
  const childUuid = String(relation.childUuid || relation.child_uuid || '').trim();
  const parentUuid = String(relation.parentUuid || relation.parent_uuid || '').trim();
  if (!childUuid || !parentUuid) return null;
  return {
    childUuid,
    parentUuid,
    childName: String(relation.childName || relation.child_name || '').trim(),
    parentName: String(relation.parentName || relation.parent_name || '').trim(),
    quantity: Number(relation.quantity) || 0
  };
}

function readSpecRelations_(ss) {
  const sheet = getSpecRelationSheet_(ss);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  return sheet.getRange(2, 1, lastRow - 1, SPEC_RELATION_COLUMNS.WIDTH).getValues()
    .map(row => normalizeSpecRelation_({
      childUuid: row[0], parentUuid: row[1], childName: row[2],
      parentName: row[3], quantity: row[4]
    }))
    .filter(Boolean);
}

/** Полностью заменяет тело зеркала, сохраняя заголовок. */
function replaceSpecRelations_(ss, relations) {
  const sheet = getSpecRelationSheet_(ss);
  const normalized = (relations || []).map(normalizeSpecRelation_).filter(Boolean);
  const maxRows = sheet.getMaxRows();
  if (maxRows > 1) sheet.getRange(2, 1, maxRows - 1, SPEC_RELATION_COLUMNS.WIDTH).clearContent();
  if (!normalized.length) return;
  if (normalized.length + 1 > sheet.getMaxRows()) {
    sheet.insertRowsAfter(sheet.getMaxRows(), normalized.length + 1 - sheet.getMaxRows());
  }
  sheet.getRange(2, 1, normalized.length, SPEC_RELATION_COLUMNS.WIDTH).setValues(
    normalized.map(r => [r.childUuid, r.parentUuid, r.childName, r.parentName, r.quantity])
  );
}

/**
 * Добавляет/обновляет зеркало. Пара child_uuid + parent_uuid уникальна.
 * При Variant B у одного child_uuid должен быть ровно один parent_uuid.
 */
function appendSpecRelations_(ss, relations) {
  const incoming = (relations || []).map(normalizeSpecRelation_).filter(Boolean);
  if (!incoming.length) return;

  const byEdge = new Map();
  for (const relation of readSpecRelations_(ss)) {
    byEdge.set(`${relation.childUuid}\u0000${relation.parentUuid}`, relation);
  }
  for (const relation of incoming) {
    byEdge.set(`${relation.childUuid}\u0000${relation.parentUuid}`, relation);
  }
  replaceSpecRelations_(ss, [...byEdge.values()]);
}

/** Legacy compatibility helper. Автоматическое слияние Variant B не использует. */
function remapSpecRelationUuid_(ss, fromUuid, toUuid) {
  const from = String(fromUuid || '').trim();
  const to = String(toUuid || '').trim();
  if (!from || !to || from === to) return;

  const byEdge = new Map();
  for (const rel of readSpecRelations_(ss)) {
    if (rel.childUuid === from) rel.childUuid = to;
    if (rel.parentUuid === from) rel.parentUuid = to;
    if (rel.childUuid === rel.parentUuid) continue;
    byEdge.set(`${rel.childUuid}\u0000${rel.parentUuid}`, rel);
  }
  replaceSpecRelations_(ss, [...byEdge.values()]);
}

/** Удаляет конкретный набор рёбер из compatibility mirror. */
function removeSpecRelationEdges_(ss, edges) {
  const remove = new Set((edges || []).map(edge => {
    const rel = normalizeSpecRelation_(edge);
    return rel ? `${rel.childUuid}\u0000${rel.parentUuid}` : '';
  }).filter(Boolean));
  if (!remove.size) return;
  replaceSpecRelations_(ss, readSpecRelations_(ss).filter(rel =>
    !remove.has(`${rel.childUuid}\u0000${rel.parentUuid}`)
  ));
}

/**
 * Metadata для старого рекурсивного генератора. writeBaseRows_ преобразует её
 * в реальные поля parent_uuid/root_uuid перед записью строки и затем временно
 * дублирует связь в compatibility mirror.
 */
function attachSpecRelation_(row, relation) {
  const normalized = normalizeSpecRelation_(relation);
  if (!row || !normalized) return row;
  Object.defineProperty(row, '_specRelation', {
    value: normalized,
    enumerable: false,
    configurable: true
  });
  return row;
}
