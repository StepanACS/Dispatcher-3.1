/**
 * Внутренний адаптер рекурсивного генератора спецификаций.
 *
 * Источник истины для иерархии — только База.parent_uuid/root_uuid. Отдельный
 * лист «Связи спецификации» новой архитектурой не читается и не записывается.
 * _specRelation существует лишь как неenumerable metadata между createSpecRow()
 * и prepareBaseHierarchyRows_(); после записи она исчезает.
 */
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
