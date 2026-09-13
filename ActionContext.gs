/**
 * Единый semantic context для бизнес-действия.
 * Обработчики этапов не должны знать физические столбцы «База» и не должны
 * предполагать, что вызываются только с role sheet.
 */
function resolveActionContext_(ss, sheet, row) {
  if (!ss || !sheet || !row) return null;
  const baseSheet = ss.getSheetByName(BASE_SHEET_NAME);
  if (!baseSheet) throw new Error('Лист «База» не найден.');

  let itemId = '';
  let itemName = '';
  let baseRow = null;

  if (sheet.getSheetId() === baseSheet.getSheetId()) {
    baseRow = row;
    const data = readBaseFields_(baseSheet, baseRow, ['uuid', 'item_name']);
    itemId = String(data.uuid || '').trim();
    itemName = String(data.item_name || '').trim();
  } else {
    const data = sheet.getRange(row, ROLE_COL_UUID, 1, ROLE_COL_ITEMNAME - ROLE_COL_UUID + 1).getValues()[0];
    itemId = String(data[0] || '').trim();
    itemName = String(data[ROLE_COL_ITEMNAME - ROLE_COL_UUID] || '').trim();
    if (itemId) baseRow = findBaseRowByUuid_(baseSheet, itemId);
  }

  if (!itemId || !baseRow) return null;
  return { ss, sheet, baseSheet, row, baseRow, itemId, itemName };
}

function resetRoleCheckbox_(context, range) {
  if (!context || !range) return;
  if (context.sheet.getSheetId() !== context.baseSheet.getSheetId()) range.setValue(false);
}

/**
 * Идемпотентная отметка этапа. Возвращает false, если этап уже был завершён.
 */
function markStageOnce_(context, fieldKey, when) {
  if (!context) return false;
  const current = readBaseFields_(context.baseSheet, context.baseRow, [fieldKey])[fieldKey];
  if (current) return false;
  writeBaseFields_(context.baseSheet, context.baseRow, { [fieldKey]: when || new Date() });
  return true;
}
