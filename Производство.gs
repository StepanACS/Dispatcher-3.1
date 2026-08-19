function handleProduction(ss, sheet, range, row) {
  const baseSheet = ss.getSheetByName(BASE_SHEET_NAME);
  const itemId = String(sheet.getRange(row, ROLE_COL_UUID).getValue() || '').trim();
  const baseRow = findBaseRowByUuid_(baseSheet, itemId);
  if (!baseRow) {
    Logger.log(`handleProduction: UUID ${itemId} не найден в Базе.`);
    return null;
  }

  if (sheet.getSheetId() !== baseSheet.getSheetId()) range.setValue(false);
  const now = new Date();
  writeBaseFields_(baseSheet, baseRow, {
    stage_production: now,
    stage_quality_check: now
  });
  return itemId;
}
