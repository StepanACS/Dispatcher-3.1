function handleProduction(ss, sheet, range, row) {
  const baseSheet = ss.getSheetByName('База');
  const itemId = String(sheet.getRange(row, COL_UUID).getValue() || '').trim();
  const baseRow = findRowById(baseSheet, itemId);

  if (!baseRow) {
    Logger.log(`handleProduction: UUID ${itemId} не найден в Базе.`);
    return null;
  }

  // Две соседние отметки записываем одним вызовом Sheets API.
  const now = new Date();
  baseSheet
    .getRange(baseRow, COL_PRODUCTION, 1, COL_QUALITY_CHECK - COL_PRODUCTION + 1)
    .setValues([[now, now]]);

  if (sheet.getSheetId() !== baseSheet.getSheetId()) {
    range.setValue(false);
  }

  return itemId;
}
