function handleAcceptance(ss, sheet, range, row) {
  const baseSheet = ss.getSheetByName(BASE_SHEET_NAME);
  if (sheet.getSheetId() === baseSheet.getSheetId()) {
    range.setValue(false);
    ss.toast('Для выполнения приемки перейдите на соответствующий лист', 'Ошибка', 5);
    return null;
  }

  const rowData = sheet.getRange(row, ROLE_COL_UUID, 1, 13 - ROLE_COL_UUID + 1).getValues()[0];
  const itemId = String(rowData[0] || '').trim();
  const factQty = rowData[13 - ROLE_COL_UUID];
  if (factQty === '' || factQty === null || typeof factQty === 'undefined') {
    range.setValue(false);
    ss.toast('Не указано фактическое количество или неверный формат записи', 'Ошибка', 5);
    return null;
  }

  const baseRow = findBaseRowByUuid_(baseSheet, itemId);
  if (!baseRow) {
    ss.toast(`Изделие ${itemId} не найдено в Базе`, 'Ошибка', 5);
    return null;
  }

  range.setValue(false);
  sheet.getRange(row, 13).clearContent();
  writeBaseFields_(baseSheet, baseRow, {
    actual_made: factQty,
    stage_acceptance: new Date()
  });
  return itemId;
}
