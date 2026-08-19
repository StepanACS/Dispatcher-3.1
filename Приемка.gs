function handleAcceptance(ss, sheet, range, row) {
  const baseSheet = ss.getSheetByName('База');

  // Приемка из «Базы» запрещена: на листе роли требуется указать фактическое количество.
  if (sheet.getSheetId() === baseSheet.getSheetId()) {
    range.setValue(false);
    ss.toast('Для выполнения приемки перейдите на соответствующий лист', 'Ошибка', 5);
    return null;
  }

  // D:M читаются одним запросом: UUID и фактическое количество (M).
  const rowData = sheet.getRange(row, COL_UUID, 1, 13 - COL_UUID + 1).getValues()[0];
  const itemId = String(rowData[0] || '').trim();
  const factQty = rowData[13 - COL_UUID];

  if (factQty === '' || factQty === null || typeof factQty === 'undefined') {
    range.setValue(false);
    ss.toast('Не указано фактическое количество или неверный формат записи', 'Ошибка', 5);
    return null;
  }

  const baseRow = findRowById(baseSheet, itemId);
  if (!baseRow) {
    ss.toast(`Изделие ${itemId} не найдено в Базе`, 'Ошибка', 5);
    return null;
  }

  // Сначала очищаем ручные поля QUERY-листа. После отметки COL_ACCEPTANCE строка
  // переходит в «Готово» и может исчезнуть с текущего листа.
  range.setValue(false);
  sheet.getRange(row, 13).clearContent();

  baseSheet.getRange(baseRow, COL_FACT_QTY).setValue(factQty);
  checkboxOk(baseSheet.getRange(baseRow, COL_ACCEPTANCE));
  return itemId;
}
