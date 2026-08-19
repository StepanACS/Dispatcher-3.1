/**
 * LEGACY ADMIN UTILITY.
 *
 * Проверяет наличие рабочих листов из «Настройки», создаёт недостающие и
 * восстанавливает базовую QUERY/checkbox-структуру. Не вызывается из onEdit.
 *
 * ВАЖНО: в текущей «Диспетчер 3.1 опытный» часть листов уже содержит дополнительные
 * служебные QUERY в M:N и более широкие исходные диапазоны. Поэтому до отдельного
 * рефакторинга архитектуры таблицы эту функцию НЕ следует запускать для массового
 * восстановления существующих листов: она описывает исторический минимальный шаблон.
 */
function checkAndUpdateSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = ss.getSheetByName('Настройки');
  if (!settingsSheet) {
    Logger.log('Лист «Настройки» не найден.');
    return;
  }

  const lastSettingsRow = settingsSheet.getLastRow();
  if (lastSettingsRow < 2) return;

  const settingsValues = settingsSheet.getRange(2, 1, lastSettingsRow - 1, 2).getValues();
  const sheetsByName = new Map(ss.getSheets().map(sheet => [sheet.getName(), sheet]));
  const checkboxRule = SpreadsheetApp.newDataValidation().requireCheckbox().build();

  for (const settingsRow of settingsValues) {
    const sheetName = String(settingsRow[0] || '').trim();
    const baseColumnLetter = String(settingsRow[1] || '').trim().toUpperCase();
    if (!sheetName || !columnLetterToNumber(baseColumnLetter)) continue;

    let targetSheet = sheetsByName.get(sheetName);
    if (!targetSheet) {
      targetSheet = ss.insertSheet(sheetName);
      sheetsByName.set(sheetName, targetSheet);
    }

    const formula = `=QUERY('База'!A:Y; "select A, B, C, D, E, F, G, H, I, J where Z = '${sheetName}'"; 1)`;
    targetSheet.getRange(1, 1, 1, 15).clearContent();
    targetSheet.getRange('A1').setFormula(formula);
    targetSheet.getRange('K1').setValue('Отметка о выполнении');

    const checkboxSettings = getSettingsForSheet(sheetName);
    if (checkboxSettings) {
      const rows = Math.max(targetSheet.getMaxRows() - 1, 1);
      targetSheet
        .getRange(2, checkboxSettings.firstColumn, rows, checkboxSettings.count)
        .setDataValidation(checkboxRule);
    }
  }
}
