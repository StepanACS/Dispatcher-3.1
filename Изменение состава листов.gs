/**
 * Проверяет наличие листов из списка настроек и создает недостающие листы.
 * На каждом листе устанавливает или обновляет формулу в ячейке A1.
 */
function checkAndUpdateSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = ss.getSheetByName("Настройки");
  const statusColumnLetter = "Z";  //на листе База
  const readyMarkColumnLetter = "K";
  const readyMarkColumn = 11;
  
  if (!settingsSheet) {
    Logger.log('Лист "Настройки" не найден.');
    return;
  }
  
  // Читаем все данные из листа "Настройки"
  const settingsRange = settingsSheet.getDataRange();
  const settingsValues = settingsRange.getValues();
  
  // Предполагаем, что заголовки в первой строке, данные начинаются со второй строки
  for (let i = 1; i < settingsValues.length; i++) {
    const sheetName = settingsValues[i][0].toString().trim(); // Столбец A
    const baseColumnLetter = settingsValues[i][1].toString().trim().toUpperCase(); // Столбец B
    if (!sheetName || !baseColumnLetter) {
      Logger.log(`Пропущена строка ${i + 1}: отсутствует название листа или столбца.`);
      continue; // Пропустить пустые строки или строки с отсутствующими данными
    }
    
    // Проверяем, существует ли лист
    let targetSheet = ss.getSheetByName(sheetName);
    if (!targetSheet) {
      // Создаем новый лист
      targetSheet = ss.insertSheet(sheetName);
      Logger.log(`Создан новый лист: ${sheetName}`);
    } else {
      Logger.log(`Лист "${sheetName}" уже существует.`);
    }
    
    // Преобразуем буквенное обозначение столбца в числовое
    const baseColumnNumber = columnLetterToNumber(baseColumnLetter);
    if (!baseColumnNumber) {
      Logger.log(`Неверное обозначение столбца "${baseColumnLetter}" для листа "${sheetName}".`);
      continue; // Пропустить, если столбец указан неверно
    }
    
    // Создаем формулу для ячейки A1 с использованием ';' в качестве разделителя
   const formula = `=QUERY('База'!A:Y; "select A, B, C, D, E, F, G, H, I, J where ${statusColumnLetter} = '${sheetName}'"; 1)`;
    
    targetSheet.getRange(1,1,1,15).clear();
    // Устанавливаем формулу в ячейку A1
    targetSheet.getRange("A1").setFormula(formula);    
    Logger.log(`Формула установлена на листе "${sheetName}": ${formula}`);
  
    //Записываем заголовок "Отметка о выполнении"
    targetSheet.getRange(readyMarkColumnLetter + "1").setValue("Отметка о выполнении");
    
    //Заполняем чекбоксами столбец "Отметка о выполнении"
    const dv = SpreadsheetApp.newDataValidation().requireCheckbox().build();
    targetSheet.getRange(2,readyMarkColumn,20,1).setDataValidation(dv)
  }
  
}

/**
 * Преобразует буквенное обозначение столбца в числовое.
 * Например, "A" -> 1, "B" -> 2, ..., "AA" -> 27 и т.д.
 * @param {string} letter - Буквенное обозначение столбца.
 * @return {number} Номер столбца.
 */
function columnLetterToNumber(letter) {
  let column = 0;
  const length = letter.length;
  for (let i = 0; i < length; i++) {
    const charCode = letter.charCodeAt(i);
    if (charCode < 65 || charCode > 90) { // Проверка на заглавные буквы A-Z
      return 0; // Возвращаем 0 при неверном формате
    }
    column += (charCode - 64) * Math.pow(26, length - i - 1);
  }
  return column;
}

