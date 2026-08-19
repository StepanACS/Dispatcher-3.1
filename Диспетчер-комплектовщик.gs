/**
 * Функция-обработчик для листа "Диспетчер- комплектовщик".
 * Вызывается, если в столбце K (номер 11) установлен чекбокс TRUE.
 */

function handleSupplyOrder(ss, sheet, range, row) {
  Logger.log ("Запущена обработка диспетчера - комплектовщика");
  const baseSheet = ss.getSheetByName("База");
  
  const itemId = sheet.getRange(row, 4).getValue().toString().trim();
  const itemName = sheet.getRange(row, 5).getValue().toString().trim();
 
 // Поиск строки в "База" по itemId
  const baseRow = findRowById(baseSheet, itemId); 
  Logger.log ("Найдена строка в базе по ID, номер строки " + baseRow);
  
  // Получаем назначение из столбца F 
  var purpose = baseSheet.getRange(baseRow, COL_DESTINATION).getValue().toString().trim();
  // Получаем количество из столбца H
  const qty = baseSheet.getRange(baseRow, COL_QUANTITY).getValue();
  //Получаем количество дефицита
  const deficit  = baseSheet.getRange(baseRow, COL_DEFICIT).getValue();
  //Получаем способ поставки из столбца Y
  const supplyWay = baseSheet.getRange(baseRow, COL_MODE).getValue();
  Logger.log (`Назначение ${purpose}, количество дефицита ${deficit} и способ поставки ${supplyWay}`);
  
  if (deficit > 0){  
    if (supplyWay == "В снабжение"){
      setSupplyOrder(itemName, qty, purpose); //Создаем строку в снабжении
    }
  }
  Logger.log ("Создана строка в снабжении ");

  // Сбрасываем исходный чекбокс, если изменение произошло вне "База"
  if (sheet.getName() != baseSheet.getName()){
    range.setValue(false);
  }
  checkboxOk(baseSheet.getRange(baseRow, COL_PICKER)); //Ставим отметку о выполнении на листе База
   Logger.log ("Поставлена отметка в Базе ");
  return itemId;
 
}

function setSupplyOrder(name, qty, purpose){
  const supplyUrl = "https://docs.google.com/spreadsheets/d/1j_5fbonKWN5O5F7EeHDRjanm5d-fhWR0K2m2SO_EvSc/edit?gid=0#gid=0"; //Ссылка на документ "снабжение"
  const supplySpreadsheet = SpreadsheetApp.openByUrl(supplyUrl);  //открываем документ "снабжение"
  const supplySheet = supplySpreadsheet.getSheetByName("Сводный");  //Открываем в снабжении лист "сводный"
  
  const currentDate = new Date(); //Присваиваем переменной сегодняшнюю дату
  const newOrderRow = getLastRowNumber(supplySheet, 2) + 1; // Определяем номер строки для новой записи

  // Создаем массив с данными для новой строки.
  // Важно: массив должен содержать элементы для всех столбцов, в которые вы хотите записать данные.
  // Если между столбцами 4 (name) и 8 (purpose) есть пустые столбцы,
  // их нужно представить пустыми строками ("") или null в массиве.
  // Здесь мы заполняем от 1 до 8 столбца.
  const rowData = [
    currentDate,                     // Столбец 1
    "Script - Диспетчер 3.0",        // Столбец 2
    qty,                             // Столбец 3
    name,                            // Столбец 4
    "",                              // Столбец 5 (пустой, если не используется)
    "",                              // Столбец 6 (пустой, если не используется)
    "",                              // Столбец 7 (пустой, если не используется)
    purpose                          // Столбец 8
  ];

  // Записываем массив данных в диапазон
  // getRange(строка, столбец, количество_строк, количество_столбцов)
  // Мы записываем 1 строку, начиная с 1 столбца, на 8 столбцов в ширину.
  supplySheet.getRange(newOrderRow, 1, 1, rowData.length).setValues([rowData]);

  Logger.log(`Новая строка заказа на снабжение добавлена в строку ${newOrderRow} листа "Сводный".`);
}

function getLastRowNumber(sheet, column) {  //Определяет последнюю занятую строку в столбце
  var values = sheet.getRange(1, column, sheet.getLastRow(), 1).getValues();
  
  for (var i = values.length - 1; i >= 0; i--) {
    if (values[i][0] !== "") {
      return i + 1; // Возвращаем номер строки (индекс + 1)
    }
  }
}

function getSpec(itemName, link){
// Открытие таблицы по ссылке: извлекаем ID из ссылки и используем openById
  Logger.log(`Начало открытия таблицы по ссылке для "${itemName}" в ${new Date().toLocaleString()}`);
  let targetSS;
  try {
    var sheetId = extractSheetId(link);
    if (!sheetId) {
      Logger.log(`Не удалось извлечь ID из ссылки: "${link}" для "${itemName}".`);
    }
    targetSS = SpreadsheetApp.openById(sheetId);
    Logger.log(`Таблица по ID "${sheetId}" для "${itemName}" успешно открыта в ${new Date().toLocaleString()}.`);
  } catch (error) {
    Logger.log(`Не удалось открыть таблицу по ссылке для "${itemName}": ${link} в ${new Date().toLocaleString()}. Ошибка: ${error}`);
  }
  
  if (!targetSS) {
    Logger.log(`Не удалось открыть таблицу по ссылке для "${itemName}": ${link}`);
  }

// Получаем лист спецификации
  const specSheet = targetSS.getSheetByName(itemName);
  if (!specSheet) {
    Logger.log(`Не найден лист '${itemName}' в открытой таблице для "${itemName}".`);
  }

// Поиск строки "Итого" в столбце B (начиная с 5-й строки)
  const firstDataRow = 5;
  const lastRowSpec = specSheet.getLastRow();
  if (lastRowSpec < firstDataRow) {
    Logger.log(`Нет данных для обработки в спецификации изделия "${itemName}".`);
  }
  //Считываем содержимое столбца B
  const colBValues = specSheet.getRange(firstDataRow, 2, lastRowSpec - firstDataRow + 1, 1).getValues();
  let itogoRow = null;
  for (let i = 0; i < colBValues.length; i++) {
    const cellValue = (colBValues[i][0] || "").toString().toLowerCase();
    if (cellValue === "итого") {
      itogoRow = firstDataRow + i;
      break;
    }
  }
  
  if (!itogoRow) {
    Logger.log(`Не найдена строка 'Итого' в столбце B спецификации изделия "${itemName}".`);
  }
  
  const lastDataRow = itogoRow - 1;
  if (lastDataRow < firstDataRow) {
    Logger.log(`Диапазон спецификации пуст для изделия "${itemName}".`);
  }

  // Получаем данные спецификации (диапазон из столбцов B-D)
  const numRows = lastDataRow - firstDataRow + 1;
  const rangeSpec = specSheet.getRange(firstDataRow, 2, numRows, 3);
  const dataSpec = rangeSpec.getValues();
  return dataSpec;
}

/**
 * Функция getSdId(name)
 * Считывает данные с листа "SD ID", где:
 *   - В колонке A содержится идентификатор,
 *   - В колонке B – имя.
 * Функция ищет строку, где во втором столбце (B) содержится переданное имя, 
 * и возвращает идентификатор из первого столбца (A).
 *
 * @param {string} name - Имя для поиска.
 * @returns {string|null} Идентификатор, если найден, или null в противном случае.
 */
function getSdId(ss, name) {
  var sdSheet = ss.getSheetByName("SD ID");
  if (!sdSheet) {
    Logger.log("Лист 'SD ID' не найден.");
    return null;
  }
  
  // Считываем все данные из колонок A и B (с первой по последнюю заполненную строку)
  var lastRow = sdSheet.getLastRow();
  if (lastRow < 1) return null;
  var data = sdSheet.getRange(1, 1, lastRow, 2).getValues();
  
  // Приводим искомое имя к стандартному виду (без лишних пробелов)
  var targetName = name.toString().trim();
  
  // Проходим по массиву и ищем совпадение во втором столбце (индекс 1)
  for (var i = 0; i < data.length; i++) {
    if (data[i][1].toString().trim() === targetName) {
      return data[i][0].toString().trim(); // Возвращаем идентификатор из столбца A (индекс 0)
    }
  }
  
  // Если ничего не найдено, возвращаем null
  return null;
}

