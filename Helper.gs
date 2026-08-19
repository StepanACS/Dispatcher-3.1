function installEditTrigger() { // функция для создания устанавливаемого триггера на изменения в документе
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  ScriptApp.newTrigger('onEditHandler') // Исправлено название триггера
    .forSpreadsheet(spreadsheet)
    .onEdit()
    .create();
}

function buildBaseRowTemplate(futureIndex, sheetColumnCount) { // <-- sheetColumnCount
    const newRow = Array(sheetColumnCount).fill(""); // <-- Используем sheetColumnCount

    // Заполняем формулы в виде строк. Ссылки на столбцы используют COL_XXX - 1
    newRow[COL_DRAWING - 1] = `=IF(H${futureIndex}="";"";IF(E${futureIndex}="";"";IFERROR(IF(FILTER('Каталог ссылок'!B:B;'Каталог ссылок'!A:A=E${futureIndex})="";"Добавить ссылку в каталог";FILTER('Каталог ссылок'!B:B;'Каталог ссылок'!A:A=E${futureIndex}));"Добавить изделие в каталог ссылок")))`;
    newRow[COL_LINK - 1] = `=IF(H${futureIndex}="";"";IF(E${futureIndex}="";"";IFERROR(IF(FILTER('Каталог ссылок'!C:C;'Каталог ссылок'!A:A=E${futureIndex})="";"Добавить ссылку в каталог";FILTER('Каталог ссылок'!C:C;'Каталог ссылок'!A:A=E${futureIndex}));"Добавить изделие в каталог ссылок")))`;
    newRow[COL_UUID - 1] = Utilities.getUuid();
    newRow[COL_IN_STOCK - 1] = `=IFERROR(IF(E${futureIndex}<>"" ; FILTER('Остатки'!B:B; 'Остатки'!A:A=E${futureIndex}); ""); 0)`;
    newRow[COL_DEFICIT - 1] = `=IF((H${futureIndex}-I${futureIndex})<0; 0; H${futureIndex}-I${futureIndex})`;

    newRow[COL_READY_FOR_PRODUCTION - 1] = `=IFERROR(IF(L${futureIndex} = "Да";IF(OR(Z${futureIndex} = "Производство"; Z${futureIndex} = "Склад производства"); "В производстве";IF(OR(Z${futureIndex} = "ОТК"; Z${futureIndex} = "Приемка"; Z${futureIndex} = "Готово"); "Готово";IF(AA${futureIndex} = "В производство";IF(SUM(FILTER(J$2:J; ISNUMBER(SEARCH(E${futureIndex}; F$2:F)))) > 0; "Нет"; "Да");"Покупной")));"Нет");"Комплектующие не найдены")`;

    newRow[COL_IS_NEED - 1] = `=IF(J${futureIndex}>0; "Да"; "Нет")`;
    newRow[COL_STATUS - 1] = `=IFNA(INDEX('Настройки'!A$2:A$13; MATCH(FALSE; M${futureIndex}:W${futureIndex}; 0)); "Готово")`;
    newRow[COL_MODE - 1] = `=IF(E${futureIndex}="";"";IF(COUNTIF('Цена материалов'!A:A; E${futureIndex})=0;"В производство";"В снабжение"))`;
    newRow[COL_SDID - 1] = `=IFERROR(FILTER('SD ID'!A:A; 'SD ID'!B:B = E${futureIndex});"")`;
    newRow[COL_TRELLO_ID - 1] = ""; 
    newRow[COL_SD_ORDER_ID - 1] = ""; 
    newRow[COL_FACT_QTY - 1] = ""; 

    newRow[COL_TECHNOLOGIEST_NAME - 1] = "";
    newRow[COL_DESIGNER_NAME - 1] = "";
    newRow[COL_SUPPLY_ORDERED - 1] = `=IFERROR(SUM(FILTER('Снабжение'!C:C; 'Снабжение'!A:A = E${futureIndex})); "")`;
    newRow[COL_TIED_CAPITAL - 1] = `=IFERROR(IF(F${futureIndex}=""; FILTER('ГТ Себестоимость'!C:C; 'ГТ Себестоимость'!A:A = E${futureIndex}); 0); 0) * H${futureIndex}`;
    newRow[COL_QTY_SD - 1] = `=IFERROR(IF(E${futureIndex}<>""; FILTER('Остатки'!F:F; 'Остатки'!E:E = E${futureIndex}); ""); "")`;
    newRow[COL_ABC_CATEGORY - 1] = `=IFERROR(IF(F${futureIndex}=""; FILTER(ABC!B:B; ABC!A:A = E${futureIndex}); ""); "")`;
    newRow[COL_PRIORITY - 1] = "";
    newRow[COL_EQIPMENT - 1] = `=IFERROR(TEXTJOIN(", "; TRUE; FILTER('Диспетчер оснастки'!A:A; 'Диспетчер оснастки'!B:B = E${futureIndex})); "")`;
    newRow[COL_DEADLINE - 1] = "";



    // Заполняет столбцы с COL_DESIGN (M) до COL_READY (Y) значениями false (для чекбоксов).
    for (let c = COL_DESIGN - 1; c <= COL_READY - 1; c++) {
        newRow[c] = false; 
    }
    
    return newRow;
}


/**
 * Преобразует букву столбца в номер.
 * Например, 'A' -> 1, 'B' -> 2, ..., 'Z' -> 26, 'AA' -> 27 и т.д.
 *
 * @param {string} letter - Буква столбца.
 * @returns {number} Номер столбца.
 */
function columnLetterToNumber(letter) {
  let column = 0;
  const length = letter.length;
  for (let i = 0; i < length; i++) {
    column += (letter.charCodeAt(i) - 64) * Math.pow(26, length - i - 1);
  }
  return column;
}


function checkboxOk(range) { 
    range.setValue(new Date()); // Устанавливаем дату вместо просто TRUE для отметки
}


function extractSheetId(link) { 
    const match = link.match(/\/d\/([a-zA-Z0-9_-]+)/);
    return match ? match[1] : null;
}


function findRowById(sheet, itemId) {             
  const data = sheet.getDataRange().getValues();            
  for (let i = 1; i < data.length; i++) {            
    if (data[i][COL_UUID - 1] === itemId) {            
      return i + 1; // Возвращаем номер строки (1-индексированный)            
    }            
  }            
  return null;            
}


/**
 * Преобразует первую букву строки в верхний регистр.
 * @param {string} str 
 * @returns {string}
 */
function capitalizeFirstLetter(str) {
  str = str.toString().trim();
  return str ? str.charAt(0).toUpperCase() + str.slice(1) : "";
}

/**
 * Обработчик ручных записей в примечании
*/
function handleNote(ss, sheet, range, row, col, val) {
  const note = val.toString().trim();  //Сохраняем значение примечания с листа роли
  range.setValue(""); //очищаем исходную ячейку от текста
  const itemId = sheet.getRange(row, 4).getValue().toString().trim();
  const baseSheet = ss.getSheetByName("База");
 
  if (!itemId) {
    SpreadsheetApp.getActiveSpreadsheet().toast(`Идентификатор отсутствует, строка ${row}`, "Ошибка", 5);
    return;
  }
      
  const baseRow = findRowById(baseSheet, itemId);
  if (baseRow) {
    // Записываем примечание в столбце COL_NOTE (Примечание) (столбец G)
    baseSheet.getRange(baseRow, col).setValue(note);
  }
  
  SpreadsheetApp.getActiveSpreadsheet().toast("Примечание перенесено в базу", "Информация", 5);
}

/**
 * Возврат изделия на предыдущий этап
*/
function handleReturn(ss, sheet, range, row, clearColumn){
  //const note = val.toString().trim();  //Сохраняем значение примечания с листа роли
  range.setValue(false); //очищаем исходную ячейку от текста
  const itemId = sheet.getRange(row, 4).getValue().toString().trim();
  const baseSheet = ss.getSheetByName("База");
 
  if (!itemId) {
    SpreadsheetApp.getActiveSpreadsheet().toast(`Идентификатор отсутствует, строка ${row}`, "Ошибка", 5);
    return;
  }
      
  const baseRow = findRowById(baseSheet, itemId);
  if (baseRow) {
    // Устнавливаем значение False в переданном в функцию столбце
    baseSheet.getRange(baseRow, clearColumn).setValue(false);
  }
  
  SpreadsheetApp.getActiveSpreadsheet().toast("Отметка о готовности снята в столбце", "Информация", 5);
}

//заполнение чекбоксами столбцов, в которых должны ставиться отметки
function fillCheckboxes(ss, itemId){
  Logger.log("Запущена функция заполнения ячеек чекбоксами. Идентификатор изделия: " + itemId);  
  const baseSheet = ss.getSheetByName("База"); //Определяем лист с базой данных
  const baseRow = findRowById(baseSheet, itemId); //определяем строку в базе данных, в которой содержится информация об этом изделии
  Logger.log("Номер строки в базе = " + baseRow);
  const status = baseSheet.getRange(baseRow, COL_STATUS).getValue();
  Logger.log("Имя целевого листа = " + status);
  const targetSheet = ss.getSheetByName(status); //Определяем лист, куда в соответствии со статусом было перенесено изделие
  const settings = getSettingsForSheet(status);
  if (settings == null){Logger.log("настройки для листа " + status + "не найдены")} 
  const lastRow = getLastRow(targetSheet, COL_DATE); //определяем последнюю строку, заполненную значениями
  const range = targetSheet.getRange(2, settings.firstColumn,lastRow, settings.count);  //определяем диапазон, который должен быть заполнен чекбоксами
  range.setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build());  //записываем чекбоксы в диапазон
  range.setValue(false);  //устанавливаем, что в чекбоксах должно быть значение false
}

// Функция для получения настроек по имени листа
function getSettingsForSheet(targetSheetName) {
  
  // Массив настроек с объектами для каждого листа
  var sheetCheckboxSettings = [
    { sheetName: "Проектирование", firstColumn: 11, count: 1 }, //K
    { sheetName: "Развертывание спецификации", firstColumn: 11, count: 2 }, //K, L
    { sheetName: "Технолог", firstColumn: 11, count: 2},  //K, L
    { sheetName: "Проверка технологии", firstColumn: 11, count: 2 }, //K, L
    { sheetName: "Диспетчер- комплектовщик", firstColumn: 11, count: 2 }, //K, L
    { sheetName: "Комплектация", firstColumn: 11, count: 2 }, //K, L
    { sheetName: "Склад", firstColumn: 11, count: 2 }, //K, L
    { sheetName: "Склад производства", firstColumn: 11, count: 2 }, //K, L
    { sheetName: "Производство", firstColumn: 11, count: 2 }, //K, L
    { sheetName: "ОТК", firstColumn: 11, count: 2 }, //K, L
    { sheetName: "Приемка", firstColumn: 11, count: 2 }, //K, L
    { sheetName: "Готово", firstColumn: 11, count: 2 } //K, L
  ];

  for (var i = 0; i < sheetCheckboxSettings.length; i++) {
    if (sheetCheckboxSettings[i].sheetName === targetSheetName) {
      return sheetCheckboxSettings[i];
    }
  }
  return null; // если настройки для данного листа не найдены
}

//Определяет последнюю заполненную строку в столбце
function getLastRow(sheet, col) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 1) return 0;
  var values = sheet.getRange(1, col, lastRow, 1).getValues();
  for (var i = values.length - 1; i >= 0; i--) {
    if (values[i][0].toString().trim() !== "") {
      return i + 1; // строки индексируются с 1
    }
  }
  return 0; // если столбец пуст
}


/**
 * Формирует строку для записи в лист "База" для спецификации.
 *
 * @param {Date|null} date - Дата создания записи. Если null, используется текущая дата.
 * @param {string} name - Название изделия (для этой строки).
 * @param {number} fullQty - Полное количество этого изделия.
 * @param {string|null} parentName - Имя непосредственного родительского изделия (для поля "Назначение").
 * @param {number} futureIndex - Будущий номер строки в листе "База". Нужен для формул.
 * @param {string} parentUUIDForDestination - UUID непосредственного родительского изделия (который будет записан в поле "Назначение").
 * @param {boolean} opCheckboxStatus - Флаг, указывающий, нужно ли устанавливать чекбоксы технолога и проверки технологии.
 * @param {number} totalColumnCount - Общее количество столбцов, используемых для записи.
 * @param {string} rowUUID - UUID, который будет присвоен *этой* строке. Это может быть как новый сгенерированный, так и существующий.
 * @param {string} [note=''] - Примечание для этой строки. <-- НОВОЕ ИЗМЕНЕНИЕ
 * @returns {Array<any>} Сформированная строка-массив с данными и формулами.
 */
function createSpecRow(date, name, fullQty, parentName, futureIndex, parentUUIDForDestination, opCheckboxStatus, totalColumnCount, rowUUID, note = '') { 
  const row = buildBaseRowTemplate(futureIndex, totalColumnCount);

  // Фиксируем "текущее время" один раз, чтобы все метки были идентичны
  const now = new Date();

  // Дата записи: либо переданная, либо "сейчас"
  row[COL_DATE - 1] = (date == null) ? now : date;
  row[COL_ITEMNAME - 1] = name;

  // UUID строки — только если явно передан 
  if (rowUUID !== null && rowUUID !== undefined) {
    row[COL_UUID - 1] = rowUUID; 
  }

  if (parentName != null) {
    const cleanedParentName = String(parentName).replace(/<[^>]*>/g, '');
    const cleanedParentUUIDForDestination = String(parentUUIDForDestination).replace(/<[^>]*>/g, '');
    row[COL_DESTINATION - 1] = `${cleanedParentName} = ${fullQty} (${cleanedParentUUIDForDestination})`;
  } else {
    row[COL_DESTINATION - 1] = ""; 
  }

  row[COL_QUANTITY - 1] = fullQty;
  row[COL_NOTE - 1] = note;

  // Все служебные отметки — одним временем (now)
  row[COL_DESIGN - 1] = now;
  row[COL_TECHNICAL_CONTROL - 1] = now;     // <-- исправление опечатки
  row[COL_DEPLOYMENT - 1] = now;

  if (opCheckboxStatus === true) { 
    row[COL_TECHNOLOGIEST - 1] = now;
    row[COL_TECHNOLOGY_CHECK - 1] = now;
  }

  return row;
}




/**
 * Очищает строку назначения от GUID в скобках (например, "(21699c79-a345-443f-a5d6-7a4e299ab96b)").
 * Функция обрабатывает несколько таких вхождений в строке.
 *
 * @param {string} destinationString Исходная строка назначения, которая может содержать GUID в скобках.
 * @returns {string} Очищенная строка назначения без GUID в скобках.
 */
function normalizeDestination(destinationString) {
  // Регулярное выражение для поиска:
  // Пробел, затем открывающая скобка, затем GUID, затем закрывающая скобка.
  // GUID состоит из букв (a-fA-F), цифр (0-9) и дефисов (-).
  // \s - пробельный символ
  // \( - экранированная открывающая скобка
  // [0-9a-fA-F]{8} - 8 шестнадцатеричных символов
  // (?:-[0-9a-fA-F]{4}){3} - 3 группы по 4 шестнадцатеричных символа, разделённые дефисом (незахватывающая группа)
  // -[0-9a-fA-F]{12} - дефис и 12 шестнадцатеричных символов
  // \) - экранированная закрывающая скобка
  // g - глобальный флаг, чтобы найти все совпадения, а не только первое.
  const guidRegex = /\s\([0-9a-fA-F]{8}(?:-[0-9a-fA-F]{4}){3}-[0-9a-fA-F]{12}\)/g;

  // Заменяем все найденные совпадения пустой строкой.
  return destinationString.replace(guidRegex, '');
}


/**
 * Гарантирует, что начиная с templateRow (включительно) внизу листа
 * есть минимум minFreeRows свободных строк.
 *
 * Если строк не хватает — добавляет и копирует "шаблон" из templateRow
 * во все добавленные строки (формулы/формат/валидации).
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {number} templateRow       Строка-шаблон (обычно первая строка пустого блока/запаса)
 * @param {number} minFreeRows       Минимальное количество свободных строк (включая templateRow)
 * @returns {{added:number, freeRows:number, totalRows:number}}
 */
function ensureFreeRows(sheet, templateRow, minFreeRows = 100) {
  const FN = "ensureFreeRows";

  if (!sheet) throw new Error(`${FN}: sheet is required`);
  if (!Number.isInteger(templateRow) || templateRow < 1) {
    throw new Error(`${FN}: templateRow must be positive integer, got ${templateRow}`);
  }
  if (!Number.isInteger(minFreeRows) || minFreeRows < 0) {
    throw new Error(`${FN}: minFreeRows must be non-negative integer, got ${minFreeRows}`);
  }

  // 1) Убеждаемся, что строк как минимум достаточно, чтобы templateRow существовала
  let totalRows = sheet.getMaxRows();
  if (templateRow > totalRows) {
    const addToReachTemplate = templateRow - totalRows;
    sheet.insertRowsAfter(totalRows, addToReachTemplate);
    totalRows = sheet.getMaxRows();
    Logger.log(`[${FN}] Expanded sheet to reach templateRow=${templateRow}. Added=${addToReachTemplate}. TotalRows=${totalRows}`);
  }

  // 2) Сколько свободных строк есть снизу, начиная с templateRow ВКЛЮЧИТЕЛЬНО
  const freeRows = totalRows - templateRow + 1;

  if (freeRows >= minFreeRows) {
    Logger.log(`[${FN}] Sheet="${sheet.getName()}": enough free rows. freeRows=${freeRows}, required=${minFreeRows}, templateRow=${templateRow}`);
    return { added: 0, freeRows, totalRows };
  }

  const rowsToAdd = minFreeRows - freeRows;

  // 3) Добавляем строки в конец листа
  sheet.insertRowsAfter(totalRows, rowsToAdd);

  const lastCol = sheet.getLastColumn();
  if (lastCol === 0) {
    // Теоретически: пустой лист без колонок
    Logger.log(`[${FN}] Sheet="${sheet.getName()}": lastCol=0, rows added=${rowsToAdd}, but nothing copied.`);
    return { added: rowsToAdd, freeRows: freeRows + rowsToAdd, totalRows: totalRows + rowsToAdd };
  }

  // 4) Копируем шаблон (1 строка) на добавленные строки
  const sourceRange = sheet.getRange(templateRow, 1, 1, lastCol);
  const targetStartRow = totalRows + 1; // первая добавленная строка
  const targetRange = sheet.getRange(targetStartRow, 1, rowsToAdd, lastCol);

  // Копирует формулы/формат/валидации и пр. (как было)
  sourceRange.copyTo(targetRange, { contentsOnly: false });

  Logger.log(`[${FN}] Sheet="${sheet.getName()}": addedRows=${rowsToAdd}, newTotalRows=${totalRows + rowsToAdd}, freeRowsNow=${(totalRows + rowsToAdd) - templateRow + 1}, templateRow=${templateRow}`);

  return { added: rowsToAdd, freeRows: (totalRows + rowsToAdd) - templateRow + 1, totalRows: totalRows + rowsToAdd };
}
