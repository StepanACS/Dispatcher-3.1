/**
 * Обрабатывает готовность изделия: отмечает в "База", закрывает Trello,
 * архивирует строку и дописывает в архив цену из внешней базы.
 */
function handleReady(ss, sheet, range, row) {
  const baseSheet    = ss.getSheetByName("База");
  const archiveSs    = SpreadsheetApp.openById("1oPsmDHUGmN0ZmzYbEMp137Px_KsxDcZRx4PhceUp-3g");
  const archiveSheet = archiveSs.getSheetByName("Архив");
  const itemId       = String(sheet.getRange(row, COL_UUID).getValue()).trim();
  const itemName     = String(sheet.getRange(row, 5).getValue()).trim();
  
  // 1) Находим строку в "База" по UUID
  const baseRow = findRowById(baseSheet, itemId);
  if (!baseRow) {
    Logger.log(`❌ itemId "${itemId}" not found in sheet База`);
    return;
  }
  
  // 2) Ставим галочку "выполнено" в БАЗА
  checkboxOk(baseSheet.getRange(baseRow, COL_ACCEPTANCE));
  checkboxOk(baseSheet.getRange(baseRow, COL_READY));
  
  
  // 3) Читаем всю строку из БАЗА
  const rowData  = baseSheet
    .getRange(baseRow, 1, 1, COL_FACT_QTY)
    .getValues();
  const trelloId = rowData[0][COL_TRELLO_ID - 1];
  Logger.log(`TrelloID = ${trelloId}`);
  closeTrelloCard(trelloId);
  
  // 4) Архивируем: дописываем строку сразу после последней
  const archiveLastRow = archiveSheet.getLastRow();
  archiveSheet
    .getRange(archiveLastRow + 1, 1, 1, COL_FACT_QTY)
    .setValues(rowData);
  
  // 5) Дописываем цену из внешней БД в колонку COL_FACT_QTY+1
  fetchAndAppendPrice(itemName, archiveSheet, archiveLastRow, COL_FACT_QTY);
  
  // 6) Удаляем исходную строку из "База"
  baseSheet.deleteRow(baseRow);
  Logger.log(`✅ Row ${baseRow} moved from "База" to "Архив"`);
  
  // 7) Сбрасываем чекбокс, если отметка ставилась не на листе "База"
  if (sheet.getName() !== baseSheet.getName()) {
    range.setValue(false);
  }
  
  return itemId;
}


/**
 * Ищет в внешней таблице цену для изделия и дописывает её в заданый архивный лист.
 * Если изделие не найдено, записывает цену 0.
 *
 * @param {string} itemName          — название изделия для поиска
 * @param {Sheet}  archiveSheet      — лист "Архив", куда дописываем цену
 * @param {number} archiveLastRow    — строка, после которой добавлена сама запись
 * @param {number} COL_FACT_QTY      — число колонок данных (без цены)
 */
function fetchAndAppendPrice(itemName, archiveSheet, archiveLastRow, COL_FACT_QTY) {
  const ss      = SpreadsheetApp.openById("1wR3wybH7UplH0ogrrELf1S27N9O5t5DS9lIkBHeagJ8");
  const dbSheet = ss.getSheetByName("БД артикулы");
  if (!dbSheet) {
    throw new Error('Sheet "БД артикулы" not found');
  }
  
  // Читаем A:C
  const lastRow = dbSheet.getLastRow();
  const data    = dbSheet.getRange(1, 1, lastRow, 3).getValues();
  
  // Ищем совпадение по названию
  let foundPrice = 0;
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]).trim() === itemName) {
      foundPrice = data[i][2];  // цена в колонке C
      break;
    }
  }
  if (foundPrice === 0) {
    Logger.log(`⚠️ Item "${itemName}" not found, defaulting price to 0`);
  } else {
    Logger.log(`Price ${foundPrice} found for "${itemName}"`);
  }
  
  // Записываем цену в следующую колонку после COL_FACT_QTY
  archiveSheet
    .getRange(archiveLastRow + 1, COL_FACT_QTY + 1)
    .setValue(foundPrice);
}



