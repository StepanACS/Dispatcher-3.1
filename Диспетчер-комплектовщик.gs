/** Обрабатывает завершение этапа «Диспетчер- комплектовщик». */
function handleSupplyOrder(ss, sheet, range, row) {
  const baseSheet = ss.getSheetByName(BASE_SHEET_NAME);
  const roleData = sheet.getRange(row, ROLE_COL_UUID, 1, ROLE_COL_ITEMNAME - ROLE_COL_UUID + 1).getValues()[0];
  const itemId = String(roleData[0] || '').trim();
  const itemName = String(roleData[ROLE_COL_ITEMNAME - ROLE_COL_UUID] || '').trim();
  const baseRow = findBaseRowByUuid_(baseSheet, itemId);
  if (!baseRow) {
    Logger.log(`handleSupplyOrder: UUID ${itemId} не найден в Базе.`);
    return null;
  }

  const data = readBaseFields_(baseSheet, baseRow, [
    'destination', 'quantity', 'deficit', 'production_mode', 'stage_picker'
  ]);
  // Повторное событие после уже выполненного этапа не должно создавать вторую заявку.
  if (data.stage_picker) {
    if (sheet.getSheetId() !== baseSheet.getSheetId()) range.setValue(false);
    return itemId;
  }

  if (data.deficit > 0 && data.production_mode === 'В снабжение') {
    setSupplyOrder(ss, itemId, itemName, data.quantity, String(data.destination || '').trim());
  }
  if (sheet.getSheetId() !== baseSheet.getSheetId()) range.setValue(false);
  checkboxOk(baseSheet.getRange(baseRow, baseCol_(ss, 'stage_picker')));
  return itemId;
}

function supplyRequestToken_(itemId) {
  return `[D31:${String(itemId || '').trim()}]`;
}

/**
 * Добавляет заявку в реальное «Снабжение». В TEST-среде запись не отключается:
 * она помечается [TEST Dispatcher 3.1] и UUID исходной строки Диспетчера.
 * Перед добавлением проверяем существующую ещё не полученную заявку с тем же token.
 */
function setSupplyOrder(ss, itemId, name, qty, purpose) {
  const supplySpreadsheet = SpreadsheetApp.openById(SUPPLY_SPREADSHEET_ID);
  const supplySheet = supplySpreadsheet.getSheetByName('Сводный');
  if (!supplySheet) throw new Error('В документе снабжения не найден лист «Сводный».');

  const token = supplyRequestToken_(itemId);
  const source = markExternalTestText_(ss, 'Script - Диспетчер 3.1');
  const comment = markExternalTestText_(ss, `${token} ${String(purpose || '').trim()}`.trim());

  const lastRow = supplySheet.getLastRow();
  if (lastRow >= 2) {
    // H = комментарий/назначение, M = получено. Повторяем только если предыдущая
    // заявка с этим UUID уже закрыта получением.
    const comments = supplySheet.getRange(2, 8, lastRow - 1, 1).getDisplayValues().flat();
    const received = supplySheet.getRange(2, 13, lastRow - 1, 1).getValues().flat();
    for (let i = comments.length - 1; i >= 0; i--) {
      if (!String(comments[i] || '').includes(token)) continue;
      const receivedValue = received[i];
      const isReceived = receivedValue instanceof Date ||
        (typeof receivedValue === 'number' && isFinite(receivedValue) && receivedValue > 0);
      if (!isReceived) return { created: false, row: i + 2, duplicate: true };
    }
  }

  const newOrderRow = getLastRowNumber(supplySheet, 2) + 1;
  supplySheet.getRange(newOrderRow, 1, 1, 8).setValues([[
    new Date(), source, qty, name, '', '', '', comment
  ]]);
  return { created: true, row: newOrderRow, duplicate: false };
}

function getLastRowNumber(sheet, column) {
  const maxRows = sheet.getMaxRows();
  if (maxRows < 1) return 0;
  return sheet.getRange(maxRows, column).getNextDataCell(SpreadsheetApp.Direction.UP).getRow();
}

function getSpec(itemName, link) {
  const sheetId = extractSheetId(link);
  if (!sheetId) throw new Error(`Не удалось извлечь ID таблицы из ссылки для «${itemName}».`);
  const targetSS = SpreadsheetApp.openById(sheetId);
  const specSheet = targetSS.getSheetByName(itemName);
  if (!specSheet) throw new Error(`Лист «${itemName}» не найден в таблице спецификации.`);
  const firstDataRow = 5;
  const lastRow = specSheet.getLastRow();
  if (lastRow < firstDataRow) return [];
  const raw = specSheet.getRange(firstDataRow, 2, lastRow - firstDataRow + 1, 3).getValues();
  const totalIndex = raw.findIndex(r => String(r[0] || '').trim().toLowerCase() === 'итого');
  return totalIndex === -1 ? [] : raw.slice(0, totalIndex);
}

function loadSdIdMap_(ss) {
  const result = new Map();
  const sdSheet = ss.getSheetByName('SD ID');
  if (!sdSheet) return result;
  const lastRow = sdSheet.getLastRow();
  if (lastRow < 1) return result;
  const data = sdSheet.getRange(1, 1, lastRow, 2).getValues();
  for (const row of data) {
    const name = String(row[1] || '').trim();
    if (name) result.set(name, String(row[0] || '').trim());
  }
  return result;
}

function getSdId(ss, name, sdIdMap) {
  const map = sdIdMap || loadSdIdMap_(ss);
  return map.get(String(name || '').trim()) || null;
}
