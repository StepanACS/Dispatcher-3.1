/**
 * Основной обработчик редактирования.
 *
 * onEdit выполняет только короткую операцию постановки события в очередь. Очередь
 * хранится в DocumentProperties отдельными записями, поэтому одно событие не может
 * превысить лимит свойства из-за роста общего JSON-массива.
 *
 * DocumentLock защищает только короткие операции с очередью. Длительная бизнес-
 * логика выполняется под ScriptLock, поэтому новые onEdit могут продолжать ставить
 * события в очередь, пока предыдущие события обрабатываются.
 */

const EDIT_QUEUE_CFG = {
  QUEUE_PREFIX: 'EDIT_QUEUE_V2_',
  MAX_QUEUE_LEN: 300,
  ENQUEUE_LOCK_WAIT_MS: 5000,
  QUEUE_LOCK_WAIT_MS: 5000,
  PROCESS_LOCK_WAIT_MS: 50,
  DRAIN_TIME_LIMIT_MS: 240000
};

const _sheetToBaseColumnMaps = new Map();

function onEditHandler(e) {
  if (!e || !e.range) {
    Logger.log('onEditHandler: событие отсутствует или не содержит range.');
    return;
  }

  try {
    const range = e.range;
    const sheet = range.getSheet();
    const spreadsheet = e.source || sheet.getParent();
    const rawValue = (typeof e.value !== 'undefined') ? e.value : range.getValue();

    const payload = {
      ts: Date.now(),
      id: Utilities.getUuid(),
      spreadsheetId: spreadsheet.getId(),
      sheetName: sheet.getName(),
      a1Notation: range.getA1Notation(),
      row: range.getRow(),
      column: range.getColumn(),
      value: rawValue
    };

    if (!enqueueEditEvent_(payload)) {
      Logger.log(`onEditHandler: не удалось поставить событие ${payload.id} в очередь.`);
      return;
    }

    tryProcessEditQueue_();
  } catch (error) {
    Logger.log('onEditHandler: ' + (error && error.stack ? error.stack : error));
  }
}

function enqueueEditEvent_(payload) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(EDIT_QUEUE_CFG.ENQUEUE_LOCK_WAIT_MS)) {
    return false;
  }

  try {
    const props = PropertiesService.getDocumentProperties();
    const key = buildQueueKey_(payload);
    props.setProperty(key, JSON.stringify(payload));

    const queueKeys = Object.keys(props.getProperties())
      .filter(keyName => keyName.indexOf(EDIT_QUEUE_CFG.QUEUE_PREFIX) === 0)
      .sort();

    const overflow = queueKeys.length - EDIT_QUEUE_CFG.MAX_QUEUE_LEN;
    if (overflow > 0) {
      for (let i = 0; i < overflow; i++) {
        props.deleteProperty(queueKeys[i]);
      }
      Logger.log(`[QUEUE] удалено ${overflow} самых старых событий из переполненной очереди.`);
    }

    return true;
  } finally {
    lock.releaseLock();
  }
}

function buildQueueKey_(payload) {
  const timestamp = String(payload.ts || Date.now()).padStart(13, '0');
  return EDIT_QUEUE_CFG.QUEUE_PREFIX + timestamp + '_' + payload.id;
}

function tryProcessEditQueue_() {
  const processorLock = LockService.getScriptLock();
  if (!processorLock.tryLock(EDIT_QUEUE_CFG.PROCESS_LOCK_WAIT_MS)) {
    return;
  }

  const startedAt = Date.now();
  try {
    while (Date.now() - startedAt <= EDIT_QUEUE_CFG.DRAIN_TIME_LIMIT_MS) {
      const next = dequeueOneEvent_();
      if (!next) return;

      try {
        mainEditLogic(next);
      } catch (error) {
        Logger.log(`[QUEUE] событие ${next.id || 'unknown'} завершилось ошибкой: ` +
          (error && error.stack ? error.stack : error));
      }
    }

    Logger.log('[QUEUE] достигнут лимит времени; оставшиеся события обработает следующий запуск.');
  } finally {
    processorLock.releaseLock();
  }
}

function dequeueOneEvent_() {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(EDIT_QUEUE_CFG.QUEUE_LOCK_WAIT_MS)) {
    return null;
  }

  try {
    const props = PropertiesService.getDocumentProperties();
    const allProps = props.getProperties();
    const queueKeys = Object.keys(allProps)
      .filter(key => key.indexOf(EDIT_QUEUE_CFG.QUEUE_PREFIX) === 0)
      .sort();

    while (queueKeys.length > 0) {
      const key = queueKeys.shift();
      const raw = allProps[key];
      props.deleteProperty(key);

      try {
        return JSON.parse(raw);
      } catch (error) {
        Logger.log(`[QUEUE] повреждённое событие ${key} удалено: ${error.message}`);
      }
    }

    return null;
  } finally {
    lock.releaseLock();
  }
}

function isTrueEditValue_(value) {
  return value === true || String(value).toUpperCase() === 'TRUE';
}

function mainEditLogic(eventData) {
  if (!eventData || !eventData.spreadsheetId || !eventData.sheetName) {
    throw new Error('mainEditLogic: неполные данные события.');
  }

  const ss = SpreadsheetApp.openById(eventData.spreadsheetId);
  const sheetName = eventData.sheetName;
  const row = eventData.row;
  const col = eventData.column;
  const val = eventData.value;
  const isTrue = isTrueEditValue_(val);

  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    Logger.log(`mainEditLogic: лист «${sheetName}» не найден.`);
    return;
  }

  const range = sheet.getRange(row, col);

  switch (sheetName) {
    case 'Запуск': {
      if (col !== COL_LAUNCH_TRIGGER || !isTrue) break;

      const launchData = sheet.getRange(row, 1, 1, 3).getValues()[0];
      const itemName = String(launchData[0] ?? '').trim();
      const quantity = String(launchData[2] ?? '').trim();

      if (!itemName || !quantity) {
        ss.toast('Недостаточно данных для внесения изделия в базу данных Диспетчера', 'Ошибка', 5);
        range.setValue(false);
        return;
      }

      handleLaunchEdit(ss, sheet, row);
      ss.toast('Изделие добавлено в базу', 'Информация', 5);
      break;
    }

    case 'Проектирование':
      if (col === COL_ITEMNAME || col === COL_NOTE || col === COL_QUANTITY) {
        handleNote(ss, sheet, range, row, col, val);
      }
      if (col === COL_CHECKBOX && isTrue) {
        const itemId = handleNewSheetEdit(ss, sheet, row, col, val);
        if (itemId) fillCheckboxes(ss, itemId);
        range.setValue(false);
      }
      break;

    case 'Техконтроль':
      if (col === COL_ITEMNAME || col === COL_NOTE || col === COL_QUANTITY) {
        handleNote(ss, sheet, range, row, col, val);
      }
      if (col === COL_CHECKBOX && isTrue) {
        const itemId = handleNewSheetEdit(ss, sheet, row, col, val);
        if (itemId) fillCheckboxes(ss, itemId);
        range.setValue(false);
      }
      if (col === COL_RETURN_CHECKBOX && isTrue) {
        handleReturn(ss, sheet, range, row, COL_DESIGN);
      }
      break;

    case 'Развертывание спецификации':
      if (col === COL_ITEMNAME || col === COL_NOTE || col === COL_QUANTITY) {
        handleNote(ss, sheet, range, row, col, val);
      }
      if (col === COL_CHECKBOX && isTrue) {
        const rowData = sheet.getRange(row, COL_UUID, 1, COL_QUANTITY - COL_UUID + 1).getValues()[0];
        const itemUUID = String(rowData[0] ?? '').trim();
        const itemName = String(rowData[COL_ITEMNAME - COL_UUID] ?? '').trim();
        const initialQuantity = rowData[COL_QUANTITY - COL_UUID];
        const baseSheet = ss.getSheetByName('База');

        if (!baseSheet) throw new Error('Лист «База» не найден.');
        handleDeployment(itemUUID, itemName, ss, baseSheet, row, initialQuantity);
        range.setValue(false);
      }
      if (col === COL_RETURN_CHECKBOX && isTrue) {
        handleReturn(ss, sheet, range, row, COL_TECHNICAL_CONTROL);
      }
      break;

    case 'Технолог':
      if (col === COL_ITEMNAME || col === COL_NOTE || col === COL_QUANTITY) {
        handleNote(ss, sheet, range, row, col, val);
      }
      if (col === COL_CHECKBOX && isTrue) {
        const itemId = handleNewSheetEdit(ss, sheet, row, col, val);
        if (itemId) fillCheckboxes(ss, itemId);
        range.setValue(false);
      }
      if (col === COL_RETURN_CHECKBOX && isTrue) {
        handleReturn(ss, sheet, range, row, COL_DEPLOYMENT);
      }
      break;

    case 'Проверка технологии':
      if (col === COL_NOTE) {
        handleNote(ss, sheet, range, row, col, val);
      }
      if (col === COL_CHECKBOX && isTrue) {
        const itemId = handleNewSheetEdit(ss, sheet, row, col, val);
        if (itemId) fillCheckboxes(ss, itemId);
        range.setValue(false);
      }
      if (col === COL_RETURN_CHECKBOX && isTrue) {
        handleReturn(ss, sheet, range, row, COL_TECHNOLOGIEST);
      }
      break;

    case 'Диспетчер- комплектовщик':
      if (col === COL_ITEMNAME || col === COL_NOTE || col === COL_QUANTITY) {
        handleNote(ss, sheet, range, row, col, val);
      }
      if (col === COL_CHECKBOX && isTrue) {
        const itemId = handleSupplyOrder(ss, sheet, range, row);
        if (itemId) fillCheckboxes(ss, itemId);
        range.setValue(false);
      }
      if (col === COL_RETURN_CHECKBOX && isTrue) {
        handleReturn(ss, sheet, range, row, COL_TECHNOLOGY_CHECK);
      }
      break;

    case 'Комплектация':
      if (col === COL_ITEMNAME || col === COL_NOTE || col === COL_QUANTITY || col === COL_DESTINATION) {
        handleNote(ss, sheet, range, row, col, val);
      }
      if (col === COL_CHECKBOX && isTrue) {
        const itemId = handlePickingOrder(ss, sheet, range, row);
        if (itemId) fillCheckboxes(ss, itemId);
        range.setValue(false);
      }
      if (col === COL_RETURN_CHECKBOX && isTrue) {
        handleReturn(ss, sheet, range, row, COL_PICKER);
      }
      break;

    case 'Склад':
      if (col === COL_NOTE) {
        handleNote(ss, sheet, range, row, col, val);
      }
      if (col === COL_CHECKBOX && isTrue) {
        const itemId = handleStock(ss, sheet, range, row);
        if (itemId) fillCheckboxes(ss, itemId);
        range.setValue(false);
      }
      if (col === COL_RETURN_CHECKBOX && isTrue) {
        handleReturn(ss, sheet, range, row, COL_PICKING);
      }
      break;

    case 'Склад производства':
      if (col === COL_NOTE) {
        handleNote(ss, sheet, range, row, col, val);
      }
      if (col === COL_CHECKBOX && isTrue) {
        const itemId = handleProductionStock(ss, sheet, range, row);
        if (itemId) fillCheckboxes(ss, itemId);
        range.setValue(false);
      }
      if (col === COL_RETURN_CHECKBOX && isTrue) {
        handleReturn(ss, sheet, range, row, COL_STOCK);
      }
      break;

    case 'Производство':
      if (col === COL_NOTE) {
        handleNote(ss, sheet, range, row, col, val);
      }
      if (col === COL_CHECKBOX && isTrue) {
        const itemId = handleProduction(ss, sheet, range, row);
        if (itemId) fillCheckboxes(ss, itemId);
        range.setValue(false);
      }
      if (col === COL_RETURN_CHECKBOX && isTrue) {
        handleReturn(ss, sheet, range, row, COL_PRODUCTION_STOCK);
      }
      break;

    case 'ОТК':
      if (col === COL_NOTE) {
        handleNote(ss, sheet, range, row, col, val);
      }
      if (col === COL_CHECKBOX && isTrue) {
        const itemId = handleNewSheetEdit(ss, sheet, row, col, val);
        if (itemId) fillCheckboxes(ss, itemId);
        range.setValue(false);
      }
      if (col === COL_RETURN_CHECKBOX && isTrue) {
        handleReturn(ss, sheet, range, row, COL_PRODUCTION);
      }
      break;

    case 'Приемка':
      if (col === COL_NOTE || col === COL_QUANTITY) {
        handleNote(ss, sheet, range, row, col, val);
      }
      if (col === COL_CHECKBOX && isTrue) {
        const itemId = handleAcceptance(ss, sheet, range, row);
        if (itemId) fillCheckboxes(ss, itemId);
        range.setValue(false);
      }
      if (col === COL_RETURN_CHECKBOX && isTrue) {
        handleReturn(ss, sheet, range, row, COL_QUALITY_CHECK);
      }
      break;

    case 'Готово':
      if (col === COL_NOTE) {
        handleNote(ss, sheet, range, row, col, val);
      }
      if (col === COL_CHECKBOX && isTrue) {
        const itemId = handleReady(ss, sheet, range, row);
        if (itemId) fillCheckboxes(ss, itemId);
        range.setValue(false);
      }
      if (col === COL_RETURN_CHECKBOX && isTrue) {
        handleReturn(ss, sheet, range, row, COL_ACCEPTANCE);
      }
      break;

    case 'База':
      handleBaseEdit_(ss, sheet, range, row, col, val, isTrue);
      break;

    default:
      handleNewSheetEdit(ss, sheet, row, col, val);
      break;
  }
}

function handleBaseEdit_(ss, sheet, range, row, col, val, isTrue) {
  switch (col) {
    case COL_DEPLOYMENT: {
      const booleanVal = (typeof val === 'string') ? val.toUpperCase() === 'TRUE' : Boolean(val);
      const rowData = sheet.getRange(row, COL_UUID, 1, COL_QUANTITY - COL_UUID + 1).getValues()[0];
      const itemUUID = String(rowData[0] ?? '').trim();
      const itemName = String(rowData[COL_ITEMNAME - COL_UUID] ?? '').trim();
      const initialQuantity = rowData[COL_QUANTITY - COL_UUID];

      if (booleanVal) {
        handleDeployment(itemUUID, itemName, ss, sheet, row, initialQuantity);
      } else {
        handleUndeployment(ss, sheet, row);
      }
      break;
    }

    case COL_PICKER:
      if (isTrue) {
        handleSupplyOrder(ss, sheet, range, row);
        range.setValue(false);
      }
      break;

    case COL_STOCK:
      if (isTrue) {
        handleStock(ss, sheet, range, row);
        range.setValue(false);
      }
      break;

    case COL_PRODUCTION_STOCK:
      if (isTrue) {
        handleProductionStock(ss, sheet, range, row);
        range.setValue(false);
      }
      break;

    case COL_ACCEPTANCE:
      if (isTrue) {
        ss.toast('Для выполнения приемки перейдите на соответствующий лист', 'Ошибка', 5);
        range.setValue(false);
      }
      break;

    case COL_READY:
      if (isTrue) {
        handleReady(ss, sheet, range, row);
        range.setValue(false);
      }
      break;

    default:
      if (isTrue) checkboxOk(range);
  }
}

/**
 * Универсальный обработчик отметки на листе роли.
 * Настройки листов читаются один раз за execution и повторно используются всеми
 * событиями, которые процессор очереди успевает обработать за один запуск.
 */
function handleNewSheetEdit(ss, sheet, row, col, val) {
  if (col !== COL_CHECKBOX || !isTrueEditValue_(val)) return null;

  const sheetName = sheet.getName();
  const sheetToBaseColumn = getSheetToBaseColumnMap_(ss);
  const baseColumn = sheetToBaseColumn.get(sheetName);
  if (!baseColumn) return null;

  const rowData = sheet.getRange(row, COL_UUID, 1, COL_ITEMNAME - COL_UUID + 1).getValues()[0];
  const itemId = String(rowData[0] ?? '').trim();
  const itemName = String(rowData[COL_ITEMNAME - COL_UUID] ?? '').trim();

  if (!itemId) {
    sheet.getRange(row, COL_CHECKBOX).setValue(false);
    Logger.log(`handleNewSheetEdit: отсутствует UUID для «${itemName}», строка ${row}.`);
    return null;
  }

  const baseSheet = ss.getSheetByName('База');
  if (!baseSheet) throw new Error('Лист «База» не найден.');

  const baseRow = findRowById(baseSheet, itemId);
  if (!baseRow) {
    Logger.log(`handleNewSheetEdit: UUID ${itemId} не найден в «База».`);
    return null;
  }

  checkboxOk(baseSheet.getRange(baseRow, baseColumn));
  sheet.getRange(row, COL_CHECKBOX).setValue(false);
  return itemId;
}

function getSheetToBaseColumnMap_(ss) {
  const spreadsheetId = ss.getId();
  if (_sheetToBaseColumnMaps.has(spreadsheetId)) {
    return _sheetToBaseColumnMaps.get(spreadsheetId);
  }

  const result = new Map();
  const settingsSheet = ss.getSheetByName('Настройки');
  if (!settingsSheet) {
    _sheetToBaseColumnMaps.set(spreadsheetId, result);
    return result;
  }

  const lastRow = settingsSheet.getLastRow();
  if (lastRow >= 2) {
    const settingsData = settingsSheet.getRange(2, 1, lastRow - 1, 2).getValues();
    for (const row of settingsData) {
      const currentSheetName = String(row[0] ?? '').trim();
      const baseColumnLetter = String(row[1] ?? '').trim().toUpperCase();
      if (currentSheetName && baseColumnLetter) {
        result.set(currentSheetName, columnLetterToNumber(baseColumnLetter));
      }
    }
  }

  _sheetToBaseColumnMaps.set(spreadsheetId, result);
  return result;
}
