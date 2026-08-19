/**
 * Основной обработчик редактирования.
 *
 * onEdit выполняет короткую постановку события в FIFO-очередь. DocumentLock
 * защищает только head/tail и одну запись очереди; длительная бизнес-логика
 * выполняется без DocumentLock, поэтому параллельные редактирования могут быстро
 * добавлять новые события.
 */

const EDIT_QUEUE_CFG = {
  EVENT_PREFIX: 'EDIT_QUEUE_V2_EVENT_',
  HEAD_KEY: 'EDIT_QUEUE_V2_HEAD',
  TAIL_KEY: 'EDIT_QUEUE_V2_TAIL',
  MAX_QUEUE_LEN: 300,
  ENQUEUE_LOCK_WAIT_MS: 5000,
  QUEUE_LOCK_WAIT_MS: 5000,
  PROCESS_LOCK_WAIT_MS: 500,
  DRAIN_TIME_LIMIT_MS: 240000
};

/**
 * Настройки системных уведомлений. Отрицательный timeout оставляет toast на
 * экране до тех пор, пока его не заменит следующий статус или пользователь не
 * закроет уведомление вручную.
 */
const STATUS_TOAST_CFG = Object.freeze({
  TITLE: 'Диспетчер',
  QUEUED_TIMEOUT: -1,
  WORKING_TIMEOUT: -1,
  SUCCESS_TIMEOUT: 5,
  NOT_DONE_TIMEOUT: 8,
  ERROR_TIMEOUT: 10
});

const _sheetToBaseColumnMaps = new Map();
let _roleSheetConfigCache = null;

/**
 * Возвращает конфигурацию однотипных листов исполнителей.
 *
 * Конфигурация создаётся лениво, а не на верхнем уровне файла. Это важно для
 * Apps Script: порядок инициализации файлов проекта не гарантирован, поэтому
 * верхнеуровневое обращение из Main.gs к const COL_* из Config.gs может дать
 * ReferenceError ещё до запуска выбранной пользователем функции.
 */
function getRoleSheetConfig_() {
  if (_roleSheetConfigCache) return _roleSheetConfigCache;

  _roleSheetConfigCache = Object.freeze({
    'Проектирование': {
      noteColumns: [COL_ITEMNAME, COL_NOTE, COL_QUANTITY],
      completion: 'generic'
    },
    'Техконтроль': {
      noteColumns: [COL_ITEMNAME, COL_NOTE, COL_QUANTITY],
      completion: 'generic',
      returnColumn: COL_DESIGN
    },
    'Технолог': {
      noteColumns: [COL_ITEMNAME, COL_NOTE, COL_QUANTITY],
      completion: 'generic',
      returnColumn: COL_DEPLOYMENT
    },
    'Проверка технологии': {
      noteColumns: [COL_NOTE],
      completion: 'generic',
      returnColumn: COL_TECHNOLOGIEST
    },
    'Диспетчер- комплектовщик': {
      noteColumns: [COL_ITEMNAME, COL_NOTE, COL_QUANTITY],
      completion: 'supply',
      returnColumn: COL_TECHNOLOGY_CHECK
    },
    'Комплектация': {
      noteColumns: [COL_ITEMNAME, COL_NOTE, COL_QUANTITY, COL_DESTINATION],
      completion: 'picking',
      returnColumn: COL_PICKER
    },
    'Склад': {
      noteColumns: [COL_NOTE],
      completion: 'stock',
      returnColumn: COL_PICKING
    },
    'Склад производства': {
      noteColumns: [COL_NOTE],
      completion: 'productionStock',
      returnColumn: COL_STOCK
    },
    'Производство': {
      noteColumns: [COL_NOTE],
      completion: 'production',
      returnColumn: COL_PRODUCTION_STOCK
    },
    'ОТК': {
      noteColumns: [COL_NOTE],
      completion: 'generic',
      returnColumn: COL_PRODUCTION
    },
    'Приемка': {
      noteColumns: [COL_NOTE, COL_QUANTITY],
      completion: 'acceptance',
      returnColumn: COL_QUALITY_CHECK
    },
    'Готово': {
      noteColumns: [COL_NOTE],
      completion: 'ready',
      returnColumn: COL_ACCEPTANCE,
      fillNextStage: false
    }
  });

  return _roleSheetConfigCache;
}

/** Возвращает название длительного/чекбоксного действия или пустую строку. */
function getTrackedActionName_(sheetName, col, value) {
  const isTrue = isTrueEditValue_(value);

  if (sheetName === 'Запуск' && col === COL_LAUNCH_TRIGGER && isTrue) {
    return 'Запуск изделия';
  }

  if (sheetName === 'Развертывание спецификации') {
    if (col === COL_CHECKBOX && isTrue) return 'Развертывание спецификации';
    if (col === COL_RETURN_CHECKBOX && isTrue) return 'Возврат на предыдущий этап';
    return '';
  }

  if (sheetName === 'База') {
    if (col === COL_DEPLOYMENT) {
      return isTrue ? 'Развертывание спецификации' : 'Сворачивание спецификации';
    }

    if (isTrue && col >= COL_DESIGN && col <= COL_READY) {
      switch (col) {
        case COL_DESIGN: return 'Проектирование';
        case COL_TECHNICAL_CONTROL: return 'Техконтроль';
        case COL_TECHNOLOGIEST: return 'Технолог';
        case COL_TECHNOLOGY_CHECK: return 'Проверка технологии';
        case COL_PICKER: return 'Диспетчер-комплектовщик';
        case COL_PICKING: return 'Комплектация';
        case COL_STOCK: return 'Склад';
        case COL_PRODUCTION_STOCK: return 'Склад производства';
        case COL_PRODUCTION: return 'Производство';
        case COL_QUALITY_CHECK: return 'ОТК';
        case COL_ACCEPTANCE: return 'Приемка';
        case COL_READY: return 'Готово';
      }
    }
    return '';
  }

  const roleConfig = getRoleSheetConfig_()[sheetName];
  if (roleConfig && isTrue) {
    if (col === COL_CHECKBOX) return `Этап «${sheetName}»`;
    if (col === COL_RETURN_CHECKBOX && roleConfig.returnColumn) {
      return `Возврат из «${sheetName}»`;
    }
  }

  // Поддерживает дополнительные однотипные листы, заданные через «Настройки».
  if (isTrue && col === COL_CHECKBOX) return `Этап «${sheetName}»`;
  return '';
}

/** Получает имя изделия только для отслеживаемых операций. */
function getStatusItemName_(sheet, row, sheetName) {
  try {
    const itemColumn = sheetName === 'Запуск' ? 1 : COL_ITEMNAME;
    return String(sheet.getRange(row, itemColumn).getDisplayValue() || '').trim();
  } catch (error) {
    Logger.log(`getStatusItemName_: ${error.message}`);
    return '';
  }
}

function buildStatusLabel_(actionName, itemName) {
  if (!actionName) return '';
  return itemName ? `${actionName}: «${itemName}»` : actionName;
}

/** Единая точка отображения состояния длительной операции. */
function showActionStatus_(ss, state, label, details) {
  if (!ss || !label) return;

  let message;
  let timeout;
  switch (state) {
    case 'queued':
      message = `⏳ В очереди: ${label}`;
      timeout = STATUS_TOAST_CFG.QUEUED_TIMEOUT;
      break;
    case 'working':
      message = `⚙️ Выполняется: ${label}`;
      timeout = STATUS_TOAST_CFG.WORKING_TIMEOUT;
      break;
    case 'success':
      message = `✅ Завершено: ${label}`;
      timeout = STATUS_TOAST_CFG.SUCCESS_TIMEOUT;
      break;
    case 'notDone':
      message = `⚠️ Операция не выполнена: ${label}`;
      timeout = STATUS_TOAST_CFG.NOT_DONE_TIMEOUT;
      break;
    case 'error':
      message = `❌ Ошибка: ${label}`;
      if (details) message += `\n${details}`;
      timeout = STATUS_TOAST_CFG.ERROR_TIMEOUT;
      break;
    default:
      return;
  }

  try {
    ss.toast(message, STATUS_TOAST_CFG.TITLE, timeout);
  } catch (error) {
    Logger.log(`showActionStatus_: ${error.message}`);
  }
}

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
    const actionName = getTrackedActionName_(sheet.getName(), range.getColumn(), rawValue);
    const itemName = actionName ? getStatusItemName_(sheet, range.getRow(), sheet.getName()) : '';

    const payload = {
      ts: Date.now(),
      id: Utilities.getUuid(),
      spreadsheetId: spreadsheet.getId(),
      sheetName: sheet.getName(),
      a1Notation: range.getA1Notation(),
      row: range.getRow(),
      column: range.getColumn(),
      value: rawValue,
      trackStatus: Boolean(actionName),
      statusLabel: buildStatusLabel_(actionName, itemName)
    };

    if (!enqueueEditEvent_(payload)) {
      Logger.log(`onEditHandler: не удалось поставить событие ${payload.id} в очередь.`);
      if (payload.trackStatus) {
        showActionStatus_(spreadsheet, 'error', payload.statusLabel, 'Не удалось поставить задачу в очередь.');
      }
      return;
    }

    if (payload.trackStatus) {
      showActionStatus_(spreadsheet, 'queued', payload.statusLabel);
    }

    tryProcessEditQueue_();
  } catch (error) {
    Logger.log('onEditHandler: ' + (error && error.stack ? error.stack : error));
  }
}

function queuePosition_(props, key, fallback) {
  const value = Number(props.getProperty(key));
  return Number.isSafeInteger(value) && value >= 0 ? value : fallback;
}

function queueEventKey_(position) {
  return EDIT_QUEUE_CFG.EVENT_PREFIX + String(position).padStart(12, '0');
}

/** Ставит одно событие в FIFO за O(1), не перечитывая остальные payload. */
function enqueueEditEvent_(payload) {
  const lock = LockService.getDocumentLock();
  if (!lock || !lock.tryLock(EDIT_QUEUE_CFG.ENQUEUE_LOCK_WAIT_MS)) return false;

  try {
    const props = PropertiesService.getDocumentProperties();
    let head = queuePosition_(props, EDIT_QUEUE_CFG.HEAD_KEY, 1);
    let tail = queuePosition_(props, EDIT_QUEUE_CFG.TAIL_KEY, 0);

    if (head < 1 || tail < head - 1) {
      head = 1;
      tail = 0;
    }

    tail += 1;
    props.setProperty(queueEventKey_(tail), JSON.stringify(payload));

    while (tail - head + 1 > EDIT_QUEUE_CFG.MAX_QUEUE_LEN) {
      props.deleteProperty(queueEventKey_(head));
      head += 1;
    }

    props.setProperties({
      [EDIT_QUEUE_CFG.HEAD_KEY]: String(head),
      [EDIT_QUEUE_CFG.TAIL_KEY]: String(tail)
    });
    return true;
  } finally {
    lock.releaseLock();
  }
}

function tryProcessEditQueue_() {
  const processorLock = LockService.getScriptLock();
  if (!processorLock.tryLock(EDIT_QUEUE_CFG.PROCESS_LOCK_WAIT_MS)) return;

  const startedAt = Date.now();
  try {
    while (Date.now() - startedAt <= EDIT_QUEUE_CFG.DRAIN_TIME_LIMIT_MS) {
      const next = dequeueOneEvent_();
      if (!next) return;

      let statusSs = null;
      try {
        if (next.trackStatus) {
          statusSs = SpreadsheetApp.openById(next.spreadsheetId);
          showActionStatus_(statusSs, 'working', next.statusLabel);
        }

        const result = mainEditLogic(next, statusSs);

        if (next.trackStatus) {
          showActionStatus_(
            statusSs,
            result === false ? 'notDone' : 'success',
            next.statusLabel
          );
        }
      } catch (error) {
        Logger.log(`[QUEUE] событие ${next.id || 'unknown'} завершилось ошибкой: ` +
          (error && error.stack ? error.stack : error));

        if (next.trackStatus) {
          if (!statusSs) {
            try {
              statusSs = SpreadsheetApp.openById(next.spreadsheetId);
            } catch (openError) {
              Logger.log(`[QUEUE] не удалось открыть таблицу для toast: ${openError.message}`);
            }
          }
          showActionStatus_(
            statusSs,
            'error',
            next.statusLabel || 'операция',
            error && error.message ? error.message : String(error)
          );
        }
      }
    }
    Logger.log('[QUEUE] достигнут лимит времени; оставшиеся события будут обработаны следующим запуском.');
  } finally {
    SpreadsheetApp.flush();
    processorLock.releaseLock();
  }
}

/** Снимает одно событие FIFO за O(1). Повреждённые/пропавшие записи пропускает. */
function dequeueOneEvent_() {
  const lock = LockService.getDocumentLock();
  if (!lock || !lock.tryLock(EDIT_QUEUE_CFG.QUEUE_LOCK_WAIT_MS)) return null;

  try {
    const props = PropertiesService.getDocumentProperties();
    let head = queuePosition_(props, EDIT_QUEUE_CFG.HEAD_KEY, 1);
    let tail = queuePosition_(props, EDIT_QUEUE_CFG.TAIL_KEY, 0);

    while (head <= tail) {
      const key = queueEventKey_(head);
      const raw = props.getProperty(key);
      props.deleteProperty(key);
      head += 1;

      if (head > tail) {
        head = 1;
        tail = 0;
      }

      props.setProperties({
        [EDIT_QUEUE_CFG.HEAD_KEY]: String(head),
        [EDIT_QUEUE_CFG.TAIL_KEY]: String(tail)
      });

      if (!raw) {
        if (tail === 0) return null;
        continue;
      }

      try {
        return JSON.parse(raw);
      } catch (error) {
        Logger.log(`[QUEUE] повреждённое событие ${key} удалено: ${error.message}`);
        if (tail === 0) return null;
      }
    }

    if (head !== 1 || tail !== 0) {
      props.setProperties({
        [EDIT_QUEUE_CFG.HEAD_KEY]: '1',
        [EDIT_QUEUE_CFG.TAIL_KEY]: '0'
      });
    }
    return null;
  } finally {
    lock.releaseLock();
  }
}

function isTrueEditValue_(value) {
  return value === true || String(value).toUpperCase() === 'TRUE';
}

/**
 * Выполняет одно событие очереди.
 * Возвращает false, если отслеживаемая операция не была выполнена; любое другое
 * значение трактуется очередью как штатное завершение.
 */
function mainEditLogic(eventData, openedSpreadsheet) {
  if (!eventData || !eventData.spreadsheetId || !eventData.sheetName) {
    throw new Error('mainEditLogic: неполные данные события.');
  }

  const ss = openedSpreadsheet || SpreadsheetApp.openById(eventData.spreadsheetId);
  const sheetName = eventData.sheetName;
  const row = eventData.row;
  const col = eventData.column;
  const val = eventData.value;
  const isTrue = isTrueEditValue_(val);
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    Logger.log(`mainEditLogic: лист «${sheetName}» не найден.`);
    return false;
  }

  const range = sheet.getRange(row, col);

  if (sheetName === 'Запуск') {
    return handleLaunchEvent_(ss, sheet, range, row, col, isTrue);
  }

  if (sheetName === 'Развертывание спецификации') {
    return handleDeploymentSheetEvent_(ss, sheet, range, row, col, val, isTrue);
  }

  if (sheetName === 'База') {
    return handleBaseEdit_(ss, sheet, range, row, col, val, isTrue);
  }

  const roleConfig = getRoleSheetConfig_()[sheetName];
  if (roleConfig) {
    return handleRoleSheetEdit_(ss, sheet, range, row, col, val, isTrue, roleConfig);
  }

  return Boolean(handleNewSheetEdit(ss, sheet, row, col, val));
}

function handleLaunchEvent_(ss, sheet, range, row, col, isTrue) {
  if (col !== COL_LAUNCH_TRIGGER || !isTrue) return false;

  const launchData = sheet.getRange(row, 1, 1, 3).getValues()[0];
  const itemName = String(launchData[0] ?? '').trim();
  const quantity = String(launchData[2] ?? '').trim();

  if (!itemName || !quantity) {
    ss.toast('Недостаточно данных для внесения изделия в базу данных Диспетчера', 'Ошибка', 5);
    range.setValue(false);
    return false;
  }

  handleLaunchEdit(ss, sheet, row);
  return true;
}

function handleDeploymentSheetEvent_(ss, sheet, range, row, col, val, isTrue) {
  if (col === COL_ITEMNAME || col === COL_NOTE || col === COL_QUANTITY) {
    handleNote(ss, sheet, range, row, col, val);
    return true;
  }

  if (col === COL_CHECKBOX && isTrue) {
    const rowData = sheet
      .getRange(row, COL_UUID, 1, COL_QUANTITY - COL_UUID + 1)
      .getValues()[0];
    const itemUUID = String(rowData[0] ?? '').trim();
    const itemName = String(rowData[COL_ITEMNAME - COL_UUID] ?? '').trim();
    const initialQuantity = rowData[COL_QUANTITY - COL_UUID];
    const baseSheet = ss.getSheetByName('База');

    if (!baseSheet) throw new Error('Лист «База» не найден.');

    range.setValue(false);
    return Boolean(handleDeployment(itemUUID, itemName, ss, baseSheet, row, initialQuantity));
  }

  if (col === COL_RETURN_CHECKBOX && isTrue) {
    handleReturn(ss, sheet, range, row, COL_TECHNICAL_CONTROL);
    return true;
  }

  return false;
}

function handleRoleSheetEdit_(ss, sheet, range, row, col, val, isTrue, config) {
  if (config.noteColumns.indexOf(col) !== -1) {
    handleNote(ss, sheet, range, row, col, val);
    return true;
  }

  if (col === COL_RETURN_CHECKBOX && isTrue && config.returnColumn) {
    handleReturn(ss, sheet, range, row, config.returnColumn);
    return true;
  }

  if (col !== COL_CHECKBOX || !isTrue) return false;

  const itemId = runRoleCompletion_(config.completion, ss, sheet, range, row, col, val);
  if (itemId && config.fillNextStage !== false) {
    fillCheckboxes(ss, itemId);
  }
  return Boolean(itemId);
}

function runRoleCompletion_(completion, ss, sheet, range, row, col, val) {
  switch (completion) {
    case 'generic':
      return handleNewSheetEdit(ss, sheet, row, col, val);
    case 'supply':
      return handleSupplyOrder(ss, sheet, range, row);
    case 'picking':
      return handlePickingOrder(ss, sheet, range, row);
    case 'stock':
      return handleStock(ss, sheet, range, row);
    case 'productionStock':
      return handleProductionStock(ss, sheet, range, row);
    case 'production':
      return handleProduction(ss, sheet, range, row);
    case 'acceptance':
      return handleAcceptance(ss, sheet, range, row);
    case 'ready':
      return handleReady(ss, sheet, range, row);
    default:
      throw new Error(`Неизвестный обработчик этапа: ${completion}`);
  }
}

function handleBaseEdit_(ss, sheet, range, row, col, val, isTrue) {
  switch (col) {
    case COL_DEPLOYMENT: {
      const booleanVal = (typeof val === 'string') ? val.toUpperCase() === 'TRUE' : Boolean(val);
      const rowData = sheet
        .getRange(row, COL_UUID, 1, COL_QUANTITY - COL_UUID + 1)
        .getValues()[0];
      const itemUUID = String(rowData[0] ?? '').trim();
      const itemName = String(rowData[COL_ITEMNAME - COL_UUID] ?? '').trim();
      const initialQuantity = rowData[COL_QUANTITY - COL_UUID];

      if (booleanVal) {
        return Boolean(handleDeployment(itemUUID, itemName, ss, sheet, row, initialQuantity));
      }

      handleUndeployment(ss, sheet, row);
      return true;
    }

    case COL_PICKER:
      return isTrue ? Boolean(handleSupplyOrder(ss, sheet, range, row)) : false;

    case COL_STOCK:
      return isTrue ? Boolean(handleStock(ss, sheet, range, row)) : false;

    case COL_PRODUCTION_STOCK:
      return isTrue ? Boolean(handleProductionStock(ss, sheet, range, row)) : false;

    case COL_ACCEPTANCE:
      if (isTrue) {
        ss.toast('Для выполнения приемки перейдите на соответствующий лист', 'Ошибка', 5);
        range.setValue(false);
      }
      return false;

    case COL_READY:
      return isTrue ? Boolean(handleReady(ss, sheet, range, row)) : false;

    default:
      if (isTrue) {
        checkboxOk(range);
        return true;
      }
      return false;
  }
}

/**
 * Универсальный обработчик отметки на листе роли. Настройки листов читаются один
 * раз за execution и повторно используются всеми событиями очереди.
 */
function handleNewSheetEdit(ss, sheet, row, col, val) {
  if (col !== COL_CHECKBOX || !isTrueEditValue_(val)) return null;

  const sheetName = sheet.getName();
  const baseColumn = getSheetToBaseColumnMap_(ss).get(sheetName);
  if (!baseColumn) return null;

  const rowData = sheet
    .getRange(row, COL_UUID, 1, COL_ITEMNAME - COL_UUID + 1)
    .getValues()[0];
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

  sheet.getRange(row, COL_CHECKBOX).setValue(false);
  checkboxOk(baseSheet.getRange(baseRow, baseColumn));
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
    for (const settingsRow of settingsData) {
      const currentSheetName = String(settingsRow[0] ?? '').trim();
      const baseColumnLetter = String(settingsRow[1] ?? '').trim().toUpperCase();
      if (currentSheetName && baseColumnLetter) {
        result.set(currentSheetName, columnLetterToNumber(baseColumnLetter));
      }
    }
  }

  _sheetToBaseColumnMaps.set(spreadsheetId, result);
  return result;
}
