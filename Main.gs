/**
 * Основной обработчик редактирования.
 *
 * Одиночное событие при свободном процессоре выполняется напрямую через
 * e.source без PropertiesService/openById. FIFO-очередь остаётся fallback для
 * параллельных/накопившихся событий. Идентичность строки фиксируется по UUID,
 * а не по физическому номеру строки динамического FILTER-листа.
 */

const EDIT_QUEUE_CFG = {
  EVENT_PREFIX: 'EDIT_QUEUE_V3_EVENT_',
  HEAD_KEY: 'EDIT_QUEUE_V3_HEAD',
  TAIL_KEY: 'EDIT_QUEUE_V3_TAIL',
  MAX_QUEUE_LEN: 1000,
  ENQUEUE_LOCK_WAIT_MS: 5000,
  QUEUE_LOCK_WAIT_MS: 5000,
  PROCESS_LOCK_WAIT_MS: 500,
  FAST_PATH_LOCK_WAIT_MS: 25,
  FAST_PATH_QUEUE_LOCK_WAIT_MS: 50,
  DRAIN_TIME_LIMIT_MS: 240000,
  CONTINUATION_DELAY_MS: 1000,
  PERF_LOG_THRESHOLD_MS: 500
};

const STATUS_TOAST_CFG = Object.freeze({
  TITLE: 'Диспетчер',
  QUEUED_TIMEOUT: -1,
  WORKING_TIMEOUT: -1,
  SUCCESS_TIMEOUT: 5,
  NOT_DONE_TIMEOUT: 8,
  ERROR_TIMEOUT: 10
});

const _sheetToBaseFieldMaps = new Map();
let _roleSheetConfigCache = null;

function getRoleSheetConfig_() {
  if (_roleSheetConfigCache) return _roleSheetConfigCache;

  _roleSheetConfigCache = Object.freeze({
    'Проектирование': {
      noteColumns: [ROLE_COL_ITEMNAME, ROLE_COL_NOTE, ROLE_COL_QUANTITY],
      completion: 'generic'
    },
    'Техконтроль': {
      noteColumns: [ROLE_COL_ITEMNAME, ROLE_COL_NOTE, ROLE_COL_QUANTITY],
      completion: 'generic',
      returnField: 'stage_design'
    },
    'Технолог': {
      noteColumns: [ROLE_COL_ITEMNAME, ROLE_COL_NOTE, ROLE_COL_QUANTITY],
      completion: 'generic',
      returnField: 'stage_deployment'
    },
    'Проверка технологии': {
      noteColumns: [ROLE_COL_NOTE],
      completion: 'generic',
      returnField: 'stage_technologist'
    },
    'Диспетчер- комплектовщик': {
      noteColumns: [ROLE_COL_ITEMNAME, ROLE_COL_NOTE, ROLE_COL_QUANTITY],
      completion: 'supply',
      returnField: 'stage_technology_check'
    },
    'Комплектация': {
      noteColumns: [ROLE_COL_ITEMNAME, ROLE_COL_NOTE, ROLE_COL_QUANTITY, ROLE_COL_DESTINATION],
      completion: 'picking',
      returnField: 'stage_picker'
    },
    'Склад': {
      noteColumns: [ROLE_COL_NOTE],
      completion: 'stock',
      returnField: 'stage_picking'
    },
    'Склад производства': {
      noteColumns: [ROLE_COL_NOTE],
      completion: 'productionStock',
      returnField: 'stage_stock'
    },
    'Производство': {
      noteColumns: [ROLE_COL_NOTE],
      completion: 'production',
      returnField: 'stage_production_stock'
    },
    'ОТК': {
      noteColumns: [ROLE_COL_NOTE],
      completion: 'generic',
      returnField: 'stage_production'
    },
    'Приемка': {
      noteColumns: [ROLE_COL_NOTE, ROLE_COL_QUANTITY],
      completion: 'acceptance',
      returnField: 'stage_quality_check'
    },
    'Готово': {
      noteColumns: [ROLE_COL_NOTE],
      completion: 'ready',
      returnField: 'stage_acceptance'
    }
  });
  return _roleSheetConfigCache;
}

const BASE_STAGE_ACTIONS = Object.freeze({
  stage_design: 'Проектирование',
  stage_technical_control: 'Техконтроль',
  stage_deployment: 'Развертывание спецификации',
  stage_technologist: 'Технолог',
  stage_technology_check: 'Проверка технологии',
  stage_picker: 'Диспетчер- комплектовщик',
  stage_picking: 'Комплектация',
  stage_stock: 'Склад',
  stage_production_stock: 'Склад производства',
  stage_production: 'Производство',
  stage_quality_check: 'ОТК',
  stage_acceptance: 'Приемка',
  stage_ready: 'Готово'
});

function getTrackedActionName_(ss, sheetName, col, value) {
  const isTrue = isTrueEditValue_(value);

  if (sheetName === 'Запуск' && col === COL_LAUNCH_TRIGGER && isTrue) return 'Запуск изделия';

  if (sheetName === 'Развертывание спецификации') {
    if (col === COL_CHECKBOX && isTrue) return 'Развертывание спецификации';
    if (col === COL_RETURN_CHECKBOX && isTrue) return 'Возврат на предыдущий этап';
    return '';
  }

  if (sheetName === BASE_SHEET_NAME) {
    const key = baseKeyForColumn_(ss, col);
    if (key === 'stage_deployment') {
      return isTrue ? 'Развертывание спецификации' : 'Сворачивание спецификации';
    }
    return isTrue ? (BASE_STAGE_ACTIONS[key] || '') : '';
  }

  const roleConfig = getRoleSheetConfig_()[sheetName];
  if (roleConfig && isTrue) {
    if (col === COL_CHECKBOX) return `Этап «${sheetName}»`;
    if (col === COL_RETURN_CHECKBOX && roleConfig.returnField) return `Возврат из «${sheetName}»`;
  }

  if (isTrue && col === COL_CHECKBOX) return `Этап «${sheetName}»`;
  return '';
}

function getStatusItemName_(ss, sheet, row, sheetName) {
  try {
    let itemColumn;
    if (sheetName === 'Запуск') itemColumn = 1;
    else if (sheetName === BASE_SHEET_NAME) itemColumn = baseCol_(ss, 'item_name');
    else itemColumn = ROLE_COL_ITEMNAME;
    return String(sheet.getRange(row, itemColumn).getDisplayValue() || '').trim();
  } catch (error) {
    Logger.log(`getStatusItemName_: ${error.message}`);
    return '';
  }
}

/** Фиксирует логическую идентичность строки в момент исходного edit. */
function captureEditIdentity_(ss, sheet, row, sheetName) {
  if (sheetName === 'Запуск') {
    return { launchData: sheet.getRange(row, 1, 1, 4).getValues()[0] };
  }
  try {
    const uuidColumn = sheetName === BASE_SHEET_NAME ? baseCol_(ss, 'uuid') : ROLE_COL_UUID;
    return { itemUuid: String(sheet.getRange(row, uuidColumn).getValue() || '').trim() };
  } catch (error) {
    Logger.log(`captureEditIdentity_: ${error.message}`);
    return {};
  }
}

function buildStatusLabel_(actionName, itemName) {
  if (!actionName) return '';
  return itemName ? `${actionName}: «${itemName}»` : actionName;
}

function showActionStatus_(ss, state, label, details) {
  if (!ss || !label) return;
  let message;
  let timeout;
  switch (state) {
    case 'queued': message = `⏳ В очереди: ${label}`; timeout = STATUS_TOAST_CFG.QUEUED_TIMEOUT; break;
    case 'working': message = `⚙️ Выполняется: ${label}`; timeout = STATUS_TOAST_CFG.WORKING_TIMEOUT; break;
    case 'success': message = `✅ Завершено: ${label}`; timeout = STATUS_TOAST_CFG.SUCCESS_TIMEOUT; break;
    case 'notDone': message = `⚠️ Операция не выполнена: ${label}`; timeout = STATUS_TOAST_CFG.NOT_DONE_TIMEOUT; break;
    case 'error':
      message = `❌ Ошибка: ${label}` + (details ? `\n${details}` : '');
      timeout = STATUS_TOAST_CFG.ERROR_TIMEOUT;
      break;
    default: return;
  }
  try { ss.toast(message, STATUS_TOAST_CFG.TITLE, timeout); }
  catch (error) { Logger.log(`showActionStatus_: ${error.message}`); }
}

function logSlowEdit_(path, payload, startedAt) {
  const elapsed = Date.now() - startedAt;
  if (elapsed < EDIT_QUEUE_CFG.PERF_LOG_THRESHOLD_MS) return;
  Logger.log(`[PERF] ${path} ${elapsed} ms | ${payload.sheetName}!${payload.a1Notation || ''} | ${payload.statusLabel || payload.id || ''}`);
}

function executeEditPayload_(payload, spreadsheet, path) {
  const startedAt = Date.now();
  try {
    setDispatcherRuntimeSpreadsheet_(spreadsheet);
    if (payload.trackStatus) showActionStatus_(spreadsheet, 'working', payload.statusLabel);
    const result = mainEditLogic(payload, spreadsheet);
    if (payload.trackStatus) {
      showActionStatus_(spreadsheet, result === false ? 'notDone' : 'success', payload.statusLabel);
    }
    return result;
  } catch (error) {
    Logger.log(`[${path}] событие ${payload.id || 'unknown'} завершилось ошибкой: ` + (error && error.stack ? error.stack : error));
    if (payload.trackStatus) {
      showActionStatus_(spreadsheet, 'error', payload.statusLabel || 'операция', error && error.message ? error.message : String(error));
    }
    return false;
  } finally {
    logSlowEdit_(path, payload, startedAt);
  }
}

function onEditHandler(e) {
  if (!e || !e.range) {
    Logger.log('onEditHandler: событие отсутствует или не содержит range.');
    return;
  }

  try {
    const range = e.range;
    if (range.getNumRows() !== 1 || range.getNumColumns() !== 1) {
      Logger.log(`onEditHandler: многоклеточное редактирование ${range.getA1Notation()} не обрабатывается автоматически.`);
      return;
    }

    const sheet = range.getSheet();
    const spreadsheet = e.source || sheet.getParent();
    setDispatcherRuntimeSpreadsheet_(spreadsheet);
    const rawValue = (typeof e.value !== 'undefined') ? e.value : range.getValue();
    const sheetName = sheet.getName();
    const actionName = getTrackedActionName_(spreadsheet, sheetName, range.getColumn(), rawValue);
    const itemName = actionName ? getStatusItemName_(spreadsheet, sheet, range.getRow(), sheetName) : '';
    const identity = captureEditIdentity_(spreadsheet, sheet, range.getRow(), sheetName);

    const payload = {
      ts: Date.now(),
      id: Utilities.getUuid(),
      spreadsheetId: spreadsheet.getId(),
      sheetName,
      a1Notation: range.getA1Notation(),
      row: range.getRow(),
      column: range.getColumn(),
      value: rawValue,
      ...identity,
      trackStatus: Boolean(actionName),
      statusLabel: buildStatusLabel_(actionName, itemName)
    };

    if (tryFastEditPath_(payload, spreadsheet)) return;

    if (!enqueueEditEvent_(payload)) {
      Logger.log(`onEditHandler: не удалось поставить событие ${payload.id} в очередь.`);
      if (payload.trackStatus) showActionStatus_(spreadsheet, 'error', payload.statusLabel, 'Очередь переполнена или недоступна.');
      return;
    }
    if (payload.trackStatus) showActionStatus_(spreadsheet, 'queued', payload.statusLabel);
    tryProcessEditQueue_(spreadsheet);
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

function isEditQueueEmpty_() {
  const lock = LockService.getDocumentLock();
  if (!lock || !lock.tryLock(EDIT_QUEUE_CFG.FAST_PATH_QUEUE_LOCK_WAIT_MS)) return false;
  try {
    const props = PropertiesService.getDocumentProperties();
    const head = queuePosition_(props, EDIT_QUEUE_CFG.HEAD_KEY, 1);
    const tail = queuePosition_(props, EDIT_QUEUE_CFG.TAIL_KEY, 0);
    return tail === 0 || tail < head;
  } finally {
    lock.releaseLock();
  }
}

function tryFastEditPath_(payload, spreadsheet) {
  const processorLock = LockService.getScriptLock();
  if (!processorLock.tryLock(EDIT_QUEUE_CFG.FAST_PATH_LOCK_WAIT_MS)) return false;
  try {
    if (!isEditQueueEmpty_()) return false;
    executeEditPayload_(payload, spreadsheet, 'FAST');
    drainEditQueueLocked_(spreadsheet, Date.now());
    return true;
  } finally {
    processorLock.releaseLock();
  }
}

function enqueueEditEvent_(payload) {
  const lock = LockService.getDocumentLock();
  if (!lock || !lock.tryLock(EDIT_QUEUE_CFG.ENQUEUE_LOCK_WAIT_MS)) return false;
  try {
    const props = PropertiesService.getDocumentProperties();
    let head = queuePosition_(props, EDIT_QUEUE_CFG.HEAD_KEY, 1);
    let tail = queuePosition_(props, EDIT_QUEUE_CFG.TAIL_KEY, 0);
    if (head < 1 || tail < head - 1) { head = 1; tail = 0; }
    if (tail - head + 1 >= EDIT_QUEUE_CFG.MAX_QUEUE_LEN) {
      Logger.log(`[QUEUE] очередь заполнена (${EDIT_QUEUE_CFG.MAX_QUEUE_LEN}); событие не удаляется молча.`);
      return false;
    }
    tail += 1;
    props.setProperty(queueEventKey_(tail), JSON.stringify(payload));
    props.setProperties({
      [EDIT_QUEUE_CFG.HEAD_KEY]: String(head),
      [EDIT_QUEUE_CFG.TAIL_KEY]: String(tail)
    });
    return true;
  } finally {
    lock.releaseLock();
  }
}

function tryProcessEditQueue_(sourceSpreadsheet) {
  const processorLock = LockService.getScriptLock();
  if (!processorLock.tryLock(EDIT_QUEUE_CFG.PROCESS_LOCK_WAIT_MS)) return;
  try {
    drainEditQueueLocked_(sourceSpreadsheet || null, Date.now());
  } finally {
    processorLock.releaseLock();
  }
}

function scheduleEditQueueContinuation_() {
  const exists = ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'processQueuedEdits_');
  if (!exists) ScriptApp.newTrigger('processQueuedEdits_').timeBased().after(EDIT_QUEUE_CFG.CONTINUATION_DELAY_MS).create();
}

function processQueuedEdits_() {
  // Одноразовый continuation trigger удаляет себя перед обработкой; если лимит снова
  // будет достигнут, drainEditQueueLocked_ создаст следующий.
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'processQueuedEdits_')
    .forEach(t => ScriptApp.deleteTrigger(t));
  tryProcessEditQueue_(null);
}

function drainEditQueueLocked_(sourceSpreadsheet, startedAt) {
  while (Date.now() - startedAt <= EDIT_QUEUE_CFG.DRAIN_TIME_LIMIT_MS) {
    const next = dequeueOneEvent_();
    if (!next) return;
    let ss = sourceSpreadsheet;
    if (!ss || ss.getId() !== next.spreadsheetId) {
      try { ss = SpreadsheetApp.openById(next.spreadsheetId); }
      catch (error) {
        Logger.log(`[QUEUE] не удалось открыть таблицу ${next.spreadsheetId}: ${error.message}`);
        continue;
      }
    }
    executeEditPayload_(next, ss, 'QUEUE');
  }
  Logger.log('[QUEUE] достигнут лимит времени; назначено автоматическое продолжение.');
  scheduleEditQueueContinuation_();
}

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
      if (head > tail) { head = 1; tail = 0; }
      props.setProperties({
        [EDIT_QUEUE_CFG.HEAD_KEY]: String(head),
        [EDIT_QUEUE_CFG.TAIL_KEY]: String(tail)
      });
      if (!raw) {
        if (tail === 0) return null;
        continue;
      }
      try { return JSON.parse(raw); }
      catch (error) {
        Logger.log(`[QUEUE] повреждённое событие ${key} удалено: ${error.message}`);
        if (tail === 0) return null;
      }
    }
    if (head !== 1 || tail !== 0) {
      props.setProperties({ [EDIT_QUEUE_CFG.HEAD_KEY]: '1', [EDIT_QUEUE_CFG.TAIL_KEY]: '0' });
    }
    return null;
  } finally {
    lock.releaseLock();
  }
}

function isTrueEditValue_(value) {
  return value === true || String(value).toUpperCase() === 'TRUE';
}

/** Возвращает актуальную строку события по UUID после возможного сдвига FILTER. */
function resolveEventRow_(ss, sheet, eventData) {
  const itemUuid = String(eventData.itemUuid || '').trim();
  if (!itemUuid || eventData.sheetName === 'Запуск') return eventData.row;
  if (eventData.sheetName === BASE_SHEET_NAME) return findBaseRowByUuid_(sheet, itemUuid);
  return findRowById(sheet, itemUuid);
}

function mainEditLogic(eventData, openedSpreadsheet) {
  if (!eventData || !eventData.spreadsheetId || !eventData.sheetName) {
    throw new Error('mainEditLogic: неполные данные события.');
  }
  const ss = openedSpreadsheet || SpreadsheetApp.openById(eventData.spreadsheetId);
  setDispatcherRuntimeSpreadsheet_(ss);
  const sheetName = eventData.sheetName;
  const col = eventData.column;
  const val = eventData.value;
  const isTrue = isTrueEditValue_(val);
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    Logger.log(`mainEditLogic: лист «${sheetName}» не найден.`);
    return false;
  }

  if (sheetName === 'Запуск') {
    const range = sheet.getRange(eventData.row, col);
    return handleLaunchEvent_(ss, sheet, range, eventData.row, col, isTrue, eventData.launchData || null);
  }

  const row = resolveEventRow_(ss, sheet, eventData);
  if (!row) {
    Logger.log(`mainEditLogic: UUID ${eventData.itemUuid || ''} больше не находится на листе «${sheetName}»; событие считаем устаревшим.`);
    return true;
  }
  const range = sheet.getRange(row, col);

  if (sheetName === 'Развертывание спецификации') return handleDeploymentSheetEvent_(ss, sheet, range, row, col, val, isTrue);
  if (sheetName === BASE_SHEET_NAME) return handleBaseEdit_(ss, sheet, range, row, col, val, isTrue);

  const roleConfig = getRoleSheetConfig_()[sheetName];
  if (roleConfig) return handleRoleSheetEdit_(ss, sheet, range, row, col, val, isTrue, roleConfig);
  return Boolean(handleNewSheetEdit(ss, sheet, row, col, val));
}

function handleLaunchEvent_(ss, sheet, range, row, col, isTrue, launchData) {
  if (col !== COL_LAUNCH_TRIGGER || !isTrue) return false;
  const data = Array.isArray(launchData) ? launchData : sheet.getRange(row, 1, 1, 4).getValues()[0];
  const itemName = String(data[0] ?? '').trim();
  const quantity = String(data[2] ?? '').trim();
  if (!itemName || !quantity) {
    ss.toast('Недостаточно данных для внесения изделия в базу данных Диспетчера', 'Ошибка', 5);
    if (row <= sheet.getMaxRows()) range.setValue(false);
    return false;
  }
  handleLaunchSnapshot_(ss, sheet, data, row);
  return true;
}

function handleDeploymentSheetEvent_(ss, sheet, range, row, col, val, isTrue) {
  if ([ROLE_COL_ITEMNAME, ROLE_COL_NOTE, ROLE_COL_QUANTITY].indexOf(col) !== -1) {
    handleNote(ss, sheet, range, row, col, val);
    return true;
  }

  if (col === COL_CHECKBOX && isTrue) {
    const rowData = sheet.getRange(row, ROLE_COL_UUID, 1, ROLE_COL_QUANTITY - ROLE_COL_UUID + 1).getValues()[0];
    const itemUUID = String(rowData[0] ?? '').trim();
    const itemName = String(rowData[ROLE_COL_ITEMNAME - ROLE_COL_UUID] ?? '').trim();
    const initialQuantity = rowData[ROLE_COL_QUANTITY - ROLE_COL_UUID];
    const baseSheet = ss.getSheetByName(BASE_SHEET_NAME);
    if (!baseSheet) throw new Error('Лист «База» не найден.');
    range.setValue(false);
    return Boolean(handleDeployment(itemUUID, itemName, ss, baseSheet, row, initialQuantity));
  }

  if (col === COL_RETURN_CHECKBOX && isTrue) {
    handleReturn(ss, sheet, range, row, 'stage_technical_control');
    return true;
  }
  return false;
}

function handleRoleSheetEdit_(ss, sheet, range, row, col, val, isTrue, config) {
  if (config.noteColumns.indexOf(col) !== -1) {
    handleNote(ss, sheet, range, row, col, val);
    return true;
  }
  if (col === COL_RETURN_CHECKBOX && isTrue && config.returnField) {
    handleReturn(ss, sheet, range, row, config.returnField);
    return true;
  }
  if (col !== COL_CHECKBOX || !isTrue) return false;

  const itemId = runRoleCompletion_(config.completion, ss, sheet, range, row, col, val);
  return Boolean(itemId);
}

function runRoleCompletion_(completion, ss, sheet, range, row, col, val) {
  switch (completion) {
    case 'generic': return handleNewSheetEdit(ss, sheet, row, col, val);
    case 'supply': return handleSupplyOrder(ss, sheet, range, row);
    case 'picking': return handlePickingOrder(ss, sheet, range, row);
    case 'stock': return handleStock(ss, sheet, range, row);
    case 'productionStock': return handleProductionStock(ss, sheet, range, row);
    case 'production': return handleProduction(ss, sheet, range, row);
    case 'acceptance': return handleAcceptance(ss, sheet, range, row);
    case 'ready': return handleReady(ss, sheet, range, row);
    default: throw new Error(`Неизвестный обработчик этапа: ${completion}`);
  }
}

function handleBaseEdit_(ss, sheet, range, row, col, val, isTrue) {
  const fieldKey = baseKeyForColumn_(ss, col);
  if (!fieldKey) return false;

  if (fieldKey === 'stage_deployment') {
    const data = readBaseFields_(sheet, row, ['uuid', 'item_name', 'quantity']);
    const itemUUID = String(data.uuid ?? '').trim();
    const itemName = String(data.item_name ?? '').trim();
    if (isTrue) return Boolean(handleDeployment(itemUUID, itemName, ss, sheet, row, data.quantity));
    handleUndeployment(ss, sheet, row);
    return true;
  }

  if (fieldKey === 'stage_picker') return isTrue ? Boolean(handleSupplyOrder(ss, sheet, range, row)) : false;
  if (fieldKey === 'stage_picking') return isTrue ? Boolean(handlePickingOrder(ss, sheet, range, row)) : false;
  if (fieldKey === 'stage_stock') return isTrue ? Boolean(handleStock(ss, sheet, range, row)) : false;
  if (fieldKey === 'stage_production_stock') return isTrue ? Boolean(handleProductionStock(ss, sheet, range, row)) : false;
  if (fieldKey === 'stage_production') return isTrue ? Boolean(handleProduction(ss, sheet, range, row)) : false;
  if (fieldKey === 'stage_acceptance') {
    if (isTrue) {
      ss.toast('Для выполнения приемки перейдите на соответствующий лист', 'Ошибка', 5);
      range.setValue(false);
    }
    return false;
  }
  if (fieldKey === 'stage_ready') return isTrue ? Boolean(handleReady(ss, sheet, range, row)) : false;

  if (isTrue && STAGE_FIELD_KEYS.indexOf(fieldKey) !== -1) {
    checkboxOk(range);
    return true;
  }
  return false;
}

function handleNewSheetEdit(ss, sheet, row, col, val) {
  if (col !== COL_CHECKBOX || !isTrueEditValue_(val)) return null;

  const sheetName = sheet.getName();
  const stageField = getSheetToBaseFieldMap_(ss).get(sheetName);
  if (!stageField) return null;

  const rowData = sheet.getRange(row, ROLE_COL_UUID, 1, ROLE_COL_ITEMNAME - ROLE_COL_UUID + 1).getValues()[0];
  const itemId = String(rowData[0] ?? '').trim();
  const itemName = String(rowData[ROLE_COL_ITEMNAME - ROLE_COL_UUID] ?? '').trim();
  if (!itemId) {
    sheet.getRange(row, COL_CHECKBOX).setValue(false);
    Logger.log(`handleNewSheetEdit: отсутствует UUID для «${itemName}», строка ${row}.`);
    return null;
  }

  const baseSheet = ss.getSheetByName(BASE_SHEET_NAME);
  if (!baseSheet) throw new Error('Лист «База» не найден.');
  const baseRow = findBaseRowByUuid_(baseSheet, itemId);
  if (!baseRow) {
    Logger.log(`handleNewSheetEdit: UUID ${itemId} не найден в «База».`);
    return null;
  }

  sheet.getRange(row, COL_CHECKBOX).setValue(false);
  checkboxOk(baseSheet.getRange(baseRow, baseCol_(ss, stageField)));
  return itemId;
}

function getSheetToBaseFieldMap_(ss) {
  const spreadsheetId = ss.getId();
  if (_sheetToBaseFieldMaps.has(spreadsheetId)) return _sheetToBaseFieldMaps.get(spreadsheetId);

  const result = new Map();
  const settingsSheet = ss.getSheetByName('Настройки');
  if (!settingsSheet) {
    _sheetToBaseFieldMaps.set(spreadsheetId, result);
    return result;
  }

  const lastRow = settingsSheet.getLastRow();
  if (lastRow >= 2) {
    const settingsData = settingsSheet.getRange(2, 1, lastRow - 1, 2).getValues();
    for (const settingsRow of settingsData) {
      const currentSheetName = String(settingsRow[0] ?? '').trim();
      const fieldKey = String(settingsRow[1] ?? '').trim();
      if (!currentSheetName || !fieldKey) continue;
      baseCol_(ss, fieldKey);
      result.set(currentSheetName, fieldKey);
    }
  }

  _sheetToBaseFieldMaps.set(spreadsheetId, result);
  return result;
}
