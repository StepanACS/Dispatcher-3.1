/**
 * Основной обработчик редактирования — выполняет задачу напрямую.
 */
/**
 * onEditHandler(e)
 * ------------------------------------------------------------
 * Идеология:
 * 1) onEdit должен выполняться быстро: только собрать контекст и положить событие в очередь.
 * 2) Очередь хранится в DocumentProperties (общая для всех пользователей этого файла).
 * 3) Один “процессор” (победитель по LockService) последовательно вынимает события из очереди
 *    и вызывает mainEditLogic(payload) строго по одному.
 *
 * Важно:
 * - В очередь нельзя класть объекты Sheet/Range (они не сериализуются). Поэтому sourceSheet не передаем.
 * - При обработке mainEditLogic(payload) при необходимости заново получайте sheet через
 *   SpreadsheetApp.getActive().getSheetByName(payload.sheetName).
 */

// ------------------------- CONFIG -------------------------
const EDIT_QUEUE_CFG = {
  QUEUE_PROP_KEY: 'EDIT_QUEUE_V1',     // где хранится очередь
  MAX_QUEUE_LEN: 300,                  // защита от раздувания (подберите)
  PROCESS_LOCK_WAIT_MS: 15000,         // ждать ли процессору захват замка
  ENQUEUE_LOCK_WAIT_MS: 2000,          // короткий замок на добавление в очередь
  DRAIN_TIME_LIMIT_MS: 240000          // ограничение “дренажа” за один запуск (≈4 мин)
};

// ------------------------- ENTRY POINT -------------------------
function onEditHandler(e) {
  Logger.log('--- Начало onEditHandler ---');

  if (!e || !e.range) {
    Logger.log("Ошибка: событие 'e' пустое или без range. Возможно, ручной запуск.");
    Logger.log('--- Конец onEditHandler ---');
    return;
  }

  try {
    const range = e.range;
    const sheet = range.getSheet();

    // Нормализуем value: для чекбокса e.value обычно строка "TRUE"/"FALSE", для некоторых случаев e.value undefined
    const rawValue = (typeof e.value !== 'undefined') ? e.value : range.getValue();

    const payload = {
      // метаданные события
      ts: Date.now(),
      id: Utilities.getUuid(),

      sheetName: sheet.getName(),
      a1Notation: range.getA1Notation(),
      row: range.getRow(),
      column: range.getColumn(),

      // значение (храним как есть, без попытки привести к типу)
      value: rawValue,

      // пользователь/режим (в установленных триггерах Session.getActiveUser() часто пустой)
      user: safeGetActiveUserEmail_(),
      authMode: e.authMode
    };

    Logger.log(`onEditHandler: enqueue ${payload.sheetName}!${payload.a1Notation}, value='${payload.value}'`);

    // 1) Кладем событие в очередь (под коротким замком)
    enqueueEditEvent_(payload);

    // 2) Пытаемся стать “процессором” и обработать очередь (если уже есть процессор — просто выходим)
    tryProcessEditQueue_();

  } catch (error) {
    Logger.log('onEditHandler: Общая ошибка: ' + (error && error.message ? error.message : error));
  }

  Logger.log('--- Конец onEditHandler ---');
}

// ------------------------- QUEUE: ENQUEUE -------------------------
function enqueueEditEvent_(payload) {
  const lock = LockService.getDocumentLock(); // именно “по документу”
  const got = lock.tryLock(EDIT_QUEUE_CFG.ENQUEUE_LOCK_WAIT_MS);

  if (!got) {
    // Не смогли быстро взять замок на очередь — лучше не зависать в onEdit.
    // Можно либо “сдаться”, либо попробовать очень короткий второй раз.
    Logger.log('[QUEUE] enqueue: lock busy, skip enqueue (rare). payload id=' + payload.id);
    return;
  }

  try {
    const props = PropertiesService.getDocumentProperties();

    // Храним как массив строк JSON (чуть компактнее, чем массив объектов)
    const current = props.getProperty(EDIT_QUEUE_CFG.QUEUE_PROP_KEY);
    const arr = current ? JSON.parse(current) : [];

    arr.push(JSON.stringify(payload));

    // защита от переполнения
    if (arr.length > EDIT_QUEUE_CFG.MAX_QUEUE_LEN) {
      const overflow = arr.length - EDIT_QUEUE_CFG.MAX_QUEUE_LEN;
      arr.splice(0, overflow); // выкидываем самые старые
      Logger.log(`[QUEUE] overflow: dropped ${overflow} oldest events`);
    }

    props.setProperty(EDIT_QUEUE_CFG.QUEUE_PROP_KEY, JSON.stringify(arr));
  } finally {
    lock.releaseLock();
  }
}

// ------------------------- QUEUE: PROCESSOR -------------------------
function tryProcessEditQueue_() {
  const lock = LockService.getDocumentLock();

  // ВАЖНО: здесь именно waitLock/tryLock — чтобы собрать очередь и обработать последовательно.
  // Если процессор уже работает, просто выходим (очередь будет обработана им).
  const got = lock.tryLock(EDIT_QUEUE_CFG.PROCESS_LOCK_WAIT_MS);
  if (!got) {
    Logger.log('[QUEUE] processor: another processor is running; exit');
    return;
  }

  const startedAt = Date.now();

  try {
    while (true) {
      // ограничение по времени, чтобы не упереться в лимиты Apps Script
      if (Date.now() - startedAt > EDIT_QUEUE_CFG.DRAIN_TIME_LIMIT_MS) {
        Logger.log('[QUEUE] drain: time limit reached; remaining events will be processed by next trigger');
        return;
      }

      const next = dequeueOneEvent_();
      if (!next) {
        Logger.log('[QUEUE] drain: queue empty');
        return;
      }

      // Обработка строго по одному событию
      try {
        // Если вашей mainEditLogic нужен sheet-объект — получите его внутри mainEditLogic:
        // const sheet = SpreadsheetApp.getActive().getSheetByName(payload.sheetName);
        mainEditLogic(next);
      } catch (err) {
        Logger.log('[QUEUE] event failed id=' + next.id + ' err=' + (err && err.message ? err.message : err));
        // Ошибку не “возвращаем” в очередь, иначе можно зациклиться.
      }
    }
  } finally {
    lock.releaseLock();
  }
}

// ------------------------- QUEUE: DEQUEUE (one) -------------------------
function dequeueOneEvent_() {
  // ВНИМАНИЕ: мы уже под PROCESS_LOCK, но всё равно работаем аккуратно.
  const props = PropertiesService.getDocumentProperties();
  const current = props.getProperty(EDIT_QUEUE_CFG.QUEUE_PROP_KEY);
  if (!current) return null;

  const arr = JSON.parse(current);
  if (!arr.length) return null;

  const firstStr = arr.shift();
  props.setProperty(EDIT_QUEUE_CFG.QUEUE_PROP_KEY, JSON.stringify(arr));

  try {
    return JSON.parse(firstStr);
  } catch (e) {
    Logger.log('[QUEUE] bad JSON event, skipped');
    return null;
  }
}

// ------------------------- HELPERS -------------------------
function safeGetActiveUserEmail_() {
  try {
    const email = Session.getActiveUser().getEmail();
    return email || '';
  } catch (e) {
    return '';
  }
}


function mainEditLogic(eventData) {
    Logger.log("Триггер сработал, вызвана mainEditLogic.");
    Logger.log(`mainEditLogic: Входящие данные eventData: ${JSON.stringify(eventData)}`);

    // *** УНИВЕРСАЛЬНОЕ ОПРЕДЕЛЕНИЕ ПЕРЕМЕННЫХ ***
    const ss = SpreadsheetApp.openById("1XffWoUclgQyFcT4DOIRoenQ9jAbBxXam6YxyY8xgv_4");
    const sheetName = eventData.sheetName;
    const row = eventData.row;
    const col = eventData.column;
    const val = eventData.value;
    const a1Notation = eventData.a1Notation;
    const sourceSheet = eventData.sourceSheet;

    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
        Logger.log(`mainEditLogic: Ошибка: Лист с именем '${sheetName}' не найден.`);
        ss.toast(`Ошибка: Лист '${sheetName}' не найден.`, "Ошибка", 5);
        return;
    }
    const range = sheet.getRange(a1Notation);
    if (!range) {
        Logger.log(`mainEditLogic: Ошибка: Диапазон '${a1Notation}' не найден на листе '${sheetName}'.`);
        ss.toast(`Ошибка: Диапазон '${a1Notation}' не найден.`, "Ошибка", 5);
        return;
    }

    Logger.log(`Зафиксировано изменение на листе ${sheet.getName()} на строке ${row}, в столбце ${col}, новое значение в ячейке ${val}`);

    let notifications = [];

    switch (sheetName) {
        case "Запуск":
            if (col === COL_LAUNCH_TRIGGER && val === "TRUE") {
                Logger.log(`mainEditLogic: Обнаружен триггер запуска в колонке ${col}. Проверяем данные в A и C.`);

                const valueCol1 = sheet.getRange(row, 1);
                const valueCol3 = sheet.getRange(row, 3);

                if (!valueCol1 || !valueCol3) {
                    Logger.log("mainEditLogic: Ошибка: Не удалось получить диапазоны для проверки A и C.");
                    ss.toast("Ошибка при проверке данных. Пожалуйста, попробуйте еще раз.", "Ошибка", 5);
                    return;
                }

                if (
                    valueCol1.getValue().toString().trim() !== "" &&
                    valueCol3.getValue().toString().trim() !== ""
                ) {
                    Logger.log(`mainEditLogic: Данные в A (${valueCol1.getValue()}) и C (${valueCol3.getValue()}) найдены. Вызываем handleLaunchEdit.`);
                    handleLaunchEdit(ss, sheet, row);
                    //var templateRow = sheet.getLastRow();
                    //testEnsureFreeRows();
                    //ensureFreeRows(sheet, 18, 50);
                    notifications.push("Изделие добавлено в базу");
                    
                } else {
                    Logger.log("mainEditLogic: Недостаточно данных в колонках A или C для внесения изделия.");
                    SpreadsheetApp.getUi().alert("Недостаточно данных для внесения изделия в базу данных Диспетчера");
                    sheet.getRange(row, COL_LAUNCH_TRIGGER).setValue(false); // Сбросить чекбокс обратно
                }
            }
            break;

        case "Проектирование":
            if (col === COL_ITEMNAME || col === COL_NOTE || col === COL_QUANTITY) {
                Logger.log("Запуск обработки примечания");
                handleNote(ss, sheet, range, row, col, val);
            }
            if (col === COL_CHECKBOX && val === "TRUE") {
                const itemId = handleNewSheetEdit(ss, sheet, row, col, val);
                Logger.log("Зафиксировано нажатие на чекбокс на странице Проектирование. Идентификатор изделия: " + itemId);
                fillCheckboxes(ss, itemId);
                sheet.getRange(row, col).setValue(false); // Сброс чекбокса после обработки
            }
            break;

        case "Техконтроль":
            if (col === COL_ITEMNAME || col === COL_NOTE || col === COL_QUANTITY) {
                Logger.log("Запуск обработки примечания");
                handleNote(ss, sheet, range, row, col, val);
            }
            if (col === COL_CHECKBOX && val === "TRUE") {
                const itemId = handleNewSheetEdit(ss, sheet, row, col, val);
                Logger.log("Зафиксировано нажатие на чекбокс на странице Техконтроль. Идентификатор изделия: " + itemId);
                fillCheckboxes(ss, itemId);
                sheet.getRange(row, col).setValue(false); // Сброс чекбокса после обработки
            }
            if (col === COL_RETURN_CHECKBOX && val === "TRUE") {
                handleReturn(ss, sheet, range, row, COL_DESIGN);
            }
            break;    

        case "Развертывание спецификации":
            if (col === COL_ITEMNAME || col === COL_NOTE || col === COL_QUANTITY) {
                Logger.log("Запуск обработки примечания");
                handleNote(ss, sheet, range, row, col, val);
            }
            if (col === COL_CHECKBOX && val === "TRUE") {
                const currentItemName = sheet.getRange(row, COL_ITEMNAME).getValue();
                const currentItemUUID = sheet.getRange(row, COL_UUID).getValue();
                const currentInitialQuantity = sheet.getRange(row, COL_QUANTITY).getValue();

                Logger.log(`Развертывание спецификации: Чекбокс TRUE. Извлечено: itemName="${currentItemName}", itemUUID="${currentItemUUID}", initialQuantity="${currentInitialQuantity}"`);

                const baseSheet = ss.getSheetByName("База");
                if (!baseSheet) {
                    Logger.log(`mainEditLogic: Ошибка: Лист "База" не найден для развертывания спецификации.`);
                    ss.toast(`Ошибка: Лист "База" не найден.`, "Ошибка", 5);
                    return;
                }

                const itemId = handleDeployment(currentItemUUID, currentItemName, ss, baseSheet, row, currentInitialQuantity);
                //fillCheckboxes(ss, itemId);
                // *** ИСПРАВЛЕНИЕ: Сброс чекбокса после успешного развертывания ***
                sourceSheet.getRange(row, col).setValue(false); 
            }
            if (col === COL_RETURN_CHECKBOX && val === "TRUE") {
                handleReturn(ss, sheet, range, row, COL_TECHNICAL_CONTROL);
            }
            break;

        case "Технолог":
            if (col === COL_ITEMNAME || col === COL_NOTE || col === COL_QUANTITY) {
                Logger.log("Запуск обработки примечания");
                handleNote(ss, sheet, range, row, col, val);
            }
            if (col === COL_CHECKBOX && val === "TRUE") {
                const itemId = handleNewSheetEdit(ss, sheet, row, col, val);
                fillCheckboxes(ss, itemId);
                sheet.getRange(row, col).setValue(false); // Сброс чекбокса после обработки
            }
            if (col === COL_RETURN_CHECKBOX && val === "TRUE") {
                handleReturn(ss, sheet, range, row, COL_DEPLOYMENT);
            }
            break;

        case "Проверка технологии":
            if (col === COL_NOTE) {
                Logger.log("Запуск обработки примечания");
                handleNote(ss, sheet, range, row, col, val);
            }
            if (col === COL_CHECKBOX && val === "TRUE") {
                const itemId = handleNewSheetEdit(ss, sheet, row, col, val);
                fillCheckboxes(ss, itemId);
                sheet.getRange(row, col).setValue(false); // Сброс чекбокса после обработки
            }
            if (col === COL_RETURN_CHECKBOX && val === "TRUE") {
                handleReturn(ss, sheet, range, row, COL_TECHNOLOGIEST);
            }
            break;

        case "Диспетчер- комплектовщик":
            if (col === COL_ITEMNAME || col === COL_NOTE || col === COL_QUANTITY) {
                Logger.log("Запуск обработки примечания");
                handleNote(ss, sheet, range, row, col, val);
            }
            if (col === COL_CHECKBOX && val === "TRUE") {
                Logger.log("зафиксировано нажатие на чекбокс на листе Диспетчер-комплектовщик");
                const itemId = handleSupplyOrder(ss, sheet, range, row);
                Logger.log("Выполнена обработка нажатия чекбокса, получен идентификатор изделия: " + itemId);
                fillCheckboxes(ss, itemId);
                sheet.getRange(row, col).setValue(false); // Сброс чекбокса после обработки
            }
            if (col === COL_RETURN_CHECKBOX && val === "TRUE") {
                handleReturn(ss, sheet, range, row, COL_TECHNOLOGY_CHECK);
            }
            break;

        case "Комплектация":
            if (col === COL_ITEMNAME || col === COL_NOTE || col === COL_QUANTITY || col === COL_DESTINATION) {
                Logger.log("Запуск обработки примечания");
                handleNote(ss, sheet, range, row, col, val);
            }
            if (col === COL_CHECKBOX && val === "TRUE") {
                const itemId = handlePickingOrder(ss, sheet, range, row);
                fillCheckboxes(ss, itemId);
                sheet.getRange(row, col).setValue(false); // Сброс чекбокса после обработки
            }
            if (col === COL_RETURN_CHECKBOX && val === "TRUE") {
                handleReturn(ss, sheet, range, row, COL_PICKER);
            }
            break;

        case "Склад":
            if (col === COL_NOTE) {
                Logger.log("Запуск обработки примечания");
                handleNote(ss, sheet, range, row, col, val);
            }
            if (col === COL_CHECKBOX && val === "TRUE") {
                const itemId = handleStock(ss, sheet, range, row);
                fillCheckboxes(ss, itemId);
                sheet.getRange(row, col).setValue(false); // Сброс чекбокса после обработки
            }
            if (col === COL_RETURN_CHECKBOX && val === "TRUE") {
                handleReturn(ss, sheet, range, row, COL_PICKING);
            }
            break;

        case "Склад производства":
            if (col === COL_NOTE) {
                Logger.log("Запуск обработки примечания");
                handleNote(ss, sheet, range, row, col, val);
            }
            if (col === COL_CHECKBOX && val === "TRUE") {
                const itemId = handleProductionStock(ss, sheet, range, row);
                fillCheckboxes(ss, itemId);
                sheet.getRange(row, col).setValue(false); // Сброс чекбокса после обработки
            }
            if (col === COL_RETURN_CHECKBOX && val === "TRUE") {
                handleReturn(ss, sheet, range, row, COL_STOCK);
            }
            break;

        case "Производство":
            if (col === COL_NOTE) {
                Logger.log("Запуск обработки примечания");
                handleNote(ss, sheet, range, row, col, val);
            }
            if (col === COL_CHECKBOX && val === "TRUE") {
                const itemId = handleProduction(ss, sheet, range, row);
                fillCheckboxes(ss, itemId);
                sheet.getRange(row, col).setValue(false); // Сброс чекбокса после обработки
            }
            if (col === COL_RETURN_CHECKBOX && val === "TRUE") {
                handleReturn(ss, sheet, range, row, COL_PRODUCTION_STOCK);
            }
            break;

        case "ОТК":
            if (col === COL_NOTE) {
                Logger.log("Запуск обработки примечания");
                handleNote(ss, sheet, range, row, col, val);
            }
            if (col === COL_CHECKBOX && val === "TRUE") {
                const itemId = handleNewSheetEdit(ss, sheet, row, col, val);
                fillCheckboxes(ss, itemId);
                sheet.getRange(row, col).setValue(false); // Сброс чекбокса после обработки
            }
            if (col === COL_RETURN_CHECKBOX && val === "TRUE") {
                handleReturn(ss, sheet, range, row, COL_PRODUCTION);
            }
            break;

        case "Приемка":
            if (col === COL_NOTE || col === COL_QUANTITY) {
                Logger.log("Запуск обработки примечания");
                handleNote(ss, sheet, range, row, col, val);
            }
            if (col === COL_CHECKBOX && val === "TRUE") {
                const itemId = handleAcceptance(ss, sheet, range, row);
                fillCheckboxes(ss, itemId);
                sheet.getRange(row, col).setValue(false); // Сброс чекбокса после обработки
            }
            if (col === COL_RETURN_CHECKBOX && val === "TRUE") {
                handleReturn(ss, sheet, range, row, COL_QUALITY_CHECK);
            }
            break;

        case "Готово":
            if (col === COL_NOTE) {
                Logger.log("Запуск обработки примечания");
                handleNote(ss, sheet, range, row, col, val);
            }
            if (col === COL_CHECKBOX && val === "TRUE") {
                const itemId = handleReady(ss, sheet, range, row);
                fillCheckboxes(ss, itemId);
                sheet.getRange(row, col).setValue(false); // Сброс чекбокса после обработки
            }
            if (col === COL_RETURN_CHECKBOX && val === "TRUE") {
                handleReturn(ss, sheet, range, row, COL_ACCEPTANCE);
            }
            break;

        case "База":
            switch (col) {
                case COL_DEPLOYMENT:
                    Logger.log(`База, столбец N: Значение чекбокса (val): ${val}, Тип: ${typeof val}`);

                    let booleanVal;
                    if (typeof val === 'string') {
                        booleanVal = (val.toUpperCase() === 'TRUE');
                    } else {
                        booleanVal = val;
                    }

                    const baseSheet = sheet;
                    const itemName = baseSheet.getRange(row, COL_ITEMNAME).getValue();
                    const itemUUID = baseSheet.getRange(row, COL_UUID).getValue();
                    const initialQuantity = baseSheet.getRange(row, COL_QUANTITY).getValue();

                    Logger.log(`База, столбец N: Извлечено: itemName="${itemName}", itemUUID="${itemUUID}", initialQuantity="${initialQuantity}"`);

                    if (booleanVal === true) {
                        handleDeployment(itemUUID, itemName, ss, baseSheet, row, initialQuantity);
                        // На листе "База" чекбокс также нужно сбросить, если он используется как триггер
                        //baseSheet.getRange(row, col).setValue(false); 
                        //baseSheet.getRange(row, col).setValue(false); 
                    } else if (booleanVal === false) {
                        Logger.log("Запуск сворачивания спецификации");
                        handleUndeployment(ss, baseSheet, row);
                    }
                    break;
                case COL_PICKER:
                    if (val === "TRUE") {
                        handleSupplyOrder(ss, sheet, range, row);
                        sheet.getRange(row, col).setValue(false); // Сброс чекбокса
                    }
                    break;
                case COL_STOCK:
                    if (val === "TRUE") {
                        handleStock(ss, sheet, range, row);
                        sheet.getRange(row, col).setValue(false); // Сброс чекбокса
                    }
                    break;
                case COL_PRODUCTION_STOCK:
                    if (val === "TRUE") {
                        handleProductionStock(ss, sheet, range, row);
                        sheet.getRange(row, col).setValue(false); // Сброс чекбокса
                    }
                    break;
                case COL_ACCEPTANCE:
                    if (val === "TRUE") {
                        //handleAcceptance(ss, sheet, range, row);
                        ss.toast("Для выполенния приемки перейдите на соответствующий лист", "Ошибка", 5);
                        sheet.getRange(row, COL_ACCEPTANCE).setValue(false); // Сброс чекбокса
                    }
                    break;
                case COL_READY:
                    if (val === "TRUE") {
                        handleReady(ss, sheet, range, row);
                        sheet.getRange(row, col).setValue(false); // Сброс чекбокса
                    }
                    break;
                default:
                    // Общий сброс чекбокса для других столбцов на листе "База", если они триггерят что-то и должны сбрасываться.
                    // Если вы хотите, чтобы `checkboxOk` просто "одобрил" чекбокс, но не сбрасывал его,
                    // то эту строку нужно будет убрать или обернуть в условие.
                    if (val === "TRUE") {
                        checkboxOk(range);
                        // sheet.getRange(row, col).setValue(false); // Возможно, здесь тоже нужен сброс?
                    }
                    break;
            }
            break;

        default:
            handleNewSheetEdit(ss, sheet, row, col, val);
            // Если handleNewSheetEdit также вызывает перемещение строки, то здесь тоже нужен сброс.
            // Однако, по логике, это, возможно, чекбоксы, которые просто "заполняются" и не перемещают строку.
            // Если это не так, добавьте: sheet.getRange(row, col).setValue(false);
            break;
    }

    notifications.forEach(message => {
        ss.toast(message, "Информация", 5);
    });
}




/**
 * Объединённый обработчик для новых листов.
 * Эта функция вызывается, если редактирование произошло на листе, которое настроено в "Настройки" 
 * (то есть, имя листа присутствует в настройках) и если редактирование произошло в столбце K (номер 11).
 * Она считывает настройки, находит нужную строку в "База" по идентификатору и устанавливает отметку в указанном столбце,
 * а затем сбрасывает исходный чекбокс.
 *
 * @param {Event} e - Объект события onEdit.
 */
function handleNewSheetEdit(ss, sheet, row, col, val) {
  const sheetName = sheet.getName();
    // Проверяем, что редактирование произошло в столбце K (номер 11) и значение = "TRUE"
  if (col !== 11 || val !== "TRUE") return;

  // Получаем настройки из листа "Настройки"
  const settingsSheet = ss.getSheetByName("Настройки");
  if (!settingsSheet) {
    Logger.log(`[${new Date().toLocaleString()}] Лист "Настройки" не найден.`);
    return;
  }
  
  // Считываем весь диапазон настроек и строим объект-отображение:
  // Ключ: имя листа, Значение: номер столбца (полученный из буквенного обозначения)
  const settingsData = settingsSheet.getDataRange().getValues();
  const sheetToBaseColumn = {};
  for (let i = 1; i < settingsData.length; i++) { // начинаем со второй строки (первая — заголовок)
    let currentSheetName = settingsData[i][0].toString().trim();
    let baseColumnLetter = settingsData[i][1].toString().trim().toUpperCase();
    if (currentSheetName && baseColumnLetter) {
      sheetToBaseColumn[currentSheetName] = columnLetterToNumber(baseColumnLetter);
    }
  }
  
  // Если имя текущего листа присутствует в настройках, выполняем обработку
  if (sheetToBaseColumn.hasOwnProperty(sheetName)) {
    // Считываем уникальный идентификатор изделия из столбца D (номер 4) текущего листа
    const itemId = sheet.getRange(row, COL_UUID).getValue().toString().trim();
    // Считываем название изделия из столбца E (номер 5)
    const itemName = sheet.getRange(row, COL_ITEMNAME).getValue().toString().trim();
    Logger.log(`[${new Date().toLocaleString()}] Новая листовая отметка, строка ${row}, изделие: "${itemName}"`);
    
    if (!itemId) {
      Logger.log(`[${new Date().toLocaleString()}] Идентификатор изделия отсутствует в строке ${row}, столбец D.`);
      sheet.getRange(row, 11).setValue(false);
      return;
    }
    
    // Получаем лист "База"
    const baseSheet = ss.getSheetByName("База");
    if (!baseSheet) {
      Logger.log(`[${new Date().toLocaleString()}] Лист "База" не найден.`);
      return;
    }
    
    // Ищем строку в "База", где в столбце D (COL_UUID) совпадает itemId
    const baseRow = findRowById(baseSheet, itemId);
    if (baseRow) {
      // Из объекта настроек получаем номер столбца, в котором необходимо установить отметку на "База"
      const baseColumn = sheetToBaseColumn[sheetName];
      // Устанавливаем отметку (чекбокс) на листе "База"
      checkboxOk(baseSheet.getRange(baseRow, baseColumn));
      // Сбрасываем исходный чекбокс на текущем листе
      sheet.getRange(row, 11).setValue(false);
      Logger.log(`[${new Date().toLocaleString()}] Отметка для "${itemName}" обновлена на листе "База" в столбце ${baseColumn}.`);
    } else {
      Logger.log(`[${new Date().toLocaleString()}] Изделие "${itemName}" не найдено на листе "База".`);
    }
  return itemId;
  }

}



