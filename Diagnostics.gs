/**
 * Diagnostics.gs
 * -----------------------------------------------------------------------------
 * Ручные тесты и диагностические утилиты проекта «Диспетчер 3.1».
 * Эти функции не участвуют в штатной обработке onEdit и запускаются вручную
 * из редактора Apps Script при проверке отдельных подсистем.
 *
 * ВАЖНО: функции, которые изменяют данные, помечены соответствующим комментарием.
 */

/**
 * Проверяет ensureFreeRows() на листе «Запуск» активной таблицы.
 * Может добавлять строки в лист, поэтому запускайте только на тестовой/опытной копии.
 */
function testEnsureFreeRows() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Запуск');
  if (!sheet) throw new Error('Лист «Запуск» не найден.');
  const result = ensureFreeRows(sheet, 18, 30);
  Logger.log(JSON.stringify(result));
}

/**
 * Проверяет чтение размеров листа «База» активной таблицы.
 * Данные не изменяет.
 */
function sheetTest() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('База');
  if (!sheet) throw new Error('Лист «База» не найден.');
  Logger.log('Количество строк на листе: ' + sheet.getMaxRows());
}

/**
 * Проверяет расчёт дефицита для тестовой позиции.
 * Данные не изменяет.
 */
function testDeficit() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log(calculateDeficit('НП6х28Ч', 10, ss));
}

/**
 * Ручной тест загрузки спецификации в Firestore.
 * ВНИМАНИЕ: может создавать/обновлять документы Firestore.
 */
function testUploadSpec() {
  Logger.log(uploadSpecToFirestoreFromSheet('Б4004'));
}

/**
 * Полный smoke-test библиотеки Firestore: создаёт тестовый документ,
 * обновляет, читает и затем удаляет его.
 * Использует отдельную коллекцию test_collection_from_gas.
 */
function testFirestoreLib() {
  Logger.log('Начинаем тест подключения к Firestore...');

  try {
    const db = getFirestoreDb();
    const collectionName = 'test_collection_from_gas';
    const docId = 'test_doc_' + Date.now();
    const fullDocPath = collectionName + '/' + docId;

    db.createDocument(fullDocPath, {
      message: 'Hello from Apps Script diagnostics',
      timestamp: new Date()
    });

    db.updateDocument(fullDocPath, {
      updated_field: 'This document was updated',
      timestamp_updated: new Date()
    }, true);

    const retrievedDoc = db.getDocument(fullDocPath);
    Logger.log('Содержимое тестового документа: ' + JSON.stringify(retrievedDoc.obj));

    db.deleteDocument(fullDocPath);
    Logger.log('Тест Firestore успешно завершён, тестовый документ удалён.');
  } catch (error) {
    Logger.log('Ошибка тестирования Firestore: ' + error.message);
    if (error.stack) Logger.log(error.stack);
  }
}

/**
 * Проверяет авторизацию Trello текущими учётными данными проекта.
 * Данные Trello не изменяет.
 */
function trelloAuthCheck() {
  const credentials = getTrelloCredentials_();
  const url = 'https://api.trello.com/1/members/me?key=' +
    encodeURIComponent(credentials.apiKey) + '&token=' +
    encodeURIComponent(credentials.token);

  const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  Logger.log('Trello HTTP status: ' + response.getResponseCode());
  Logger.log(response.getContentText());
}

/**
 * Проверяет, что ссылка в «Каталог ссылок» ведёт в папку, имя которой
 * совпадает с корневым кодом изделия из столбца A.
 * Работает с активной таблицей и ничего в ней не изменяет.
 */
function checkCatalogFolderLinks() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Каталог ссылок');
  if (!sheet) throw new Error('Лист «Каталог ссылок» не найден.');

  const startRow = 2;
  const lastRow = sheet.getLastRow();
  if (lastRow < startRow) {
    Logger.log('[ok] Нет данных для проверки.');
    return;
  }

  const values = sheet.getRange(startRow, 1, lastRow - startRow + 1, 2).getValues();
  let checked = 0;
  let mismatches = 0;
  let skippedEmpty = 0;
  let skippedBadLink = 0;

  values.forEach((row, index) => {
    const rowNumber = startRow + index;
    const productRaw = String(row[0] ?? '').trim();
    const linkRaw = String(row[1] ?? '').trim();
    if (!productRaw) return;
    if (!linkRaw) {
      skippedEmpty++;
      return;
    }

    const expected = normalizeProductToRoot_(productRaw);
    const folderId = extractDriveFolderId_(linkRaw);
    if (!folderId) {
      skippedBadLink++;
      Logger.log(`[warn] Не удалось извлечь ID папки. row=${rowNumber}, изделие="${productRaw}"`);
      return;
    }

    try {
      const folderName = DriveApp.getFolderById(folderId).getName().trim();
      checked++;
      if (folderName !== expected) {
        mismatches++;
        Logger.log(`[mismatch] изделие="${productRaw}", ожидалось="${expected}", папка="${folderName}", row=${rowNumber}`);
      }
    } catch (error) {
      skippedBadLink++;
      Logger.log(`[warn] Ошибка доступа к папке. row=${rowNumber}, folderId="${folderId}", err="${error.message}"`);
    }
  });

  Logger.log(`[done] checked=${checked}, mismatches=${mismatches}, skippedEmpty=${skippedEmpty}, skippedBadLink=${skippedBadLink}`);
}

function normalizeProductToRoot_(productCode) {
  const value = String(productCode ?? '').trim();
  const dotPos = value.indexOf('.');
  return dotPos === -1 ? value : value.slice(0, dotPos).trim();
}

function extractDriveFolderId_(input) {
  const value = String(input ?? '').trim();
  if (!value) return '';
  if (!value.includes('/') && !value.includes('http') && value.length >= 20) return value;

  let match = value.match(/\/folders\/([a-zA-Z0-9_-]+)/);
  if (match && match[1]) return match[1];

  match = value.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  if (match && match[1]) return match[1];

  match = value.match(/([a-zA-Z0-9_-]{20,})/);
  return match && match[1] ? match[1] : '';
}

/**
 * Ручной диагностический запуск removeSpecification().
 * ВНИМАНИЕ: функция удаляет/изменяет строки «База».
 * Чтобы исключить случайный запуск с зашитым UUID, целевой UUID берётся из
 * Script Property `diagnostic_remove_spec_uuid`.
 */
function startRemoveSpec() {
  const parentUuid = PropertiesService.getScriptProperties()
    .getProperty('diagnostic_remove_spec_uuid');

  if (!parentUuid) {
    throw new Error('Сначала задайте Script Property diagnostic_remove_spec_uuid.');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  removeSpecification(parentUuid, ss);
  ss.toast('Диагностический откат спецификации завершён.', 'Информация', 5);
}
