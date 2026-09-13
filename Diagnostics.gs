/**
 * Diagnostics.gs
 * -----------------------------------------------------------------------------
 * Все ручные тесты и диагностические утилиты проекта «Диспетчер 3.1».
 * Они не участвуют в штатной обработке onEdit.
 */

/**
 * Главная проверка новой архитектуры таблицы. Ничего не изменяет.
 * Проверяет field_key -> header -> column, вычисляемые spill-якоря,
 * «Настройки» и наличие нормализованного реестра связей.
 */
function validateTableArchitecture() {
  const result = validateTableArchitecture_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (result.ok) {
    ss.toast(`Схема корректна: ${result.mappedFields} полей, ${result.computedFields} вычисляемых`, 'Диагностика', 8);
  } else {
    ss.toast(`Найдено проблем: ${result.issues.length}. Подробности в Execution log.`, 'Диагностика', 10);
    result.issues.forEach(issue => Logger.log('[ARCHITECTURE] ' + issue));
  }
  return result;
}

/** Может добавлять строки на «Запуск»; использовать только на тестовой копии. */
function testEnsureFreeRows() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Запуск');
  if (!sheet) throw new Error('Лист «Запуск» не найден.');
  Logger.log(JSON.stringify(ensureFreeRows(sheet, 18, 30)));
}

/** Читает размеры «База», ничего не изменяет. */
function sheetTest() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(BASE_SHEET_NAME);
  if (!sheet) throw new Error('Лист «База» не найден.');
  Logger.log(`Строк: ${sheet.getMaxRows()}, last UUID row: ${baseLastDataRow_(sheet)}`);
}

/** Проверяет расчёт дефицита. */
function testDeficit() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log(calculateDeficit('НП6х28Ч', 10, ss));
}

/** ВНИМАНИЕ: может создавать/обновлять документы Firestore. */
function testUploadSpec() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log(uploadSpecToFirestoreFromSheet('Б4004', ss));
}

/** Smoke-test Firestore с созданием и удалением отдельного тестового документа. */
function testFirestoreLib() {
  Logger.log('Начинаем тест подключения к Firestore...');
  try {
    const db = getFirestoreDb();
    const collectionName = 'test_collection_from_gas';
    const docId = 'test_doc_' + Date.now();
    const fullDocPath = collectionName + '/' + docId;
    db.createDocument(fullDocPath, { message: 'Hello from Apps Script diagnostics', timestamp: new Date() });
    db.updateDocument(fullDocPath, { updated_field: 'This document was updated', timestamp_updated: new Date() }, true);
    const retrievedDoc = db.getDocument(fullDocPath);
    Logger.log('Содержимое: ' + JSON.stringify(retrievedDoc.obj));
    db.deleteDocument(fullDocPath);
    Logger.log('Тест Firestore завершён.');
  } catch (error) {
    Logger.log('Ошибка Firestore: ' + error.message);
    if (error.stack) Logger.log(error.stack);
  }
}

/** Проверяет Trello-авторизацию, ничего не меняет. */
function trelloAuthCheck() {
  const credentials = getTrelloCredentials_();
  const url = 'https://api.trello.com/1/members/me?key=' +
    encodeURIComponent(credentials.apiKey) + '&token=' + encodeURIComponent(credentials.token);
  const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  Logger.log('Trello HTTP status: ' + response.getResponseCode());
  Logger.log(response.getContentText());
}

/** Проверяет соответствие ссылок «Каталог ссылок» именам корневых папок. */
function checkCatalogFolderLinks() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Каталог ссылок');
  if (!sheet) throw new Error('Лист «Каталог ссылок» не найден.');
  const startRow = 2;
  const lastRow = sheet.getLastRow();
  if (lastRow < startRow) { Logger.log('[ok] Нет данных для проверки.'); return; }

  const values = sheet.getRange(startRow, 1, lastRow - startRow + 1, 2).getValues();
  let checked = 0, mismatches = 0, skippedEmpty = 0, skippedBadLink = 0;
  values.forEach((r, index) => {
    const rowNumber = startRow + index;
    const productRaw = String(r[0] ?? '').trim();
    const linkRaw = String(r[1] ?? '').trim();
    if (!productRaw) return;
    if (!linkRaw) { skippedEmpty++; return; }
    const expected = normalizeProductToRoot_(productRaw);
    const folderId = extractDriveFolderId_(linkRaw);
    if (!folderId) { skippedBadLink++; Logger.log(`[warn] bad folder link row=${rowNumber}`); return; }
    try {
      const folderName = DriveApp.getFolderById(folderId).getName().trim();
      checked++;
      if (folderName !== expected) {
        mismatches++;
        Logger.log(`[mismatch] изделие="${productRaw}", ожидалось="${expected}", папка="${folderName}", row=${rowNumber}`);
      }
    } catch (error) {
      skippedBadLink++;
      Logger.log(`[warn] row=${rowNumber}, folderId="${folderId}", err="${error.message}"`);
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
 * ВНИМАНИЕ: очищает/изменяет записи «База» через removeSpecification().
 * Целевой UUID берётся из Script Property diagnostic_remove_spec_uuid.
 */
function startRemoveSpec() {
  const parentUuid = PropertiesService.getScriptProperties().getProperty('diagnostic_remove_spec_uuid');
  if (!parentUuid) throw new Error('Сначала задайте Script Property diagnostic_remove_spec_uuid.');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  removeSpecification(parentUuid, ss);
  ss.toast('Диагностический откат спецификации завершён.', 'Информация', 5);
}
