///*********************** Глобальные Константы *************************/

/**
 * ВАЖНО: номера ниже относятся только к стабильному пользовательскому интерфейсу
 * листов исполнителей A:L. Лист «База» НЕ имеет фиксированных COL_*: его поля
 * разрешаются по field_key через TableArchitecture.gs.
 */
const ROLE_COL_DATE             = 1;  // A
const ROLE_COL_DRAWING          = 2;  // B
const ROLE_COL_ORDER_LINK       = 3;  // C
const ROLE_COL_UUID             = 4;  // D
const ROLE_COL_ITEMNAME         = 5;  // E
const ROLE_COL_DESTINATION      = 6;  // F
const ROLE_COL_NOTE             = 7;  // G
const ROLE_COL_QUANTITY         = 8;  // H
const ROLE_COL_STOCK            = 9;  // I
const ROLE_COL_DEFICIT          = 10; // J
const COL_CHECKBOX              = 11; // K
const COL_RETURN_CHECKBOX       = 12; // L
const COL_LAUNCH_TRIGGER        = 5;  // E на листе «Запуск»

// Служебная проекция листов исполнителей находится далеко справа и скрыта.
const ROLE_HELPER_FIRST_COLUMN = 53; // BA

// «Каталог ссылок».
const CATALOG_COL_ITEMNAME = 1;
const CATALOG_COL_LINK = 3;

// Таблицы самого Диспетчера. Среда определяется по ID, а не по имени файла.
const DISPATCHER_PRODUCTION_SPREADSHEET_ID = '1JTr3AFm5Q1ycipTsrC5qY-SGKWY3v_w_2Jgfhq2xHgU';
const DISPATCHER_TEST_SPREADSHEET_ID = '1qcl4mumcHdxUmWCV4Cyl0eAb1bK4vTJ2GVm4EyPdQeQ';
const TEST_EXTERNAL_MARKER = '[TEST Dispatcher 3.1]';

// Внешние таблицы. Все ID собраны здесь, чтобы они не были размазаны по handlers.
const MAIN_CATALOG_SHEET_ID = DISPATCHER_PRODUCTION_SPREADSHEET_ID;
const SUPPLY_SPREADSHEET_ID = '1j_5fbonKWN5O5F7EeHDRjanm5d-fhWR0K2m2SO_EvSc';
const ARCHIVE_SPREADSHEET_ID = '1oPsmDHUGmN0ZmzYbEMp137Px_KsxDcZRx4PhceUp-3g';
const PRICE_DB_SPREADSHEET_ID = '1wR3wybH7UplH0ogrrELf1S27N9O5t5DS9lIkBHeagJ8';

function getDispatcherEnvironment_(ss) {
  const spreadsheet = ss || getDispatcherRuntimeSpreadsheet_();
  const id = spreadsheet && spreadsheet.getId ? spreadsheet.getId() : '';
  if (id === DISPATCHER_PRODUCTION_SPREADSHEET_ID) return 'PROD';
  if (id === DISPATCHER_TEST_SPREADSHEET_ID) return 'TEST';
  // Любая неизвестная копия считается тестовой: безопаснее явно пометить внешнее действие.
  return 'TEST';
}

function isTestDispatcher_(ss) {
  return getDispatcherEnvironment_(ss) !== 'PROD';
}

function markExternalTestText_(ss, text) {
  const value = String(text || '').trim();
  if (!isTestDispatcher_(ss)) return value;
  return value ? `${TEST_EXTERNAL_MARKER} ${value}` : TEST_EXTERNAL_MARKER;
}

// Firestore.
const PRODUCTS_COLLECTION = 'Products';
const MAX_SPEC_DEPTH = 10;
let _firestoreConfigCache = null;

/**
 * Читает учётные данные Firestore только тогда, когда Firestore действительно
 * используется. Обычные onEdit, не связанные с Firestore, ScriptProperties не читают.
 */
function getFirestoreConfig_() {
  if (_firestoreConfigCache) return _firestoreConfigCache;

  const props = PropertiesService.getScriptProperties();
  const clientEmail = props.getProperty('client_email');
  const privateKeyRaw = props.getProperty('private_key');
  const projectId = props.getProperty('project_id');

  if (!clientEmail || !privateKeyRaw || !projectId) {
    throw new Error('Не заданы обязательные Script Properties для Firestore.');
  }

  _firestoreConfigCache = {
    clientEmail,
    privateKey: privateKeyRaw.replace(/\\n/g, '\n'),
    projectId
  };
  return _firestoreConfigCache;
}
