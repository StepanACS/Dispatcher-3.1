///*********************** Глобальные Константы *************************/
// Номера столбцов листа «База» (1-индексация).
const COL_DATE                    = 1;  // A
const COL_DRAWING                 = 2;  // B
const COL_LINK                    = 3;  // C
const COL_UUID                    = 4;  // D
const COL_ITEMNAME                = 5;  // E
const COL_DESTINATION             = 6;  // F
const COL_NOTE                    = 7;  // G
const COL_QUANTITY                = 8;  // H
const COL_IN_STOCK                = 9;  // I
const COL_DEFICIT                 = 10; // J
const COL_READY_FOR_PRODUCTION    = 11; // K
const COL_IS_NEED                 = 12; // L
const COL_DESIGN                  = 13; // M
const COL_TECHNICAL_CONTROL       = 14; // N
const COL_DEPLOYMENT              = 15; // O
const COL_TECHNOLOGIEST           = 16; // P
const COL_TECHNOLOGY_CHECK        = 17; // Q
const COL_PICKER                  = 18; // R
const COL_PICKING                 = 19; // S
const COL_STOCK                   = 20; // T
const COL_PRODUCTION_STOCK        = 21; // U
const COL_PRODUCTION              = 22; // V
const COL_QUALITY_CHECK           = 23; // W
const COL_ACCEPTANCE              = 24; // X
const COL_READY                   = 25; // Y
const COL_STATUS                  = 26; // Z
const COL_MODE                    = 27; // AA
const COL_SDID                    = 28; // AB
const COL_TRELLO_ID               = 29; // AC
const COL_SD_ORDER_ID             = 30; // AD
const COL_FACT_QTY                = 31; // AE
const COL_TECHNOLOGIEST_NAME      = 32; // AF
const COL_DESIGNER_NAME           = 33; // AG
const COL_SUPPLY_ORDERED          = 34; // AH
const COL_TIED_CAPITAL            = 35; // AI
const COL_QTY_SD                  = 36; // AJ
const COL_ABC_CATEGORY            = 37; // AK

/**
 * ВАЖНО: в текущей «Диспетчер 3.1 опытный» заголовки AL:AN уже не совпадают
 * с историческими семантическими именами этих констант. Значения оставлены
 * неизменными, чтобы этот performance-refactor не менял формульную/табличную
 * архитектуру. Согласовать AL:AN нужно отдельным рефакторингом таблицы.
 */
const COL_PRIORITY                = 38; // AL — историческое назначение
const COL_EQIPMENT                = 39; // AM — историческое назначение
const COL_DEADLINE                = 40; // AN — историческое назначение
var LAST_BASE_COLUMN              = 40;

// Универсальные столбцы на листах исполнителей.
const COL_CHECKBOX                = 11; // K
const COL_RETURN_CHECKBOX         = 12; // L
const COL_LAUNCH_TRIGGER          = 5;  // E на листе «Запуск»

// «Каталог ссылок».
const CATALOG_COL_ITEMNAME = 1;
const CATALOG_COL_LINK = 3;

// Внешние таблицы. Все ID собраны здесь, чтобы они не были размазаны по handlers.
const MAIN_CATALOG_SHEET_ID = '1JTr3AFm5Q1ycipTsrC5qY-SGKWY3v_w_2Jgfhq2xHgU';
const SUPPLY_SPREADSHEET_ID = '1j_5fbonKWN5O5F7EeHDRjanm5d-fhWR0K2m2SO_EvSc';
const ARCHIVE_SPREADSHEET_ID = '1oPsmDHUGmN0ZmzYbEMp137Px_KsxDcZRx4PhceUp-3g';
const PRICE_DB_SPREADSHEET_ID = '1wR3wybH7UplH0ogrrELf1S27N9O5t5DS9lIkBHeagJ8';

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
