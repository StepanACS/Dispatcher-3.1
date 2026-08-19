///*********************** Глобальные Константы *************************/
// Эти константы определяют номера столбцов в листе База (используется 1-индексация)
const COL_DATE                    = 1;  // Столбец A – Дата создания записи
const COL_DRAWING                 = 2;  // Столбец B – Ссылка на КД
const COL_LINK                    = 3;  // Столбец C – Ссылка на наряд
const COL_UUID                    = 4;  // Столбец D – Уникальный идентификатор (UUID)
const COL_ITEMNAME                = 5;  // Столбец E – Название изделия (ключовое поле)
const COL_DESTINATION             = 6;  // Столбец F – Назначение
const COL_NOTE                    = 7;  // Столбец G – Примечание
const COL_QUANTITY                = 8;  // Столбец H – Количество изделий
const COL_IN_STOCK                = 9;  // Столбец I – Количество на складе
const COL_DEFICIT                 = 10; // Столбец J – Дефицит
const COL_READY_FOR_PRODUCTION    = 11; // Столбец K – Готовность к производству
const COL_IS_NEED                 = 12; // Столбец L – Есть ли потребность
const COL_DESIGN                  = 13; // Столбец M – Наличие КД
const COL_TECHNICAL_CONTROL       = 14; // Столбец N – Отметка о выполнении технического контроля после проектирования
const COL_DEPLOYMENT              = 15; // Столбец O – Отметка о развертывании спецификации на листе "База"
const COL_TECHNOLOGIEST           = 16; // Столбец P – Отметка технолога
const COL_TECHNOLOGY_CHECK        = 17; // Столбец Q – Отметка о проверке технологии
const COL_PICKER                  = 18; // Столбец R – Отметка о выполнении задачи диспетчера - комплектовщика на листе "База"
const COL_PICKING                 = 19; // Столбец S – Отметка о комплектации заказа
const COL_STOCK                   = 20; // Столбец T – Отметка о выполнении задачи склада на листе "База"
const COL_PRODUCTION_STOCK        = 21; // Столбец U – Отметка о выдаче зарезервированных товаров в производство
const COL_PRODUCTION              = 22; // Столбец V – Отметка о выполнении задач производства
const COL_QUALITY_CHECK           = 23; // Столбец W – Отметка о выполнении ОТК
const COL_ACCEPTANCE              = 24; // Столбец X – Отметка о приемке изделий с фиксацией фактического количества
const COL_READY                   = 25; // Столбец Y – Отметка о передаче в архив (Готово) на листе "База"
const COL_STATUS                  = 26; // Столбец Z – Статус - на каком этапе обработки находится строка
const COL_MODE                    = 27; // Столбец AA – Способ получения ("В производство" или "В снабжение")
const COL_SDID                    = 28; // Столбец AB – Уникальный идентификатор из системы SalesDrive
const COL_TRELLO_ID               = 29; // Столбец AC – Уникальный идентификатор карточки в Trello
const COL_SD_ORDER_ID             = 30; // Столбец AD – Уникальный идентификатор заказа в системе SalesDrive
const COL_FACT_QTY                = 31; // Столбец AE – Фактически принятое количество

const COL_TECHNOLOGIEST_NAME      = 32; // Столбец AF – Имя технолога
const COL_DESIGNER_NAME           = 33; // Столбец AG – Имя конструктора
const COL_SUPPLY_ORDERED          = 34; // Столбец AH – Заказано в снабжении
const COL_TIED_CAPITAL            = 35; // Столбец AI – Связанный капитал
const COL_QTY_SD                  = 36; // Столбец AJ – Количество из СД
const COL_ABC_CATEGORY            = 37; // Столбец AK – Категория из ABC анализа
const COL_PRIORITY                = 38; // Столбец AL – Приоритет в производстве
const COL_EQIPMENT                = 39; // Столбец AM – Оснастка
const COL_DEADLINE                = 40; // Столбец AN – Срок сдачи продукции

var LAST_BASE_COLUMN              = 40; // Столбец AN – Последний используемый столбец

// Универсальные столбцы с чекбоксами на страницах исполнителей
const COL_CHECKBOX                = 11; // Столбец K – Отметка о готовности
const COL_RETURN_CHECKBOX         = 12; // Столбец L – Возврат на предыдущий этап
const COL_LAUNCH_TRIGGER          = 5;  // Столбец E – запуск переноса с листа «Запуск»

// Лист «Каталог ссылок»
const CATALOG_COL_ITEMNAME = 1;
const CATALOG_COL_LINK = 3;

// Firestore
const PRODUCTS_COLLECTION = 'Products';
let _firestoreConfigCache = null;

/**
 * Читает учётные данные Firestore только тогда, когда Firestore действительно
 * используется. Раньше три обращения к ScriptProperties выполнялись при каждом
 * запуске проекта, включая обычные onEdit, не связанные с Firestore.
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

const MAX_SPEC_DEPTH = 10;
const MAIN_CATALOG_SHEET_ID = '1JTr3AFm5Q1ycipTsrC5qY-SGKWY3v_w_2Jgfhq2xHgU';
