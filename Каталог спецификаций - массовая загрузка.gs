/**
 * Единый каталог состава изделий: первичная массовая загрузка.
 * Самостоятельный модуль; не изменяет логику Диспетчера 3.0.
 *
 * Источник: "Каталог ссылок", A=артикул, B=папка КД.
 * Только основные изделия (без точки в артикуле). Только файлы корня
 * папки КД. В "Архив" и другие вложенные каталоги НЕ заходим.
 *
 * Выполнить вручную:
 *   startFullSpecificationImport()   - начать новую полную загрузку
 *   getFullSpecificationImportStatus() - проверить прогресс
 *   resumeFullSpecificationImport()  - возобновить после PAUSED/STOPPED
 *   stopFullSpecificationImport()    - остановить; результат сохранится
 *   testFullSpecificationImportParser() - локальные тесты парсера
 *
 * Обработка продолжается однократными таймерными триггерами, которые
 * удаляются после окончания. Рабочие листы Диспетчера не затрагиваются.
 * ВНИМАНИЕ: это первичный снимок, не непрерывная синхронизация с КД.
 */
const SPEC_ALL_ = Object.freeze({
  spreadsheetId: '1XffWoUclgQyFcT4DOIRoenQ9jAbBxXam6YxyY8xgv_4',
  catalogName: 'Каталог ссылок',
  stagingName: 'Каталог спецификаций - загрузка',
  finishedName: 'Каталог спецификаций',
  journalName: 'Импорт СП - журнал',
  stateKey: 'SPEC_FULL_CATALOG_IMPORT_V1',
  handler: 'continueFullSpecificationImport',
  maxItemsPerExecution: 30,
  timeBudgetMs: 210000,
  maxSourceRows: 600
});

const SPEC_ALL_HEADERS_ = [
  'Артикул изделия', 'Поз.', 'Обозначение', 'Наименование', 'Кол-во',
  'Категория', 'Спецификация', 'Ссылка на СП', 'Ссылка на КД',
  'Лист СП', 'Строка СП', 'Резиновое кольцо', 'Доп. поле спецификации'
];
const SPEC_ALL_JOURNAL_HEADERS_ = [
  'Артикул', 'Ссылка на КД', 'Статус', 'Файл СП',
  'Строк состава', 'Колец', 'Комментарий', 'Проверено'
];

/** Создать снимок очереди и приступить к обработке первой порции. */
function startFullSpecificationImport() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error('Импорт уже выполняется');
  try {
    const ss = SpreadsheetApp.openById(SPEC_ALL_.spreadsheetId);
    const prev = specAllLoadState_();
    if (prev && prev.status !== 'FINISHED' && prev.status !== 'FINISHED_WITH_GAPS') {
      throw new Error('Есть незавершённый импорт; используйте resumeFullSpecificationImport()');
    }
    if (ss.getSheetByName(SPEC_ALL_.finishedName)) {
      throw new Error('Итоговый лист уже существует. Не перезаписываем его автоматически');
    }
    if (ss.getSheetByName(SPEC_ALL_.stagingName) ||
        ss.getSheetByName(SPEC_ALL_.journalName)) {
      throw new Error('Обнаружены листы предыдущей загрузки. Сначала проверьте их');
    }

    const source = ss.getSheetByName(SPEC_ALL_.catalogName);
    if (!source) throw new Error('Отсутствует лист "Каталог ссылок"');
    const snapshot = specAllBuildSnapshot_(source);
    if (!snapshot.queue.length) throw new Error('Нет подходящих ссылок на папки КД');

    const staging = ss.insertSheet(SPEC_ALL_.stagingName);
    specAllSetupSheet_(staging, SPEC_ALL_HEADERS_, [135, 70, 230, 265, 80,
      160, 190, 310, 310, 100, 90, 135, 195]);
    const journal = ss.insertSheet(SPEC_ALL_.journalName);
    specAllSetupSheet_(journal, SPEC_ALL_JOURNAL_HEADERS_,
      [150, 320, 175, 315, 95, 65, 420, 160]);
    specAllEnsureRows_(journal, snapshot.queue.length + 1);
    journal.getRange(2, 1, snapshot.queue.length, 8)
      .setValues(snapshot.queue);
    journal.getRange(2, 8, snapshot.queue.length, 1)
      .setNumberFormat('yyyy-mm-dd hh:mm:ss');

    const state = {
      status: 'RUNNING',
      cursor: 0,
      outputRow: 2,
      total: snapshot.queue.length,
      imported: 0,
      noSpec: 0,
      warning: 0,
      ringCount: 0,
      componentLinksExcluded: snapshot.componentLinksExcluded,
      withoutFolderExcluded: snapshot.withoutFolderExcluded,
      duplicateCatalogEntries: snapshot.duplicateCatalogEntries,
      startedAt: new Date().toISOString(),
      lastError: ''
    };
    specAllSaveState_(state);
    specAllWriteStatus_(journal, state);
  } finally {
    lock.releaseLock();
  }
  return continueFullSpecificationImport();
}

/** Worker: вызывается автоматически таймером или вручную для продолжения. */
function continueFullSpecificationImport() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) return { status: 'ALREADY_RUNNING' };
  const started = Date.now();
  let state = null;
  try {
    specAllRemoveTriggers_();
    state = specAllLoadState_();
    if (!state || state.status !== 'RUNNING') {
      throw new Error('Нет запущенной массовой загрузки');
    }

    const ss = SpreadsheetApp.openById(SPEC_ALL_.spreadsheetId);
    const stage = ss.getSheetByName(SPEC_ALL_.stagingName);
    const log = ss.getSheetByName(SPEC_ALL_.journalName);
    if (!stage || !log) throw new Error('Не найдены рабочие листы импорта');

    // Снимок очереди хранится в журнале: между запусками каталог может меняться.
    const remaining = state.total - state.cursor;
    const size = Math.min(SPEC_ALL_.maxItemsPerExecution, remaining);
    const input = size > 0
      ? log.getRange(state.cursor + 2, 1, size, 8).getValues() : [];
    const statuses = [];
    const output = [];
    let completed = 0;

    for (let i = 0; i < input.length; i++) {
      if (i > 0 && Date.now() - started > SPEC_ALL_.timeBudgetMs) break;
      const item = input[i].slice();
      const article = String(item[0] || '').trim();
      const url = String(item[1] || '').trim();
      let result;

      if (item[2] === 'CONFLICT' || item[2] === 'BAD_KD_LINK') {
        result = { status: item[2], rows: [], comment: String(item[6] || '') };
      } else {
        try {
          result = specAllReadArticle_(article, url);
        } catch (e) {
          if (specAllIsQuotaError_(e)) throw e;
          result = { status: 'ERROR', rows: [], comment: String(e.message || e) };
        }
      }

      const rows = result.rows || [];
      Array.prototype.push.apply(output, rows);
      item[2] = result.status;
      item[3] = result.fileUrl || '';
      item[4] = rows.length;
      item[5] = rows.filter(row => row[11] === 'Да').length;
      item[6] = String(result.comment || '').slice(0, 500);
      item[7] = new Date();
      statuses.push(item);
      completed++;

      if (result.status === 'IMPORTED') state.imported++;
      else if (result.status === 'NO_SPEC') state.noSpec++;
      else state.warning++;
      state.ringCount += item[5];
    }

    if (completed === 0 && remaining > 0) {
      throw new Error('Worker не обработал ни одного артикула');
    }

    // Фиксированные позиции записи: при повторе после прерывания один и
    // тот же блок переписывается поверх себя, а не дублируется внизу.
    if (output.length) {
      specAllEnsureRows_(stage, state.outputRow + output.length - 1);
      stage.getRange(state.outputRow, 3, output.length, 1)
        .setNumberFormat('@'); // сохранение ведущих нулей в кодах колец
      stage.getRange(state.outputRow, 1, output.length, 13).setValues(output);
    }
    if (statuses.length) {
      log.getRange(state.cursor + 2, 1, statuses.length, 8)
        .setValues(statuses);
    }

    SpreadsheetApp.flush();
    state.outputRow += output.length;
    state.cursor += completed;
    state.updatedAt = new Date().toISOString();
    specAllSaveState_(state);

    if (state.cursor < state.total) {
      specAllWriteStatus_(log, state);
      specAllScheduleTrigger_();
      Logger.log('SPEC_CATALOG_PROGRESS ' + JSON.stringify(state));
      return state;
    }

    // Публикуем готовый снимок только после проверки ВСЕЙ очереди.
    if (ss.getSheetByName(SPEC_ALL_.finishedName)) {
      throw new Error('Итоговый лист создан параллельно; публикация отменена');
    }
    stage.setName(SPEC_ALL_.finishedName);
    const last = Math.max(2, state.outputRow - 1);
    stage.getRange(1, 1, last, 13).createFilter();
    state.status = state.warning > 0 || state.noSpec > 0
      ? 'FINISHED_WITH_GAPS' : 'FINISHED';
    state.finishedAt = new Date().toISOString();
    specAllSaveState_(state);
    specAllWriteStatus_(log, state);
    specAllRemoveTriggers_();
    Logger.log('SPEC_CATALOG_FINISHED ' + JSON.stringify(state));
    return state;
  } catch (e) {
    if (state && state.status === 'RUNNING') {
      state.status = 'PAUSED';
      state.lastError = String(e.message || e).slice(0, 600);
      state.updatedAt = new Date().toISOString();
      specAllSaveState_(state);
      try {
        const journal = SpreadsheetApp.openById(SPEC_ALL_.spreadsheetId)
          .getSheetByName(SPEC_ALL_.journalName);
        if (journal) specAllWriteStatus_(journal, state);
      } catch (ignored) {}
    }
    specAllRemoveTriggers_();
    Logger.log('SPEC_CATALOG_PAUSED: ' + String(e.message || e));
    throw e;
  } finally {
    lock.releaseLock();
  }
}

/** Возобновление: сохранённые строки не пересчитываются. */
function resumeFullSpecificationImport() {
  const state = specAllLoadState_();
  if (!state || ['FINISHED', 'FINISHED_WITH_GAPS'].indexOf(state.status) >= 0) {
    throw new Error('Нет незавершённого импорта');
  }
  state.status = 'RUNNING';
  state.lastError = '';
  specAllSaveState_(state);
  return continueFullSpecificationImport();
}

/** Ручная остановка; журнал и уже записанные строки не удаляются. */
function stopFullSpecificationImport() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error('Идёт запись порции; повторите остановку');
  try {
    const state = specAllLoadState_();
    if (!state || state.status !== 'RUNNING') return state;
    state.status = 'STOPPED';
    specAllSaveState_(state);
    specAllRemoveTriggers_();
    return state;
  } finally {
    lock.releaseLock();
  }
}

/** Прогресс, в том числе из редактора Apps Script. */
function getFullSpecificationImportStatus() {
  const state = specAllLoadState_() || { status: 'NOT_STARTED' };
  Logger.log('SPEC_CATALOG_STATUS ' + JSON.stringify(state));
  return state;
}

function specAllBuildSnapshot_(catalog) {
  const last = catalog.getLastRow();
  if (last < 2) return { queue: [], componentLinksExcluded: 0,
    withoutFolderExcluded: 0, duplicateCatalogEntries: 0 };
  const range = catalog.getRange(2, 1, last - 1, 2);
  const display = range.getDisplayValues();
  const rich = catalog.getRange(2, 2, last - 1, 1).getRichTextValues();
  const seen = new Map();
  const queue = [];
  let componentLinksExcluded = 0;
  let withoutFolderExcluded = 0;
  let duplicateCatalogEntries = 0;

  display.forEach(function(r, index) {
    const article = String(r[0] || '').trim();
    const textUrl = String(r[1] || '').trim();
    const richUrl = rich[index][0] && rich[index][0].getLinkUrl();
    const url = /\/folders\//.test(textUrl) ? textUrl : (richUrl || textUrl);
    if (!article || !url) return;

    // Согласно ТЗ, отдельные детали и подсборки не индексируем.
    if (article.indexOf('.') >= 0) {
      componentLinksExcluded++;
      return;
    }
    const match = String(url).match(/\/folders\/([a-zA-Z0-9_-]+)/);
    if (!match) {
      withoutFolderExcluded++;
      return;
    }
    const folderId = match[1];
    if (seen.has(article)) {
      duplicateCatalogEntries++;
      const prior = queue[seen.get(article)];
      const oldId = String(prior[1]).match(/\/folders\/([a-zA-Z0-9_-]+)/);
      if (!oldId || oldId[1] !== folderId) {
        prior[2] = 'CONFLICT';
        prior[6] = 'Для артикула указаны разные папки КД';
      }
      return;
    }
    seen.set(article, queue.length);
    queue.push([article, url, 'PENDING', '', '', '', '', '']);
  });
  return { queue, componentLinksExcluded, withoutFolderExcluded,
    duplicateCatalogEntries };
}

/** Чтение только корня папки, выбор одной основной СП. */
function specAllReadArticle_(article, folderUrl) {
  const match = String(folderUrl).match(/\/folders\/([a-zA-Z0-9_-]+)/);
  if (!match) return { status: 'BAD_KD_LINK', rows: [], comment: 'Не папка Drive' };
  const resourceKey = String(folderUrl).match(/[?&]resourcekey=([a-zA-Z0-9_-]+)/);
  const folder = resourceKey
    ? DriveApp.getFolderByIdAndResourceKey(match[1], resourceKey[1])
    : DriveApp.getFolderById(match[1]);

  // В архивную ветку не заходим, даже если на неё ошибочно указывает B.
  if (specAllIsArchivePath_(folder)) {
    return { status: 'ARCHIVE_PATH', rows: [],
      comment: 'Папка КД находится в архивной ветке' };
  }
  const files = folder.getFiles(); // getFiles, НЕ getFolders, НЕ рекурсия
  const pattern = new RegExp('^' + specAllEscape_(article) + '\\s+СП\\s*\\d{6,8}(?:\\D|$)', 'iu');
  const sheets = [];
  const unsupported = [];
  while (files.hasNext()) {
    const file = files.next();
    if (!pattern.test(file.getName().trim())) continue;
    if (file.getMimeType() === MimeType.GOOGLE_SHEETS) sheets.push(file);
    else unsupported.push(file.getName() + ' [' + file.getMimeType() + ']');
  }
  if (sheets.length > 1) {
    return { status: 'AMBIGUOUS', rows: [],
      comment: 'Несколько основных СП: ' + sheets.map(f => f.getName()).join('; ') };
  }
  if (!sheets.length && unsupported.length) {
    return { status: 'UNSUPPORTED_FORMAT', rows: [],
      comment: unsupported.join('; ') };
  }
  if (!sheets.length) return { status: 'NO_SPEC', rows: [],
    comment: 'Основной спецификации с обозначением артикула в корне нет' };

  const file = sheets[0];
  const wb = SpreadsheetApp.openById(file.getId());
  const variants = [];
  wb.getSheets().forEach(function(tab) {
    const top = tab.getRange(1, 1, Math.min(15, tab.getMaxRows()), 5)
      .getDisplayValues();
    const declared = String((top[0] || [])[2] || '').trim().replace(/\s+/g, ' ');
    const expected = article.replace(/\s+/g, ' ');
    if (declared !== expected && declared !== expected + ' СБ') return;
    const header = top.findIndex(row => /Поз/i.test(String(row[0] || '')) &&
      /Обознач/i.test(String(row[1] || '')) &&
      /Наимен/i.test(String(row[2] || '')));
    if (header >= 0) variants.push({ tab, header });
  });
  if (!variants.length) return { status: 'INVALID_SOURCE', rows: [],
    fileUrl: file.getUrl(), comment: 'Не найдена таблица СП с совпадающим C1' };

  const selected = variants.find(v => v.tab.getName().toUpperCase() === 'СП') ||
    variants[0];
  if (selected.tab.getLastRow() > SPEC_ALL_.maxSourceRows) {
    return { status: 'SOURCE_TOO_LARGE', rows: [], fileUrl: file.getUrl(),
      comment: 'Слишком много строк в источнике; требуется ручная проверка' };
  }
  const matrix = selected.tab.getRange(1, 1,
    Math.max(3, selected.tab.getLastRow()), 7).getDisplayValues();
  const parsed = specAllParse_(matrix, selected.header, article,
    file.getName(), file.getUrl(), folderUrl, selected.tab.getName());
  if (!parsed.length) return { status: 'INVALID_SOURCE', rows: [],
    fileUrl: file.getUrl(), comment: 'В основной таблице нет элементов состава' };

  return { status: 'IMPORTED', rows: parsed, fileUrl: file.getUrl(), comment:
    'Лист: ' + selected.tab.getName() };
}

/** Запрет чтения содержимого Архива, включая ошибочные ссылки на его потомков. */
function specAllIsArchivePath_(folder) {
  let level = [folder];
  for (let depth = 0; depth < 12 && level.length; depth++) {
    const parents = [];
    for (let i = 0; i < level.length; i++) {
      const current = level[i];
      if (/(?:архив|archive)/iu.test(current.getName())) return true;
      const iterator = current.getParents();
      while (iterator.hasNext()) parents.push(iterator.next());
    }
    level = parents;
  }
  // Если дерево слишком глубокое, предпочтительнее пропустить, чем
  // рискнуть чтением из папки, лежащей под неизвестным архивом.
  return level.length > 0;
}

function specAllParse_(matrix, header, article, name, specUrl, folderUrl, tab) {
  const result = [];
  for (let i = header + 1; i < matrix.length; i++) {
    const fields = [0, 1, 2, 3, 4, 5]
      .map(j => String(matrix[i][j] == null ? '' : matrix[i][j]).trim());
    const [pos, code, title, qty, category, extra] = fields;
    if (/Обозначение изделия|Аннулированные сборочные|Исполнение с раскроем/i.test(pos)) {
      break;
    }
    if (!/^\d+$/.test(pos) || !qty || (!code && !title)) continue;
    if ((code === '2' && title === '3') || /^Списать$/i.test(extra)) continue;
    const count = Number(qty.replace(/\s/g, '').replace(',', '.'));
    if (!Number.isFinite(count) || count < 0) {
      throw new Error(article + ': неверное количество в строке ' + (i + 1));
    }
    const ring = /^\d{3}[-‐‑–—−]\d{3}[-‐‑–—−]\d{2}$/u.test(code) ||
      /(?:уплотнительн\w*\s+кольц|кольц\w*\s+уплотнительн|ущільнювальн\w*\s+кільц|кільц\w*\s+ущільнювальн)/iu.test(title);
    result.push([
      article, pos, code, title, count, category,
      name, specUrl, folderUrl, tab, i + 1, ring ? 'Да' : '', extra
    ]);
  }
  return result;
}

function specAllSetupSheet_(sheet, headers, widths) {
  if (sheet.getMaxColumns() < headers.length) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(),
      headers.length - sheet.getMaxColumns());
  }
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground('#E8EDF4').setFontColor('#243247')
    .setFontWeight('bold').setWrap(true);
  sheet.setRowHeight(1, 42);
  sheet.setFrozenRows(1);
  sheet.setHiddenGridlines(true);
  widths.forEach((value, index) => sheet.setColumnWidth(index + 1, value));
}

function specAllEnsureRows_(sheet, targetRows) {
  if (sheet.getMaxRows() < targetRows) {
    sheet.insertRowsAfter(sheet.getMaxRows(), targetRows - sheet.getMaxRows());
  }
}

function specAllWriteStatus_(log, state) {
  const summary = [
    ['Состояние', state.status], ['Обработано', state.cursor],
    ['В очереди', state.total], ['Импортировано СП', state.imported],
    ['Без СП', state.noSpec], ['Требуют проверки', state.warning],
    ['Колец', state.ringCount],
    ['Ссылок на детали исключено', state.componentLinksExcluded],
    ['Не папки КД исключено', state.withoutFolderExcluded],
    ['Повторяющиеся артикулы', state.duplicateCatalogEntries],
    ['Последняя ошибка', state.lastError || ''],
    ['Последнее обновление', state.updatedAt || state.startedAt]
  ];
  log.getRange(1, 10, summary.length, 2).setValues(summary);
  log.getRange(1, 10, 1, 2).setFontWeight('bold');
  log.setColumnWidth(10, 245);
  log.setColumnWidth(11, 380);
}

function specAllLoadState_() {
  const json = PropertiesService.getScriptProperties()
    .getProperty(SPEC_ALL_.stateKey);
  return json ? JSON.parse(json) : null;
}
function specAllSaveState_(value) {
  PropertiesService.getScriptProperties()
    .setProperty(SPEC_ALL_.stateKey, JSON.stringify(value));
}
function specAllRemoveTriggers_() {
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === SPEC_ALL_.handler) {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}
function specAllScheduleTrigger_() {
  specAllRemoveTriggers_();
  ScriptApp.newTrigger(SPEC_ALL_.handler).timeBased().after(60000).create();
}
function specAllIsQuotaError_(error) {
  return /quota|too many times|rate limit|429|limit exceeded|service invoked too many|execution time/i
    .test(String(error && (error.message || error)));
}
function specAllEscape_(text) {
  return String(text).replace(/[.*+?^$()|[\]{}\\]/g, '\\$&');
}

/** Быстрая проверка правил разбора: шапка, списание, повторная СП, кольцо. */
function testFullSpecificationImportParser() {
  const matrix = [
    ['Обозначение изделия:', '', 'Р999'],
    ['Наименование изделия:', '', 'Патрубок'],
    ['Поз. на СБ', 'Обозначение', 'Наименование', 'Кол-во', 'Категория'],
    ['1', '2', '3', '4', '5', '6'],
    ['1', '024-029-30', 'Кільце ущільнювальне', '2', 'Покупное изделие'],
    ['2', 'Р999.01', 'Деталь', '1', 'Деталь специальная'],
    ['3', 'Болт', 'Болт', '1', 'Покупное изделие', 'Списать'],
    ['Обозначение изделия:', '', 'Р999'],
    ['Поз. на СБ', 'Обозначение', 'Наименование', 'Кол-во', 'Категория'],
    ['1', '999-999-99', 'Кільце ущільнювальне', '9', 'Покупное изделие']
  ];
  const actual = specAllParse_(matrix, 2, 'Р999', 'Р999 СП 010126',
    'test://spec', 'test://folder', 'СП');
  if (actual.length !== 2 || actual[0][2] !== '024-029-30' ||
      actual[0][4] !== 2 || actual[0][11] !== 'Да' ||
      actual[1][2] !== 'Р999.01') {
    throw new Error('Тест парсера не пройден: ' + JSON.stringify(actual));
  }
  Logger.log('SPEC_CATALOG_PARSER_TESTS_PASSED');
  return 'PASS';
}
