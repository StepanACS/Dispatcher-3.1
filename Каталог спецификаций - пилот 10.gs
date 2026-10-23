/**
 * Пилот индексации 10 основных спецификаций для "Диспетчер 3.0".
 * Вход: "Каталог ссылок" (A = артикул, B = ссылка на папку КД).
 * Читаются только Google-таблицы непосредственно в этих папках:
 * подпапки, включая "Архив", никогда не обходятся.
 *
 * Ручные функции: previewSpecificationCatalogPrototype10 (без записи),
 * runSpecificationCatalogPrototype10 (перезапись только тестового листа).
 * Триггеров нет, другие модули диспетчера не вызываются.
 */

const SPEC_PILOT_ = {
  id: '1XffWoUclgQyFcT4DOIRoenQ9jAbBxXam6YxyY8xgv_4',
  catalog: 'Каталог ссылок',
  result: 'Спецификации (тест 10)',
  articles: ['Р350', 'Р351', 'Р352', 'Р358', 'Р359',
             'Р361', 'Р336', 'Р337', 'Р275', 'Р276']
};
const SPEC_PILOT_HEADERS_ = [
  'Артикул изделия', 'Поз.', 'Обозначение', 'Наименование', 'Кол-во',
  'Категория', 'Спецификация', 'Ссылка на СП', 'Ссылка на КД',
  'Лист СП', 'Строка СП', 'Резиновое кольцо', 'Доп. поле спецификации'
];

/** Прочитать и проверить 10 основных СП, ничего не записывая. */
function previewSpecificationCatalogPrototype10() {
  const result = specPilotCollect_();
  Logger.log(JSON.stringify(result.stats));
  return result.stats;
}

/** Ручной запуск. Полная замена только тестового листа. */
function runSpecificationCatalogPrototype10() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    // Сначала проверяем все источники, чтобы не стереть старый
    // результат при недоступной/неоднозначной спецификации.
    const result = specPilotCollect_();
    const ss = SpreadsheetApp.openById(SPEC_PILOT_.id);
    let sheet = ss.getSheetByName(SPEC_PILOT_.result);
    if (!sheet) sheet = ss.insertSheet(SPEC_PILOT_.result);
    const data = [SPEC_PILOT_HEADERS_].concat(result.rows);
    const width = SPEC_PILOT_HEADERS_.length;

    if (sheet.getMaxColumns() < width) {
      sheet.insertColumnsAfter(sheet.getMaxColumns(), width - sheet.getMaxColumns());
    }
    if (sheet.getMaxRows() < data.length) {
      sheet.insertRowsAfter(sheet.getMaxRows(), data.length - sheet.getMaxRows());
    }

    // Только создаваемый нами лист, никаких изменений "Каталога ссылок".
    if (sheet.getLastRow()) {
      sheet.getRange(1, 1, sheet.getLastRow(), width).clearContent();
    }
    sheet.getRange(1, 1, data.length, width).setNumberFormat('@');
    if (data.length > 1) {
      sheet.getRange(2, 5, data.length - 1, 1).setNumberFormat('0.###');
    }
    sheet.getRange(1, 1, data.length, width).setValues(data);
    sheet.setFrozenRows(1);
    sheet.setHiddenGridlines(true);
    sheet.getRange(1, 1, 1, width)
      .setBackground('#E8EDF4').setFontColor('#243247')
      .setFontWeight('bold').setWrap(true);
    sheet.setRowHeight(1, 38);
    [135, 70, 230, 265, 80, 160, 190,
     320, 320, 120, 100, 145, 135].forEach(
      (pixels, index) => sheet.setColumnWidth(index + 1, pixels)
    );
    const oldFilter = sheet.getFilter();
    if (oldFilter) oldFilter.remove();
    sheet.getRange(1, 1, data.length, width).createFilter();

    Logger.log('SPEC_PILOT_OK ' + JSON.stringify(result.stats));
    return result.stats;
  } finally {
    lock.releaseLock();
  }
}

function specPilotCollect_() {
  const ss = SpreadsheetApp.openById(SPEC_PILOT_.id);
  const sheet = ss.getSheetByName(SPEC_PILOT_.catalog);
  if (!sheet) throw new Error('Не найден лист "Каталог ссылок"');
  const count = sheet.getLastRow() - 1;
  if (count < 1) throw new Error('Каталог ссылок пуст');
  const links = sheet.getRange(2, 1, count, 2).getDisplayValues();
  const folders = {};

  links.forEach(function(row) {
    const article = String(row[0] || '').trim();
    if (SPEC_PILOT_.articles.indexOf(article) < 0) return;
    const found = String(row[1] || '').match(/\/folders\/([a-zA-Z0-9_-]+)/);
    if (!found) return;
    if (folders[article] && folders[article] !== found[1]) {
      throw new Error(article + ': в каталоге неоднозначные ссылки');
    }
    folders[article] = found[1];
  });

  const output = [];
  const stats = { documents: 0, parts: 0, rings: 0, byArticle: {} };

  SPEC_PILOT_.articles.forEach(function(article) {
    const folderId = folders[article];
    if (!folderId) throw new Error(article + ': отсутствует ссылка на КД в B');
    const folder = DriveApp.getFolderById(folderId);
    const candidates = [];
    const files = folder.getFiles();  // Никакой рекурсии по подпапкам!
    const nameRegex = new RegExp('^' + article + '\\s+СП\\s+\\d', 'iu');
    while (files.hasNext()) {
      const file = files.next();
      if (file.getMimeType() !== MimeType.GOOGLE_SHEETS) continue;
      if (nameRegex.test(file.getName().trim())) candidates.push(file);
    }
    if (candidates.length !== 1) {
      throw new Error(article + ': основных СП в корне ' + candidates.length +
                      ', ожидается ровно одна');
    }

    const file = candidates[0];
    const source = SpreadsheetApp.openById(file.getId());
    let specSheet = null;
    source.getSheets().forEach(function(tab) {
      const top = tab.getRange(1, 1, 3, 5).getDisplayValues();
      const declared = String(top[0][2] || '').trim();
      if (declared !== article && declared !== article + ' СБ') return;
      if (!/Поз/i.test(top[2][0]) || !/Обознач/i.test(top[2][1]) ||
          !/Наимен/i.test(top[2][2])) return;
      if (!specSheet || tab.getName().toUpperCase() === 'СП') specSheet = tab;
    });
    if (!specSheet) throw new Error(article + ': основная таблица не найдена');

    const last = Math.min(Math.max(3, specSheet.getLastRow()), 1000);
    const cells = specSheet.getRange(1, 1, last, 7).getDisplayValues();
    const header = cells.findIndex(row =>
      /Поз/i.test(String(row[0] || '')) &&
      /Обознач/i.test(String(row[1] || '')) &&
      /Наимен/i.test(String(row[2] || '')));
    if (header < 0) throw new Error(article + ': нет шапки таблицы');

    let parts = 0;
    for (let i = header + 1; i < cells.length; i++) {
      const [position, designation, title, qtyText, category, usage] =
        [0, 1, 2, 3, 4, 5].map(n => String(cells[i][n] || '').trim());
      if (/Обозначение изделия|Аннулированные сборочные|Исполнение с раскроем/i.test(position)) break;
      if (!/^\d+$/.test(position) || !qtyText || (!designation && !title)) continue;
      // Служебная строка с номерами столбцов, статус аннулированного задела.
      if ((designation === '2' && title === '3') || /^Списать$/i.test(usage)) continue;

      const quantity = Number(qtyText.replace(/\s/g, '').replace(',', '.'));
      if (!Number.isFinite(quantity)) {
        throw new Error(article + ': ошибочное количество в строке ' + (i + 1));
      }
      const ring = /^\d{3}[-‐‑–—−]\d{3}[-‐‑–—−]\d{2}$/u.test(designation);
      output.push([
        article, position, designation, title, quantity, category,
        file.getName(), 'https://docs.google.com/spreadsheets/d/' +
          file.getId() + '/edit',
        'https://drive.google.com/drive/folders/' + folderId,
        specSheet.getName(), i + 1, ring ? 'Да' : '', usage
      ]);
      parts++;
      if (ring) stats.rings++;
    }
    if (!parts) throw new Error(article + ': отсутствуют строки состава');
    stats.byArticle[article] = parts;
    stats.documents++;
    stats.parts += parts;
  });

  if (stats.documents !== SPEC_PILOT_.articles.length) {
    throw new Error('Прототип прочитал не все 10 спецификаций');
  }
  return { rows: output, stats: stats };
}
