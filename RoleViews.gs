/**
 * Декларативная модель исполнительских листов.
 *
 * A:J — вычисляемая проекция «База»; K/L — ручные checkbox-поля; M+ —
 * дополнительные вычисляемые либо ручные поля. Единственный тяжёлый FILTER
 * каждого листа находится в скрытом блоке, начиная с BA.
 */
const ROLE_VIEW_COMMON_FIELDS = Object.freeze([
  'date', 'drawing_link', 'order_link', 'uuid', 'item_name',
  'destination', 'note', 'quantity', 'stock', 'deficit'
]);

const ROLE_VIEW_CONFIG = Object.freeze({
  'Проектирование': { extras: [] },
  'Техконтроль': { extras: [] },
  'Развертывание спецификации': {
    extras: ['abc_components_2', 'abc', 'responsible_technologist']
  },
  'Технолог': {
    extras: ['need_status', 'production_mode', 'stage_technical_control', 'abc_components_2']
  },
  'Проверка технологии': { extras: ['responsible_technologist'] },
  'Диспетчер- комплектовщик': {
    extras: ['need_status', 'production_mode', 'salesdrive_id', 'supply_ordered'],
    extraHeaders: { supply_ordered: 'Заказано' }
  },
  'Комплектация': {
    extras: ['need_status', 'production_mode', 'salesdrive_id'],
    derived: [{ key: '__completion_percent', header: 'Процент готовности' }]
  },
  'Склад': {
    extras: ['salesdrive_order_id', 'trello_id'],
    manual: [{ column: 15, header: 'Техпаспорт шаблон' }]
  },
  'Склад производства': { extras: ['salesdrive_order_id', 'trello_id'] },
  'Производство': { extras: ['salesdrive_order_id'] },
  'ОТК': { extras: [], manual: [{ column: 13, header: 'Кол-во готово' }] },
  'Приемка': { extras: [], manual: [{ column: 13, header: 'Принято фактически' }] },
  'Готово': {
    extras: ['actual_made', 'salesdrive_stock'],
    manual: [{ column: 16, header: 'Диспетчер 3.0 - архив выполненных' }]
  }
});

function columnNumberToLetter_(column) {
  let n = Number(column);
  if (!Number.isInteger(n) || n < 1) throw new Error(`Некорректный номер столбца: ${column}`);
  let result = '';
  while (n > 0) {
    n--;
    result = String.fromCharCode(65 + (n % 26)) + result;
    n = Math.floor(n / 26);
  }
  return result;
}

function roleViewFieldExpression_(fieldKey) {
  return `CHOOSECOLS(d;MATCH("${baseHeader_(getDispatcherRuntimeSpreadsheet_(), fieldKey)}";h;0))`;
}

function buildRoleViewSourceFormula_(ss, sheetName, config) {
  setDispatcherRuntimeSpreadsheet_(ss);
  const arch = getBaseArchitecture_(ss);
  const lastCol = columnNumberToLetter_(arch.lastColumn);
  const expressions = ROLE_VIEW_COMMON_FIELDS.map(roleViewFieldExpression_);
  (config.extras || []).forEach(key => expressions.push(roleViewFieldExpression_(key)));

  for (const derived of config.derived || []) {
    if (derived.key === '__completion_percent') {
      expressions.push(`IFERROR(${roleViewFieldExpression_('stock')}/${roleViewFieldExpression_('quantity')}*100;"")`);
    }
  }

  const statusHeader = baseHeader_(ss, 'status');
  return `=LET(h;'${BASE_SHEET_NAME}'!$A$1:$${lastCol}$1;d;'${BASE_SHEET_NAME}'!$A$2:$${lastCol};st;CHOOSECOLS(d;MATCH("${statusHeader}";h;0));IFNA(FILTER(HSTACK(${expressions.join(';')});st="${sheetName}");""))`;
}

function roleExtraHeader_(ss, config, fieldKey) {
  if (config.extraHeaders && config.extraHeaders[fieldKey]) return config.extraHeaders[fieldKey];
  return baseHeader_(ss, fieldKey);
}

/**
 * Пересобирает все исполнительские листы из semantic config.
 * K/L и явно перечисленные manual-поля не очищаются.
 */
function installRoleViews_(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  setDispatcherRuntimeSpreadsheet_(ss);
  clearTableArchitectureCache_(ss);
  getBaseArchitecture_(ss); // fail fast before any writes

  const helperFirstCol = ROLE_HELPER_FIRST_COLUMN; // BA
  const helperFirstLetter = columnNumberToLetter_(helperFirstCol);
  const checkboxRule = SpreadsheetApp.newDataValidation().requireCheckbox().build();

  for (const [sheetName, config] of Object.entries(ROLE_VIEW_CONFIG)) {
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) sheet = ss.insertSheet(sheetName);

    const totalSourceColumns = ROLE_VIEW_COMMON_FIELDS.length +
      (config.extras || []).length + (config.derived || []).length;
    const requiredCols = helperFirstCol + Math.max(totalSourceColumns, 18) - 1;
    if (sheet.getMaxColumns() < requiredCols) {
      sheet.insertColumnsAfter(sheet.getMaxColumns(), requiredCols - sheet.getMaxColumns());
    }

    // Common visible projection A:J.
    sheet.getRange(1, 1, 1, 10).setValues([ROLE_VIEW_COMMON_FIELDS.map(key => baseHeader_(ss, key))]);
    sheet.getRange(2, 1, Math.max(sheet.getMaxRows() - 1, 1), 10).clearContent();
    sheet.getRange('A2').setFormula(`=ARRAYFORMULA(${helperFirstLetter}2:${columnNumberToLetter_(helperFirstCol + 9)})`);
    sheet.getRange(2, 1, Math.max(sheet.getMaxRows() - 1, 1), 1).setNumberFormat('dd.MM.yyyy');

    // Checkbox interface remains physically stable by design.
    sheet.getRange(1, COL_CHECKBOX).setValue('Отметка о выполнении');
    const checkboxCount = getSettingsForSheet(sheetName)?.count || 1;
    if (checkboxCount > 1) sheet.getRange(1, COL_RETURN_CHECKBOX).setValue('Вернуть на предыдущий этап');
    else sheet.getRange(1, COL_RETURN_CHECKBOX).clearContent();
    sheet.getRange(2, COL_CHECKBOX, Math.max(sheet.getMaxRows() - 1, 1), checkboxCount).setDataValidation(checkboxRule);

    // Sourced extras start in M; clear only fields owned by the view.
    const extraHeaders = [];
    (config.extras || []).forEach(key => extraHeaders.push(roleExtraHeader_(ss, config, key)));
    (config.derived || []).forEach(item => extraHeaders.push(item.header));
    if (extraHeaders.length) {
      sheet.getRange(1, 13, 1, extraHeaders.length).setValues([extraHeaders]);
      sheet.getRange(2, 13, Math.max(sheet.getMaxRows() - 1, 1), extraHeaders.length).clearContent();
      const helperExtraStart = helperFirstCol + ROLE_VIEW_COMMON_FIELDS.length;
      const helperExtraEnd = helperExtraStart + extraHeaders.length - 1;
      sheet.getRange(2, 13).setFormula(
        `=ARRAYFORMULA(${columnNumberToLetter_(helperExtraStart)}2:${columnNumberToLetter_(helperExtraEnd)})`
      );
    }

    for (const manual of config.manual || []) sheet.getRange(1, manual.column).setValue(manual.header);

    // One hidden source formula per role sheet.
    const helperWidth = Math.max(totalSourceColumns, 18);
    sheet.getRange(1, helperFirstCol, sheet.getMaxRows(), helperWidth).clearContent();
    sheet.getRange(2, helperFirstCol).setFormula(buildRoleViewSourceFormula_(ss, sheetName, config));
    sheet.hideColumns(helperFirstCol, helperWidth);
  }
  SpreadsheetApp.flush();
}

/** Backward-compatible admin entry point; old letter-based QUERY generator removed. */
function checkAndUpdateSheets() {
  installRoleViews_(SpreadsheetApp.getActiveSpreadsheet());
}
