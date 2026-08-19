/**
 * Канонические формулы «База».
 *
 * Формулы генерируются по field_key и устанавливаются только в строку 2.
 * После установки Google Sheets самостоятельно поддерживает ссылки при
 * физической перестановке/вставке столбцов.
 */
function baseColumnLetter_(ss, fieldKey) {
  return columnNumberToLetter_(baseCol_(ss, fieldKey));
}

function baseColumnRange_(ss, fieldKey, startRow = 2) {
  const c = baseColumnLetter_(ss, fieldKey);
  return `${c}${startRow}:${c}`;
}

function buildBaseComputedFormulas_(ss) {
  const r = key => baseColumnRange_(ss, key);
  const item = r('item_name');
  const uuid = r('uuid');
  const destination = r('destination');
  const quantity = r('quantity');
  const stock = r('stock');
  const deficit = r('deficit');
  const needFlag = r('need_flag');
  const status = r('status');
  const mode = r('production_mode');
  const note = r('note');
  const abc = r('abc');
  const abc1 = r('abc_components');
  const abc2 = r('abc_components_2');

  const stageRanges = STAGE_FIELD_KEYS.map(key => r(key));

  return Object.freeze({
    drawing_link: `=ARRAYFORMULA(IF(${item}="";"";LET(link;IFNA(XLOOKUP(${item};'Каталог ссылок'!A:A;'Каталог ссылок'!B:B);"§MISS§");IF(link="§MISS§";"Добавить изделие в каталог ссылок";IF(link="";"Добавить ссылку в каталог";link)))))`,
    order_link: `=ARRAYFORMULA(IF(${item}="";"";LET(link;IFNA(XLOOKUP(${item};'Каталог ссылок'!A:A;'Каталог ссылок'!C:C);"§MISS§");IF(link="§MISS§";"Добавить изделие в каталог ссылок";IF(link="";"Добавить ссылку в каталог";link)))))`,
    stock: `=ARRAYFORMULA(IF(${item}="";"";IFNA(XLOOKUP(${item};'Остатки'!A:A;'Остатки'!B:B);0)))`,
    deficit: `=ARRAYFORMULA(IF(${item}="";"";IF(${quantity}-${stock}<0;0;${quantity}-${stock})))`,
    need_flag: `=ARRAYFORMULA(IF(${item}="";"";IF(${deficit}>0;"Да";"Нет")))`,
    status: `=BYROW(HSTACK(${item};${stageRanges.join(';')});LAMBDA(row;IF(INDEX(row;1;1)="";"";IFNA(INDEX('Настройки'!A$2:A$14;MATCH(FALSE;CHOOSECOLS(row;2;3;4;5;6;7;8;9;10;11;12;13;14);0));"Готово"))))`,
    production_mode: `=ARRAYFORMULA(IF(${item}="";"";IF(IFNA(XLOOKUP(${item};'Цена материалов'!A:A;'Цена материалов'!A:A);"")="";"В производство";"В снабжение")))`,
    salesdrive_id: `=ARRAYFORMULA(IF(${item}="";"";IFNA(XLOOKUP(${item};'SD ID'!B:B;'SD ID'!A:A);"")))`,
    supply_ordered: `=MAP(${item};LAMBDA(x;IF(x="";"";IFERROR(SUMIF('Снабжение'!A:A;x;'Снабжение'!C:C);""))))`,
    tied_capital: `=MAP(${item};${destination};${quantity};LAMBDA(x;d;q;IF(x="";"";IF(d="";IFNA(XLOOKUP(x;'ГТ Себестоимость'!B:B;'ГТ Себестоимость'!C:C);0)*q;0))))`,
    salesdrive_stock: `=ARRAYFORMULA(IF(${item}="";"";IFNA(XLOOKUP(REGEXREPLACE(${item};"Л$";"");REGEXREPLACE('Остатки'!E:E;"Л$";"");'Остатки'!F:F);"")))`,
    abc: `=ARRAYFORMULA(IF(${item}="";"";IF(${destination}<>"";"";IFNA(XLOOKUP(${item};ABC!A:A;ABC!B:B);"F"))))`,
    abc_components: `=ARRAYFORMULA(IF(${uuid}="";"";LET(parent;IFNA(XLOOKUP(${uuid};'${SPEC_RELATIONS_SHEET_NAME}'!A:A;'${SPEC_RELATIONS_SHEET_NAME}'!B:B);"");IF(parent="";"";IFNA(XLOOKUP(parent;${uuid};${abc});"")))))`,
    abc_components_2: `=ARRAYFORMULA(IF(${uuid}="";"";LET(parent;IFNA(XLOOKUP(${uuid};'${SPEC_RELATIONS_SHEET_NAME}'!A:A;'${SPEC_RELATIONS_SHEET_NAME}'!B:B);"");IF(parent="";"";IFNA(XLOOKUP(parent;${uuid};${abc1});"")))))`,
    priority: `=BYROW(HSTACK(${item};${note};${abc};${abc1};${abc2});LAMBDA(row;IF(INDEX(row;1;1)="";"";IF(OR(IFERROR(SEARCH("ждет клиент";INDEX(row;1;2));0);IFERROR(SEARCH("срочно";INDEX(row;1;2));0));1;MIN(ARRAYFORMULA(IFNA(XLOOKUP(CHOOSECOLS(row;3;4;5);'Служебная'!$O$2:$O$11;'Служебная'!$P$2:$P$11);999)))))))`,
    tooling: `=MAP(${item};LAMBDA(x;IF(x="";"";IFERROR(TEXTJOIN(", ";TRUE;FILTER('Диспетчер оснастки'!A:A;'Диспетчер оснастки'!B:B=x));""))))`,
    cost: `=ARRAYFORMULA(IF(${item}="";"";IF(${destination}="";IFNA(XLOOKUP(${item};'ГТ Себестоимость'!B:B;'ГТ Себестоимость'!C:C);"");"")))`,
    need_status: `=MAP(${uuid};${item};${needFlag};${status};${mode};LAMBDA(u;x;n;s;m;IF(x="";"";IF(n<>"Да";"Нет";IF(OR(s="Производство";s="Склад производства");"В производстве";IF(OR(s="ОТК";s="Приемка";s="Готово");"Готово";IF(m<>"В производство";"Покупной";IF(COUNTIF('${SPEC_RELATIONS_SHEET_NAME}'!B:B;u)=0;"Комплектующие не найдены";IF(IFERROR(SUM(XLOOKUP(FILTER('${SPEC_RELATIONS_SHEET_NAME}'!A:A;'${SPEC_RELATIONS_SHEET_NAME}'!B:B=u);${uuid};${deficit}));0)>0;"Нет";"Да")))))))))`
  });
}

/**
 * Удаляет построчные копии вычисляемых формул и устанавливает один якорь
 * на каждый computed field. Ручные/data/stage поля не изменяются.
 */
function installBaseFormulaModel_(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  setDispatcherRuntimeSpreadsheet_(ss);
  clearTableArchitectureCache_(ss);
  const architecture = getBaseArchitecture_(ss);
  const formulas = buildBaseComputedFormulas_(ss);
  const baseSheet = architecture.baseSheet;
  const maxRows = baseSheet.getMaxRows();

  for (const key of architecture.computedKeys) {
    const formula = formulas[key];
    if (!formula) throw new Error(`Для computed field «${key}» не определена каноническая формула.`);
    const col = baseCol_(ss, key);
    if (maxRows >= 2) baseSheet.getRange(2, col, maxRows - 1, 1).clearContent();
    baseSheet.getRange(2, col).setFormula(formula);
  }
  SpreadsheetApp.flush();
  return validateTableArchitecture_();
}
