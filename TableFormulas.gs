/**
 * Канонические формулы «База».
 * Формулы генерируются по field_key и устанавливаются только в строку 2.
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
  const c = key => baseColumnLetter_(ss, key);
  // Используется внутри LET(...;last;...): ограничивает тяжёлые массивные операции
  // последней реально занятой строкой, не фиксируя физический размер «База».
  const bounded = key => `${c(key)}2:INDEX(${c(key)}:${c(key)};last)`;

  const launchDate = r('date');
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
  const parentUuid = r('parent_uuid');
  const rootUuid = r('root_uuid');
  const parentName = r('parent_name');

  const stageRanges = STAGE_FIELD_KEYS.map(key => r(key));

  return Object.freeze({
    drawing_link: `=ARRAYFORMULA(IF(${item}="";"";LET(link;IFNA(XLOOKUP(${item};'Каталог ссылок'!A:A;'Каталог ссылок'!B:B);"§MISS§");IF(link="§MISS§";"Добавить изделие в каталог ссылок";IF(link="";"Добавить ссылку в каталог";link)))))`,
    order_link: `=ARRAYFORMULA(IF(${item}="";"";LET(link;IFNA(XLOOKUP(${item};'Каталог ссылок'!A:A;'Каталог ссылок'!C:C);"§MISS§");IF(link="§MISS§";"Добавить изделие в каталог ссылок";IF(link="";"Добавить ссылку в каталог";link)))))`,
    minimum_batch: `=ARRAYFORMULA(IF(${item}="";"";IFNA(XLOOKUP(${item};'Каталог ссылок'!A:A;'Каталог ссылок'!D:D);"")))`,
    stock: `=ARRAYFORMULA(IF(${item}="";"";IFNA(XLOOKUP(${item};'Остатки'!A:A;'Остатки'!B:B);0)))`,

    // Variant B допускает несколько строк одного артикула у разных родителей/root.
    // Один и тот же складской остаток нельзя вычитать из каждой строки независимо.
    // Сортируем применения по артикулу и исходной строке, считаем накопительную
    // потребность одним SCAN, а затем выделяем каждой строке только прирост общего
    // дефицита. Поэтому сумма deficit по артикулу = MAX(SUM(quantity) - stock; 0).
    // Тяжёлые операции ограничены последним UUID, а не миллионом строк листа.
    deficit: `=IFNA(LET(last;MAX(FILTER(ARRAYFORMULA(ROW(${uuid}));${uuid}<>""));items;${bounded('item_name')};qty;${bounded('quantity')};st;${bounded('stock')};rows;ARRAYFORMULA(ROW(items));data;FILTER(HSTACK(items;qty;st;rows);items<>"");s;SORT(data;1;TRUE;4;TRUE);si;CHOOSECOLS(s;1);sq;CHOOSECOLS(s;2);sst;CHOOSECOLS(s;3);sr;CHOOSECOLS(s;4);cum;SCAN(0;sq;LAMBDA(a;v;a+v));base;ARRAYFORMULA(XLOOKUP(si;si;ARRAYFORMULA(cum-sq)));gc;ARRAYFORMULA(cum-base);pgc;ARRAYFORMULA(gc-sq);alloc;ARRAYFORMULA(IF(gc>sst;gc-sst;0)-IF(pgc>sst;pgc-sst;0));back;SORT(HSTACK(sr;alloc);1;TRUE);ARRAYFORMULA(IF(items="";"";IFNA(XLOOKUP(rows;CHOOSECOLS(back;1);CHOOSECOLS(back;2));""))));"")`,

    need_flag: `=ARRAYFORMULA(IF(${item}="";"";IF(${deficit}>0;"Да";"Нет")))`,
    status: `=BYROW(HSTACK(${item};${stageRanges.join(';')});LAMBDA(row;IF(INDEX(row;1;1)="";"";IFNA(INDEX('Настройки'!A$2:A$14;MATCH(FALSE;CHOOSECOLS(row;2;3;4;5;6;7;8;9;10;11;12;13;14);0));"Готово"))))`,
    production_mode: `=ARRAYFORMULA(IF(${item}="";"";IF(IFNA(XLOOKUP(${item};'Цена материалов'!A:A;'Цена материалов'!A:A);"")="";"В производство";"В снабжение")))`,
    supplier: `=ARRAYFORMULA(IF(${item}="";"";IF(${mode}<>"В снабжение";"";IFNA(XLOOKUP(${item};'Цена материалов'!A:A;'Цена материалов'!D:D);""))))`,
    supplier_request_date: `=MAP(${item};${launchDate};${mode};LAMBDA(item;launch;mode;IF(item="";"";IF(mode<>"В снабжение";"";LET(d;MAXIFS('Импорт снабжения'!C$2:C$5000;'Импорт снабжения'!B$2:B$5000;item;'Импорт снабжения'!A$2:A$5000;">="&INT(launch);'Импорт снабжения'!E$2:E$5000;TRUE);IF(d=0;"";d))))))`,
    salesdrive_id: `=ARRAYFORMULA(IF(${item}="";"";IFNA(XLOOKUP(${item};'SD ID'!B:B;'SD ID'!A:A);"")))`,

    // Вместо SUMIF по всему «Снабжение» для каждой строки сначала один раз
    // агрегируем снабжение по изделию, затем выполняем один векторный XLOOKUP.
    supply_ordered: `=LET(agg;QUERY('Снабжение'!A2:C;"select A,sum(C) where A is not null group by A label sum(C) ''";0);ARRAYFORMULA(IF(${item}="";"";IFNA(XLOOKUP(${item};CHOOSECOLS(agg;1);CHOOSECOLS(agg;2));0))))`,

    tied_capital: `=MAP(${item};${parentUuid};${quantity};LAMBDA(x;p;q;IF(x="";"";IF(p="";IFNA(XLOOKUP(x;'ГТ Себестоимость'!B:B;'ГТ Себестоимость'!C:C);0)*q;0))))`,
    salesdrive_stock: `=ARRAYFORMULA(IF(${item}="";"";IFNA(XLOOKUP(REGEXREPLACE(${item};"Л$";"");REGEXREPLACE('Остатки'!E:E;"Л$";"");'Остатки'!F:F);"")))`,
    abc: `=ARRAYFORMULA(IF(${item}="";"";IF(${parentUuid}<>"";"";IFNA(XLOOKUP(${item};ABC!A:A;ABC!B:B);"F"))))`,
    abc_components: `=ARRAYFORMULA(IF(${uuid}="";"";IF(${parentUuid}="";"";IFNA(XLOOKUP(${parentUuid};${uuid};${abc});""))))`,
    abc_components_2: `=ARRAYFORMULA(IF(${uuid}="";"";IF(${parentUuid}="";"";IFNA(XLOOKUP(${parentUuid};${uuid};${abc1});""))))`,
    priority: `=BYROW(HSTACK(${item};${note};${abc};${abc1};${abc2});LAMBDA(row;IF(INDEX(row;1;1)="";"";IF(OR(IFERROR(SEARCH("ждет клиент";INDEX(row;1;2));0);IFERROR(SEARCH("срочно";INDEX(row;1;2));0));1;MIN(ARRAYFORMULA(IFNA(XLOOKUP(CHOOSECOLS(row;3;4;5);'Служебная'!$O$2:$O$11;'Служебная'!$P$2:$P$11);999)))))))`,
    tooling: `=MAP(${item};LAMBDA(x;IF(x="";"";IFERROR(TEXTJOIN(", ";TRUE;FILTER('Диспетчер оснастки'!A:A;'Диспетчер оснастки'!B:B=x));""))))`,
    cost: `=ARRAYFORMULA(IF(${item}="";"";IF(${parentUuid}="";IFNA(XLOOKUP(${item};'ГТ Себестоимость'!B:B;'ГТ Себестоимость'!C:C);"");"")))`,
    parent_name: `=ARRAYFORMULA(IF(${uuid}="";"";IF(${parentUuid}="";"";IFNA(XLOOKUP(${parentUuid};${uuid};${item});IFERROR(REGEXEXTRACT(${destination};"^(.+?)\\s*=\\s*");"")))))`,
    root_name: `=ARRAYFORMULA(IF(${uuid}="";"";IFNA(XLOOKUP(${rootUuid};${uuid};${item});IF(${rootUuid}=${parentUuid};${parentName};""))))`,

    // Variant B позволяет агрегировать дочерний дефицит один раз по parent_uuid.
    // Раньше каждая строка отдельно делала COUNTIF+FILTER+XLOOKUP по всей «База».
    need_status: `=LET(agg;QUERY(HSTACK(${parentUuid};${deficit});"select Col1,count(Col1),sum(Col2) where Col1 is not null group by Col1 label count(Col1) '', sum(Col2) ''";0);childCount;IFNA(XLOOKUP(${uuid};CHOOSECOLS(agg;1);CHOOSECOLS(agg;2));0);childDeficit;IFNA(XLOOKUP(${uuid};CHOOSECOLS(agg;1);CHOOSECOLS(agg;3));0);ARRAYFORMULA(IF(${item}="";"";IF(${needFlag}<>"Да";"Нет";IF(((${status}="Производство")+(${status}="Склад производства"))>0;"В производстве";IF(((${status}="ОТК")+(${status}="Приемка")+(${status}="Готово"))>0;"Готово";IF(${mode}<>"В производство";"Покупной";IF(childCount=0;"Комплектующие не найдены";IF(childDeficit>0;"Нет";"Да")))))))))`
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
  return validateTableArchitecture_();
}
