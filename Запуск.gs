/**
 * Переносит одну строку (onEdit) либо все заполненные строки листа «Запуск» в «База».
 */
function handleLaunchEdit(ss, launchSheet, rowNumber) {
  if (!launchSheet) return;

  const baseSheet = ss.getSheetByName('База');
  if (!baseSheet) throw new Error('Лист «База» не найден.');

  const baseColumnCount = baseSheet.getLastColumn();
  const startRowForBase = baseSheet.getLastRow() + 1;
  const firstDataRow = 2;
  const lastColumn = Math.max(launchSheet.getLastColumn(), 1);

  let range;
  let sourceStartRow;

  if (rowNumber) {
    sourceStartRow = rowNumber;
    range = launchSheet.getRange(rowNumber, 1, 1, lastColumn);
  } else {
    const lastRow = launchSheet.getLastRow();
    if (lastRow < firstDataRow) {
      ss.toast('Нет новых изделий для добавления', 'Информация', 3);
      return;
    }
    sourceStartRow = firstDataRow;
    range = launchSheet.getRange(firstDataRow, 1, lastRow - firstDataRow + 1, lastColumn);
  }

  const sourceData = range.getValues();
  const outputRows = [];
  const rowsToDelete = [];

  sourceData.forEach((sourceRow, index) => {
    const itemName = String(sourceRow[0] || '').trim();
    const quantity = String(sourceRow[2] ?? '').trim();
    if (!itemName || !quantity) return;

    const futureBaseRow = startRowForBase + outputRows.length;
    outputRows.push(buildBaseRow(sourceRow, futureBaseRow, baseColumnCount));
    rowsToDelete.push(sourceStartRow + index);
  });

  if (outputRows.length === 0) {
    ss.toast('Нет новых изделий для добавления', 'Информация', 3);
    return;
  }

  writeRowsToSheetForLaunch(baseSheet, startRowForBase, outputRows, baseColumnCount);

  const checkboxRange = baseSheet.getRange(
    startRowForBase,
    COL_DESIGN,
    outputRows.length,
    COL_READY - COL_DESIGN + 1
  );
  checkboxRange.setDataValidation(
    SpreadsheetApp.newDataValidation().requireCheckbox().setAllowInvalid(false).build()
  );

  // deleteRows() физически уменьшает число строк листа. Чтобы лист «Запуск» не
  // сжимался после каждого переноса, возвращаем ровно столько пустых строк,
  // сколько было удалено. Это сохраняет текущий размер листа без исторического
  // механизма расширения до 50+ строк.
  deleteRowsInBlocks_(launchSheet, rowsToDelete);
  restoreLaunchInputRows_(launchSheet, rowsToDelete.length);

  ss.toast('Изделие(я) добавлено(ы) в базу', 'Информация', 5);
}

/** Формирует одну строку «Базы» из строки листа «Запуск». */
function buildBaseRow(rowData, futureIndex, totalColumnCount) {
  const newRow = buildBaseRowTemplate(futureIndex, totalColumnCount);

  newRow[COL_DATE - 1] = new Date();
  newRow[COL_ITEMNAME - 1] = capitalizeFirstLetter(rowData[0]);
  newRow[COL_DESTINATION - 1] = capitalizeFirstLetter(rowData[1]);
  newRow[COL_NOTE - 1] = String(rowData[3] || '').trim();
  newRow[COL_QUANTITY - 1] = rowData[2];

  return newRow;
}

/** Записывает подготовленные строки одним setValues(). */
function writeRowsToSheetForLaunch(sheet, startRow, rows, columnCount) {
  if (!rows || rows.length === 0) return;

  const maxRows = sheet.getMaxRows();
  const neededEndRow = startRow + rows.length - 1;
  if (neededEndRow > maxRows) {
    sheet.insertRowsAfter(maxRows, neededEndRow - maxRows);
  }

  sheet.getRange(startRow, 1, rows.length, columnCount).setValues(rows);
}

/**
 * Возвращает в конец «Запуск» удалённое количество строк и восстанавливает
 * в столбце E checkbox-validation. Данные A:D остаются пустыми.
 */
function restoreLaunchInputRows_(sheet, count) {
  if (!count || count < 1) return;

  const previousMaxRows = sheet.getMaxRows();
  sheet.insertRowsAfter(previousMaxRows, count);

  const firstNewRow = previousMaxRows + 1;
  const rule = SpreadsheetApp.newDataValidation()
    .requireCheckbox()
    .setAllowInvalid(false)
    .build();

  sheet
    .getRange(firstNewRow, COL_LAUNCH_TRIGGER, count, 1)
    .setDataValidation(rule)
    .setValue(false);
}

/**
 * Удаляет набор строк снизу вверх непрерывными блоками, сокращая количество
 * структурных операций Sheets по сравнению с deleteRow() для каждой строки.
 */
function deleteRowsInBlocks_(sheet, rowNumbers) {
  if (!rowNumbers || rowNumbers.length === 0) return;

  const sorted = [...new Set(rowNumbers)].sort((a, b) => b - a);
  let blockHigh = sorted[0];
  let blockLow = sorted[0];

  const flushBlock = () => {
    sheet.deleteRows(blockLow, blockHigh - blockLow + 1);
  };

  for (let i = 1; i < sorted.length; i++) {
    if (sorted[i] === blockLow - 1) {
      blockLow = sorted[i];
    } else {
      flushBlock();
      blockHigh = sorted[i];
      blockLow = sorted[i];
    }
  }
  flushBlock();
}
