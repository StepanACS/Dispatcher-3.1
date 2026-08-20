/** Переносит одну строку (onEdit) либо все заполненные строки «Запуск» в «База». */
function handleLaunchEdit(ss, launchSheet, rowNumber) {
  if (!launchSheet) return;
  const baseSheet = ss.getSheetByName(BASE_SHEET_NAME);
  if (!baseSheet) throw new Error('Лист «База» не найден.');

  const startRowForBase = baseLastDataRow_(baseSheet) + 1;
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
    outputRows.push(buildBaseRow(sourceRow));
    rowsToDelete.push(sourceStartRow + index);
  });

  if (!outputRows.length) {
    ss.toast('Нет новых изделий для добавления', 'Информация', 3);
    return;
  }

  writeBaseRows_(baseSheet, startRowForBase, outputRows);
  applyBaseStageValidation_(ss, baseSheet, startRowForBase, outputRows.length);

  deleteRowsInBlocks_(launchSheet, rowsToDelete);
  restoreLaunchInputRows_(launchSheet, rowsToDelete.length);
  ss.toast('Изделие(я) добавлено(ы) в базу', 'Информация', 5);
}

function buildBaseRow(rowData) {
  const newRow = buildBaseRowTemplate();
  newRow.date = new Date();
  newRow.item_name = capitalizeFirstLetter(rowData[0]);
  newRow.destination = capitalizeFirstLetter(rowData[1]);
  newRow.note = String(rowData[3] || '').trim();
  newRow.quantity = rowData[2];
  return newRow;
}

/** Ставит checkbox-validation на этапные столбцы независимо от их порядка. */
function applyBaseStageValidation_(ss, baseSheet, startRow, rowCount) {
  if (!rowCount) return;
  const rule = SpreadsheetApp.newDataValidation().requireCheckbox().setAllowInvalid(false).build();
  for (const fieldKey of STAGE_FIELD_KEYS) {
    baseSheet.getRange(startRow, baseCol_(ss, fieldKey), rowCount, 1).setDataValidation(rule);
  }
}

function restoreLaunchInputRows_(sheet, count) {
  if (!count || count < 1) return;
  const previousMaxRows = sheet.getMaxRows();
  sheet.insertRowsAfter(previousMaxRows, count);
  const firstNewRow = previousMaxRows + 1;
  const rule = SpreadsheetApp.newDataValidation().requireCheckbox().setAllowInvalid(false).build();
  sheet.getRange(firstNewRow, COL_LAUNCH_TRIGGER, count, 1).setDataValidation(rule).setValue(false);
}

function deleteRowsInBlocks_(sheet, rowNumbers) {
  if (!rowNumbers || !rowNumbers.length) return;
  const sorted = [...new Set(rowNumbers)].sort((a, b) => b - a);
  let blockHigh = sorted[0];
  let blockLow = sorted[0];
  const flushBlock = () => sheet.deleteRows(blockLow, blockHigh - blockLow + 1);

  for (let i = 1; i < sorted.length; i++) {
    if (sorted[i] === blockLow - 1) blockLow = sorted[i];
    else {
      flushBlock();
      blockHigh = sorted[i];
      blockLow = sorted[i];
    }
  }
  flushBlock();
}
