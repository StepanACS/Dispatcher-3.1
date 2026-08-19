/**
 * Объединяет подходящие дубли в «База».
 *
 * Для определения дублей нужны только столбцы E:V. Раньше функция дважды
 * загружала все 40 столбцов (values + formulas), а затем перезаписывала целые
 * строки. Теперь изменяются только F (назначение) и H (количество).
 */
function checkAndMergeDuplicates(baseSheet, columnCount) {
  if (!baseSheet) return;

  const startRow = 2;
  const lastRow = baseSheet.getLastRow();
  if (lastRow < startRow) return;

  const firstCol = COL_ITEMNAME;
  const lastCol = COL_PRODUCTION;
  const data = baseSheet
    .getRange(startRow, firstCol, lastRow - startRow + 1, lastCol - firstCol + 1)
    .getValues();

  const offset = col => col - firstCol;
  const uniqueItems = new Map();
  const rowsToDelete = [];

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const productionMarker = row[offset(COL_PRODUCTION)];
    if (productionMarker !== false && productionMarker !== null && productionMarker !== '') {
      continue;
    }

    const itemName = row[offset(COL_ITEMNAME)];
    const destination = row[offset(COL_DESTINATION)];
    const quantity = row[offset(COL_QUANTITY)];
    const pickerMarker = row[offset(COL_PICKER)];

    if (!itemName || pickerMarker !== false || !destination) continue;

    if (!uniqueItems.has(itemName)) {
      uniqueItems.set(itemName, {
        rowIndex: i,
        destination: String(destination || ''),
        quantity: typeof quantity === 'number' ? quantity : parseFloat(quantity) || 0,
        changed: false
      });
      continue;
    }

    const target = uniqueItems.get(itemName);
    const currentDestination = String(destination || '');
    if (currentDestination && !target.destination.includes(currentDestination)) {
      target.destination = target.destination
        ? target.destination + '\n' + currentDestination
        : currentDestination;
      target.changed = true;
    }

    const currentQty = typeof quantity === 'number' ? quantity : parseFloat(quantity) || 0;
    if (currentQty !== 0) {
      target.quantity += currentQty;
      target.changed = true;
    }

    rowsToDelete.push(i + startRow);
  }

  // Записываем только реально изменившиеся ячейки, не затрагивая формулы остальных столбцов.
  for (const entry of uniqueItems.values()) {
    if (!entry.changed) continue;
    const sheetRow = entry.rowIndex + startRow;
    baseSheet.getRange(sheetRow, COL_DESTINATION).setValue(entry.destination);
    baseSheet.getRange(sheetRow, COL_QUANTITY).setValue(entry.quantity);
  }

  deleteRowsInBlocks_(baseSheet, rowsToDelete);
}
