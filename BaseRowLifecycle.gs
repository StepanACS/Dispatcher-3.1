/**
 * Операции жизненного цикла строк «База».
 *
 * После перехода на spill-формулы физические deleteRow/deleteRows для «База»
 * запрещены: они могут сдвинуть или удалить формулы-якоря строки 2. Вместо этого
 * очищаются только фактические (невычисляемые) поля записи.
 */
function clearBaseRawRow_(baseSheet, row) {
  clearBaseRawRows_(baseSheet, [row]);
}

function clearBaseRawRows_(baseSheet, rowNumbers) {
  if (!baseSheet || !rowNumbers || !rowNumbers.length) return;
  const ss = baseSheet.getParent();
  const architecture = getBaseArchitecture_(ss);
  const rawColumns = [];

  for (const [key, col] of architecture.keyToColumn) {
    if (!architecture.computedKeys.has(key)) rawColumns.push({ key, col });
  }
  const groups = groupContiguousColumns_(rawColumns);
  const rows = [...new Set(rowNumbers.map(Number).filter(row => row >= 2))];

  for (const row of rows) {
    for (const group of groups) {
      baseSheet.getRange(row, group[0].col, 1, group.length).clearContent();
    }
  }
}

/** Находит первую свободную строку по UUID, не затрагивая строку 2 с якорями. */
function findReusableBaseRow_(baseSheet, fromRow = 3) {
  const ss = baseSheet.getParent();
  const uuidCol = baseCol_(ss, 'uuid');
  const lastRow = Math.max(baseLastDataRow_(baseSheet), fromRow);
  if (lastRow < fromRow) return fromRow;

  const values = baseSheet.getRange(fromRow, uuidCol, lastRow - fromRow + 1, 1).getValues();
  for (let i = 0; i < values.length; i++) {
    if (!String(values[i][0] ?? '').trim()) return fromRow + i;
  }
  return lastRow + 1;
}
