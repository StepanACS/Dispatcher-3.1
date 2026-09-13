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

function groupContiguousRows_(rowNumbers) {
  const rows = [...new Set((rowNumbers || []).map(Number).filter(row => row >= 2))]
    .sort((a, b) => a - b);
  if (!rows.length) return [];
  const groups = [];
  let start = rows[0];
  let end = rows[0];
  for (let i = 1; i < rows.length; i++) {
    if (rows[i] === end + 1) end = rows[i];
    else { groups.push({ start, end }); start = rows[i]; end = rows[i]; }
  }
  groups.push({ start, end });
  return groups;
}

/**
 * Очищает произвольный набор строк одним RangeList вместо N×M отдельных RPC.
 * Формульные столбцы намеренно исключены, поэтому spill-якоря не затрагиваются.
 */
function clearBaseRawRows_(baseSheet, rowNumbers) {
  if (!baseSheet || !rowNumbers || !rowNumbers.length) return;
  const ss = baseSheet.getParent();
  const architecture = getBaseArchitecture_(ss);
  const rawColumns = [];

  for (const [key, col] of architecture.keyToColumn) {
    if (!architecture.computedKeys.has(key)) rawColumns.push({ key, col });
  }
  const columnGroups = groupContiguousColumns_(rawColumns);
  const rowGroups = groupContiguousRows_(rowNumbers);
  if (!columnGroups.length || !rowGroups.length) return;

  const ranges = [];
  for (const rows of rowGroups) {
    for (const cols of columnGroups) {
      const left = columnNumberToLetter_(cols[0].col);
      const right = columnNumberToLetter_(cols[cols.length - 1].col);
      ranges.push(`${left}${rows.start}:${right}${rows.end}`);
    }
  }
  baseSheet.getRangeList(ranges).clearContent();
}
