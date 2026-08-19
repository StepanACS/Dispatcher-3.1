/**
 * Объединяет подходящие дубли в «База» по семантическим полям.
 * Физические строки не удаляются: дубли очищаются, чтобы не сдвигать spill-якоря.
 */
function checkAndMergeDuplicates(baseSheet, columnCount) {
  if (!baseSheet) return;
  const ss = baseSheet.getParent();
  const startRow = 2;
  const lastRow = baseLastDataRow_(baseSheet);
  if (lastRow < startRow) return;

  const keys = ['item_name', 'destination', 'quantity', 'stage_picker', 'stage_production'];
  const columns = keys.map(key => ({ key, col: baseCol_(ss, key) }));
  const minCol = Math.min(...columns.map(x => x.col));
  const maxCol = Math.max(...columns.map(x => x.col));
  const data = baseSheet.getRange(startRow, minCol, lastRow - startRow + 1, maxCol - minCol + 1).getValues();
  const idx = key => columns.find(x => x.key === key).col - minCol;

  const uniqueItems = new Map();
  const rowsToClear = [];

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const itemName = String(row[idx('item_name')] || '').trim();
    if (!itemName) continue;

    const productionMarker = row[idx('stage_production')];
    const pickerMarker = row[idx('stage_picker')];
    const destination = String(row[idx('destination')] || '').trim();
    if ((productionMarker !== false && productionMarker !== null && productionMarker !== '') ||
        pickerMarker !== false || !destination) continue;

    const quantityRaw = row[idx('quantity')];
    const quantity = typeof quantityRaw === 'number' ? quantityRaw : parseFloat(quantityRaw) || 0;
    const sheetRow = i + startRow;

    if (!uniqueItems.has(itemName)) {
      uniqueItems.set(itemName, { sheetRow, destination, quantity, changed: false });
      continue;
    }

    const target = uniqueItems.get(itemName);
    if (destination && !target.destination.split('\n').includes(destination)) {
      target.destination = target.destination ? `${target.destination}\n${destination}` : destination;
      target.changed = true;
    }
    if (quantity !== 0) {
      target.quantity += quantity;
      target.changed = true;
    }
    rowsToClear.push(sheetRow);
  }

  for (const entry of uniqueItems.values()) {
    if (!entry.changed) continue;
    writeBaseFields_(baseSheet, entry.sheetRow, {
      destination: entry.destination,
      quantity: entry.quantity
    });
  }
  clearBaseRawRows_(baseSheet, rowsToClear);
}
