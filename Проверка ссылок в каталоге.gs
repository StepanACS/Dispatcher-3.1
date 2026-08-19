/**
 * Проверяет, что ссылка на папку (столбец B) ведёт в папку с именем,
 * совпадающим с кодом изделия из столбца A (с обрезкой после точки).
 *
 * Таблица: 1XffWoUclgQyFcT4DOIRoenQ9jAbBxXam6YxyY8xgv_4
 * Лист: "Каталог ссылок"
 * A: код изделия ("Р01" или "Р01.001")
 * B: ссылка на папку в Google Drive
 */
function checkCatalogFolderLinks() {
  const SPREADSHEET_ID = "1XffWoUclgQyFcT4DOIRoenQ9jAbBxXam6YxyY8xgv_4";
  const SHEET_NAME = "Каталог ссылок";
  const COL_PRODUCT = 1; // A
  const COL_LINK = 2;    // B
  const START_ROW = 2;   // предполагаем заголовок в 1-й строке

  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sh = ss.getSheetByName(SHEET_NAME);
    if (!sh) throw new Error(`Лист не найден: ${SHEET_NAME}`);

    const lastRow = sh.getLastRow();
    if (lastRow < START_ROW) {
      Logger.log(`[ok] Нет данных для проверки (lastRow=${lastRow}).`);
      return;
    }

    const range = sh.getRange(START_ROW, 1, lastRow - START_ROW + 1, 2);
    const values = range.getValues();

    let checked = 0;
    let mismatches = 0;
    let skippedEmpty = 0;
    let skippedBadLink = 0;

    values.forEach((row, idx) => {
      const rowNumber = START_ROW + idx;
      const productRaw = String(row[COL_PRODUCT - 1] ?? "").trim();
      const linkRaw = String(row[COL_LINK - 1] ?? "").trim();

      // Если нет кода изделия — пропускаем (не оговорено, но безопасно)
      if (!productRaw) return;

      // Пустая ссылка — пропускаем по ТЗ
      if (!linkRaw) {
        skippedEmpty++;
        return;
      }

      const expected = normalizeProductToRoot_(productRaw); // "Р01.001" -> "Р01"
      const folderId = extractDriveFolderId_(linkRaw);

      if (!folderId) {
        skippedBadLink++;
        Logger.log(`[warn] Не удалось извлечь ID папки. row=${rowNumber}, изделие="${productRaw}", ссылка="${linkRaw}"`);
        return;
      }

      let folderName = "";
      try {
        folderName = DriveApp.getFolderById(folderId).getName().trim();
      } catch (e) {
        skippedBadLink++;
        Logger.log(`[warn] Ошибка доступа к папке. row=${rowNumber}, изделие="${productRaw}", folderId="${folderId}", err="${e.message}"`);
        return;
      }

      checked++;
      if (folderName !== expected) {
        mismatches++;
        Logger.log(`[mismatch] изделие="${productRaw}" (ожидалось "${expected}") != папка="${folderName}" | row=${rowNumber} | folderId=${folderId}`);
      }
    });

    Logger.log(`[done] checked=${checked}, mismatches=${mismatches}, skippedEmpty=${skippedEmpty}, skippedBadLink=${skippedBadLink}`);
  } catch (err) {
    Logger.log(`[ERROR] checkCatalogFolderLinks: ${err && err.message ? err.message : err}`);
    throw err;
  }
}

/**
 * "Р01" -> "Р01"
 * "Р01.001" -> "Р01"
 * "Р01.001.02" -> "Р01" (на всякий случай)
 */
function normalizeProductToRoot_(productCode) {
  const s = String(productCode ?? "").trim();
  const dotPos = s.indexOf(".");
  return dotPos === -1 ? s : s.slice(0, dotPos).trim();
}

/**
 * Извлекает ID папки из разных форматов ссылок:
 * - https://drive.google.com/drive/folders/<ID>
 * - https://drive.google.com/open?id=<ID>
 * - ...?usp=sharing и т.п.
 * Также поддерживает ситуацию, когда в ячейке уже просто ID.
 */
function extractDriveFolderId_(input) {
  const s = String(input ?? "").trim();
  if (!s) return "";

  // Если это уже похоже на ID (без пробелов, без слэшей, достаточно длинное)
  if (!s.includes("/") && !s.includes("http") && s.length >= 20) return s;

  // /folders/<id>
  let m = s.match(/\/folders\/([a-zA-Z0-9_-]+)/);
  if (m && m[1]) return m[1];

  // id=<id>
  m = s.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  if (m && m[1]) return m[1];

  // Иногда встречается .../u/0/folders/<id> (покрывается первым), либо другие варианты:
  // Пробуем вытащить самый длинный токен похожий на ID
  m = s.match(/([a-zA-Z0-9_-]{20,})/);
  if (m && m[1]) return m[1];

  return "";
}