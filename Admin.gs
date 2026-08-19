function testEnsureFreeRows(){
  const ss = SpreadsheetApp.openById("1XffWoUclgQyFcT4DOIRoenQ9jAbBxXam6YxyY8xgv_4");
  const sheet = ss.getSheetByName("Запуск");
  const templateRow = 18;
  ensureFreeRows(sheet, templateRow, 30)
}