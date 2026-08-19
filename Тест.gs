function sheetTest() {
  var ss = SpreadsheetApp.openById("1JTr3AFm5Q1ycipTsrC5qY-SGKWY3v_w_2Jgfhq2xHgU");
  sheet = ss.getSheetByName("База"); 
  Logger.log ("Количество строк на листе: " + sheet.getMaxRows());
}

function testDeficit(){ 
  var ss = SpreadsheetApp.openById("1JTr3AFm5Q1ycipTsrC5qY-SGKWY3v_w_2Jgfhq2xHgU");
  Logger.log(calculateDeficit("НП6х28Ч", 10, ss));
}

function testUploadSpec(){
Logger.log(uploadSpecToFirestoreFromSheet("Б4004"));
}