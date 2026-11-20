function getDatabaseData() {
  var ss = SpreadsheetApp.openById('1ZiIFxFHYgc6qqnDw5np8tro6NyA0Ws9iS5iH9bQPcj8');
  var sheet = ss.getSheetByName('Sheet1'); // আপনার ট্যাবের সঠিক নাম দিন
  var data = sheet.getDataRange().getValues();
  Logger.log(data);
}
