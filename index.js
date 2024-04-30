function uniqueTranslate() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet();
    var sourceSheet = sheet.getSheetByName('Лист1'); // Замените на имя вашего листа с данными
    var targetSheet = sheet.getSheetByName('Перевод') || sheet.insertSheet('Перевод');
    var range = sourceSheet.getDataRange();
    var values = range.getValues();
  
    // Собираем уникальные строки для перевода
    var uniqueStrings = new Set();
    values.forEach(function(row) {
      row.forEach(function(cell) {
        if (typeof cell === 'string' && cell.trim() !== '') {
          uniqueStrings.add(cell.trim());
        }
      });
    });
  
    // Переводим уникальные строки
    var translations = {};
    uniqueStrings.forEach(function(text) {
      var translatedText = LanguageApp.translate(text, 'ka', 'ru');
      translations[text] = translatedText;
    });
  
    // Заменяем исходные строки на переведенные во всем массиве данных
    var translatedValues = values.map(function(row) {
      return row.map(function(cell) {
        return translations[cell] || cell; // Возвращаем перевод, если он есть, иначе оригинальное значение
      });
    });
  
    // Записываем переведенные данные в новый лист
    targetSheet.clear(); // Очищаем лист перед записью новых данных
    targetSheet.getRange(1, 1, translatedValues.length, translatedValues[0].length).setValues(translatedValues);
  }
  