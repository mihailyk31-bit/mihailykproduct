function setupSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  sheet.setName('Записи');
  sheet.clearContents();

  var headers = ['№','Дата записи','Имя','Телефон','Врач','Услуга','Цена','Дата визита','Время','Промокод','Скидка','Аллергия','Страх боли','Первый визит','Комментарий','Оценка NPS'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#1a237e');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');
  headerRange.setFontSize(11);
  headerRange.setHorizontalAlignment('center');
  headerRange.setVerticalAlignment('middle');
  sheet.setRowHeight(1, 40);
  sheet.setFrozenRows(1);
  headerRange.setBorder(true, true, true, true, true, true, '#3949ab', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  var widths = [40,150,160,150,180,180,120,120,80,100,80,150,100,100,200,80];
  widths.forEach(function(w, i) { sheet.setColumnWidth(i + 1, w); });

  SpreadsheetApp.flush();
}

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Записи') || ss.getActiveSheet();

    if (sheet.getLastRow() === 0) setupSheet();

    var rowNum = sheet.getLastRow();
    var bgColor = rowNum % 2 === 0 ? '#e8eaf6' : '#ffffff';

    var newRow = [
      rowNum,
      new Date().toLocaleString('ru-RU'),
      data.name || '',
      data.phone || '',
      data.doctor || '',
      data.service || '',
      data.price || '',
      data.date || '',
      data.time || '',
      data.promo || '',
      data.discount > 0 ? '-' + data.discount + '%' : '',
      data.allergy || '',
      data.fear === 3 ? 'Сильный' : data.fear === 2 ? 'Средний' : '',
      data.firstVisit ? 'Да' : '',
      data.comment || '',
      data.npsScore ? data.npsScore + '/10' : ''
    ];

    var range = sheet.getRange(rowNum + 1, 1, 1, newRow.length);
    range.setValues([newRow]);
    range.setBackground(bgColor);
    range.setFontSize(10);
    range.setVerticalAlignment('middle');
    sheet.setRowHeight(rowNum + 1, 32);
    range.setBorder(true, true, true, true, true, true, '#c5cae9', SpreadsheetApp.BorderStyle.SOLID);

    if (data.fear === 3) sheet.getRange(rowNum+1,13).setBackground('#ffcdd2').setFontColor('#c62828').setFontWeight('bold');
    if (data.allergy) sheet.getRange(rowNum+1,12).setBackground('#fff9c4').setFontColor('#f57f17');
    if (data.firstVisit) sheet.getRange(rowNum+1,14).setBackground('#c8e6c9').setFontColor('#2e7d32').setFontWeight('bold');
    if (data.npsScore >= 9) sheet.getRange(rowNum+1,16).setBackground('#c8e6c9').setFontColor('#2e7d32').setFontWeight('bold');
    if (data.npsScore && data.npsScore <= 6) sheet.getRange(rowNum+1,16).setBackground('#ffcdd2').setFontColor('#c62828').setFontWeight('bold');

    SpreadsheetApp.flush();
    return ContentService.createTextOutput('ok');
  } catch(err) {
    return ContentService.createTextOutput('error: ' + err.message);
  }
}

function doGet(e) {
  setupSheet();
  return ContentService.createTextOutput('done');
}
