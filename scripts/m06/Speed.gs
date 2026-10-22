/**
 * Модуль 6, урок 1. Диапазон как массив.
 *
 * Две функции делают одно и то же: считают заказы дольше 8 часов.
 * Запустите обе и сравните время в журнале выполнения.
 */

/** Медленно: по одному обращению к таблице на каждую строку. */
function countLongOrdersSlow() {
  const start = Date.now();
  const sheet = SpreadsheetApp.getActive().getSheetByName('Orders');
  let count = 0;
  for (let row = 2; row <= sheet.getLastRow(); row++) {
    if (sheet.getRange(row, 9).getValue() > 8) count++;
  }
  console.log('Slow: ' + count + ' orders, ' + (Date.now() - start) + ' ms');
}

/** Быстро: одно обращение к таблице, дальше работа с массивом в памяти. */
function countLongOrdersFast() {
  const start = Date.now();
  const values = SpreadsheetApp.getActive().getSheetByName('Orders').getDataRange().getValues();
  const count = values.slice(1).filter(function (row) { return row[8] > 8; }).length;
  console.log('Fast: ' + count + ' orders, ' + (Date.now() - start) + ' ms');
}

/**
 * Шаблон любой обработки данных: прочитать → обработать → записать.
 * Здесь: в колонку P «Hours as Text» пишем «2 h 15 min» для каждого заказа.
 */
function writeHoursAsText() {
  const sheet = SpreadsheetApp.getActive().getSheetByName('Orders');
  const n = sheet.getLastRow() - 1;

  // 1. Прочитать — один вызов
  const hours = sheet.getRange(2, 9, n, 1).getValues();

  // 2. Обработать — в памяти, без обращений к таблице
  const texts = hours.map(function (row) {
    const h = Math.floor(row[0]);
    const m = Math.round((row[0] - h) * 60);
    return [h + ' h ' + m + ' min'];
  });

  // 3. Записать — один вызов
  sheet.getRange(1, 16).setValue('Hours as Text');
  sheet.getRange(2, 16, n, 1).setValues(texts);
}
