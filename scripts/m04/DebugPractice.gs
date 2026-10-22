/**
 * Модуль 4, урок 3. Тренажёр отладки.
 *
 * В каждой функции ниже — одна типичная ошибка. Запустите функцию,
 * прочитайте текст ошибки в журнале выполнения, найдите и исправьте причину.
 * Правильные версии — в файле DebugSolutions.gs. Не подглядывайте сразу!
 */

/** Задача 1. Посчитать заказы со статусом «Problem». */
function bug1_countProblems() {
  const sheet = SpreadsheetApp.getActive().getSheetByName('Orders ');
  const values = sheet.getDataRange().getValues();
  const problems = values.filter(function (row) { return row[4] === 'Problem'; });
  console.log('Problem orders: ' + problems.length);
}

/** Задача 2. Записать на лист «Summary» таблицу «Тип — число заказов». */
function bug2_writeSummary() {
  const rows = [
    ['Type', 'Orders'],
    ['FP', 95],
    ['ESX', 82],
    ['FF', 23],
  ];
  const sheet = SpreadsheetApp.getActive().getSheetByName('Summary');
  sheet.getRange(1, 1, 3, 2).setValues(rows);
}

/**
 * Задача 3. Покрасить в жёлтый часы больше 8.
 * Ошибки нет, но на большом листе функция работает очень долго. Почему?
 */
function bug3_slowHighlight() {
  const sheet = SpreadsheetApp.getActive().getSheetByName('Orders');
  for (let row = 2; row <= sheet.getLastRow(); row++) {
    const hours = sheet.getRange(row, 9).getValue();
    if (hours > 8) {
      sheet.getRange(row, 9).setBackground('#fff59d');
    }
  }
}

/** Задача 4. Вывести исполнителя первого заказа. */
function bug4_firstWorker() {
  const values = SpreadsheetApp.getActive().getSheetByName('Orders').getDataRange().getValues();
  // Хотим колонку G «Assignee» первой строки данных
  console.log('First assignee: ' + values[1][7]);
}
