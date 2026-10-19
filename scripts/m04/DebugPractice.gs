/**
 * Модуль 4, урок 4.3. Тренажёр отладки.
 *
 * В каждой функции ниже — одна типичная ошибка. Запустите функцию,
 * прочитайте текст ошибки в журнале выполнения, найдите и исправьте причину.
 * Правильные версии — в файле DebugSolutions.gs. Не подглядывайте сразу!
 */

/** Задача 1. Посчитать заказы со статусом «Проблема». */
function bug1_countProblems() {
  const sheet = SpreadsheetApp.getActive().getSheetByName('Заказы ');
  const values = sheet.getDataRange().getValues();
  const problems = values.filter(function (row) { return row[4] === 'Проблема'; });
  console.log('Проблемных заказов: ' + problems.length);
}

/** Задача 2. Записать на лист «Сводка» таблицу «Тип — число заказов». */
function bug2_writeSummary() {
  const rows = [
    ['Тип', 'Заказов'],
    ['FP', 95],
    ['ESX', 82],
    ['FF', 23],
  ];
  const sheet = SpreadsheetApp.getActive().getSheetByName('Сводка');
  sheet.getRange(1, 1, 3, 2).setValues(rows);
}

/**
 * Задача 3. Покрасить в жёлтый часы больше 8.
 * Ошибки нет, но на большом листе функция работает очень долго. Почему?
 */
function bug3_slowHighlight() {
  const sheet = SpreadsheetApp.getActive().getSheetByName('Заказы');
  for (let row = 2; row <= sheet.getLastRow(); row++) {
    const hours = sheet.getRange(row, 9).getValue();
    if (hours > 8) {
      sheet.getRange(row, 9).setBackground('#fff59d');
    }
  }
}

/** Задача 4. Вывести исполнителя первого заказа. */
function bug4_firstWorker() {
  const values = SpreadsheetApp.getActive().getSheetByName('Заказы').getDataRange().getValues();
  // Хотим колонку G «Исполнитель» первой строки данных
  console.log('Первый исполнитель: ' + values[1][7]);
}
