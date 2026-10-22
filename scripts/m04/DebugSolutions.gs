/**
 * Модуль 4, урок 3. Ответы к тренажёру отладки.
 */

/**
 * Задача 1. Ошибка: TypeError: Cannot read properties of null (reading 'getDataRange').
 * Причина: в имени листа лишний пробел — 'Заказы ' вместо 'Заказы'.
 * getSheetByName не нашёл лист и вернул null.
 */
function fix1_countProblems() {
  const sheet = SpreadsheetApp.getActive().getSheetByName('Заказы');
  const values = sheet.getDataRange().getValues();
  const problems = values.filter(function (row) { return row[4] === 'Проблема'; });
  console.log('Проблемных заказов: ' + problems.length);
}

/**
 * Задача 2. Ошибка: The number of rows in the data does not match
 * the number of rows in the range. The data has 4 but the range has 3.
 * Причина: в массиве 4 строки (вместе с шапкой), а диапазон на 3.
 * Лечение: брать размер из самого массива, а не писать числа руками.
 */
function fix2_writeSummary() {
  const rows = [
    ['Тип', 'Заказов'],
    ['FP', 95],
    ['ESX', 82],
    ['FF', 23],
  ];
  const sheet = SpreadsheetApp.getActive().getSheetByName('Сводка');
  sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
}

/**
 * Задача 3. Ошибки нет, но каждый getValue и setBackground — отдельный
 * запрос к таблице. 200 строк = до 400 запросов.
 * Лечение: прочитать колонку одним getValues, посчитать цвета в массиве
 * и записать их одним setBackgrounds.
 */
function fix3_fastHighlight() {
  const sheet = SpreadsheetApp.getActive().getSheetByName('Заказы');
  const range = sheet.getRange(2, 9, sheet.getLastRow() - 1, 1);
  const colors = range.getValues().map(function (row) {
    return [row[0] > 8 ? '#fff59d' : null]; // null — без заливки
  });
  range.setBackgrounds(colors);
}

/**
 * Задача 4. Ошибки нет, но выводится супервайзер, а не исполнитель.
 * Причина: в массиве колонки считаются с 0. G — седьмая буква, значит индекс 6.
 */
function fix4_firstWorker() {
  const values = SpreadsheetApp.getActive().getSheetByName('Заказы').getDataRange().getValues();
  console.log('Первый исполнитель: ' + values[1][6]);
}
