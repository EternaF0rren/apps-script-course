/**
 * Модуль 8, урок 8.3. Тренды и выбросы.
 */

/**
 * Добавляет на «Сводку» две колонки к блоку по неделям:
 *   F — изменение числа заказов к прошлой неделе, в процентах;
 *   G — скользящее среднее заказов за 3 недели (сглаживает случайные скачки).
 * Запускайте после buildSummary.
 */
function addTrendColumns() {
  const sheet = SpreadsheetApp.getActive().getSheetByName('Сводка');
  const lastRow = sheet.getRange('A1:A').getValues().filter(function (row) { return row[0] !== ''; }).length;
  if (lastRow < 3) return; // нужно хотя бы две недели

  const counts = sheet.getRange(2, 2, lastRow - 1, 1).getValues().map(function (row) { return row[0]; });

  const rows = counts.map(function (count, i) {
    const change = i === 0 ? '' : (count - counts[i - 1]) / counts[i - 1];
    const window = counts.slice(Math.max(0, i - 2), i + 1);
    const moving = window.length < 3 ? '' : round1_(mean_(window));
    return [change, moving];
  });

  sheet.getRange(1, 6, 1, 2).setValues([['Изменение к прошлой неделе', 'Скользящее среднее (3 нед.)']])
    .setFontWeight('bold');
  sheet.getRange(2, 6, rows.length, 2).setValues(rows);
  sheet.getRange(2, 6, rows.length, 1).setNumberFormat('+0%;-0%;0%');
}

/**
 * Подсвечивает на листе «Заказы» заказы с необычным числом часов.
 * Правило межквартильного размаха (IQR): всё, что дальше чем 1,5 × IQR
 * от «средней половины» данных, считаем выбросом.
 */
function highlightHourOutliers() {
  const sheet = SpreadsheetApp.getActive().getSheetByName('Заказы');
  const lastRow = sheet.getLastRow();
  const range = sheet.getRange(2, 9, lastRow - 1, 1); // колонка I — часы
  const hours = range.getValues().map(function (row) { return row[0]; });

  const numbers = hours.filter(function (h) { return typeof h === 'number'; });
  const q1 = percentile_(numbers, 0.25);
  const q3 = percentile_(numbers, 0.75);
  const iqr = q3 - q1;
  const low = q1 - 1.5 * iqr;
  const high = q3 + 1.5 * iqr;

  let found = 0;
  const colors = hours.map(function (h) {
    const isOutlier = typeof h === 'number' && (h < low || h > high);
    if (isOutlier) found++;
    return [isOutlier ? '#f8d7da' : null]; // null — убрать заливку
  });
  range.setBackgrounds(colors);

  SpreadsheetApp.getActive().toast(
    'Норма: от ' + round1_(Math.max(low, 0)) + ' до ' + round1_(high) + ' ч. Выбросов: ' + found,
    'Выбросы по часам');
}
