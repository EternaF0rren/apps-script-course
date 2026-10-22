/**
 * Модуль 2, урок 3. Форматирование скриптом.
 *
 * Одна функция приводит лист «Orders» к единому виду:
 * шапка, закрепление, чередование строк, форматы дат и чисел,
 * подсветка просроченных и проблемных заказов.
 * Запускать можно сколько угодно раз — результат будет одинаковым.
 */

const STYLE = {
  FONT: 'Arial',
  HEADER_BG: '#1f2a44',
  HEADER_FONT: '#ffffff',
  OVERDUE_BG: '#fde2e1',   // просрочено — светло-красный
  PROBLEM_FONT: '#c62828', // статус «Problem» — красный текст
  DONE_FONT: '#9e9e9e',    // готовые — серым, чтобы не отвлекали
};

function formatOrdersSheet() {
  const sheet = SpreadsheetApp.getActive().getSheetByName('Orders');
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  const body = sheet.getRange(2, 1, lastRow - 1, lastCol);

  // 1. Шрифт для всего листа
  sheet.getDataRange().setFontFamily(STYLE.FONT).setFontSize(10);

  // 2. Шапка
  sheet.getRange(1, 1, 1, lastCol)
    .setBackground(STYLE.HEADER_BG)
    .setFontColor(STYLE.HEADER_FONT)
    .setFontWeight('bold');
  sheet.setFrozenRows(1);

  // 3. Чередование строк. Старое чередование сначала убираем:
  //    второй раз на тот же диапазон его не наложить — будет ошибка.
  sheet.getBandings().forEach(function (b) { b.remove(); });
  body.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, false, false);

  // 4. Форматы: даты (B, C, K), часы (I) с одним знаком после запятой
  ['B', 'C', 'K'].forEach(function (col) {
    sheet.getRange(col + '2:' + col + lastRow).setNumberFormat('dd.MM.yyyy HH:mm');
  });
  sheet.getRange('I2:I' + lastRow).setNumberFormat('0.0');

  // 5. Условное форматирование. В формулах — английские имена функций
  //    и запятые: так их понимает Apps Script при любом языке таблицы.
  const overdue = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($C2<NOW(), $E2<>"Done")')
    .setBackground(STYLE.OVERDUE_BG)
    .setRanges([body])
    .build();

  const problem = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Problem')
    .setFontColor(STYLE.PROBLEM_FONT)
    .setBold(true)
    .setRanges([sheet.getRange('E2:E' + lastRow)])
    .build();

  const done = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$E2="Done"')
    .setFontColor(STYLE.DONE_FONT)
    .setRanges([body])
    .build();

  // Порядок важен: первое подходящее правило побеждает
  sheet.setConditionalFormatRules([problem, overdue, done]);

  // 6. Ширина колонок под содержимое
  sheet.autoResizeColumns(1, lastCol);

  SpreadsheetApp.getActive().toast('Orders sheet formatted', 'Done', 3);
}

/**
 * Задание *: единый стиль для всех листов книги, кроме служебных.
 * Применяет шрифт и оформление шапки; данные не трогает.
 */
function applyStyleToAllSheets() {
  const SKIP = ['Help', 'Settings'];

  SpreadsheetApp.getActive().getSheets().forEach(function (sheet) {
    if (SKIP.indexOf(sheet.getName()) !== -1) return; // пропускаем служебные
    if (sheet.getLastColumn() === 0) return;           // пустой лист

    sheet.getDataRange().setFontFamily(STYLE.FONT).setFontSize(10);
    sheet.getRange(1, 1, 1, sheet.getLastColumn())
      .setBackground(STYLE.HEADER_BG)
      .setFontColor(STYLE.HEADER_FONT)
      .setFontWeight('bold');
    sheet.setFrozenRows(1);
  });
}
