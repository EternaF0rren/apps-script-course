/**
 * Модуль 2, урок 2. Читаем код макроса.
 *
 * Так выглядит макрос, который Google Таблицы записали сами:
 * мы выделили шапку, залили её тёмным цветом, сделали шрифт белым и жирным
 * и закрепили первую строку.
 */
function FormatHeader() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A1:L1').activate();
  spreadsheet.getActiveRangeList().setBackground('#1f2a44')
  .setFontColor('#ffffff')
  .setFontWeight('bold');
  spreadsheet.getActiveSheet().setFrozenRows(1);
};

/**
 * Тот же макрос, переписанный руками.
 * Что поменяли и почему:
 *  1. Нет activate(): скрипту не нужно «выделять» ячейки, как человеку. Так быстрее.
 *  2. Диапазон шапки считается по числу колонок, а не жёстко A1:L1.
 *     Добавили колонку — макрос всё равно покрасит всю шапку.
 *  3. Цвет вынесен в переменную: поменять стиль можно в одном месте.
 */
function formatHeader() {
  const HEADER_COLOR = '#1f2a44';
  const sheet = SpreadsheetApp.getActiveSheet();
  const header = sheet.getRange(1, 1, 1, sheet.getLastColumn());

  header
    .setBackground(HEADER_COLOR)
    .setFontColor('#ffffff')
    .setFontWeight('bold');
  sheet.setFrozenRows(1);
}
