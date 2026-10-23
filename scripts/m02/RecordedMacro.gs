/**
 * Модуль 2, урок 2. Читаем код макроса.
 *
 * Так выглядит макрос, который Google Таблицы записали сами:
 * мы выделили шапку, залили её синим, поменяли шрифт на IBM Plex Sans,
 * сделали его курсивом и выровняли текст по центру.
 */
function Formatheader() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A1:L1').activate();
  spreadsheet.getActiveRangeList().setBackground('#1155cc')
  .setFontFamily('IBM Plex Sans')
  .setFontStyle('italic')
  .setHorizontalAlignment('center');
};

/**
 * Тот же макрос, переписанный руками.
 * Что поменяли и почему:
 *  1. Нет activate(): скрипту не нужно «выделять» ячейки, как человеку. Так быстрее.
 *  2. Диапазон шапки считается по числу колонок, а не жёстко A1:L1.
 *     Добавили колонку — макрос всё равно оформит всю шапку.
 *  3. Цвет и шрифт вынесены в константы: поменять стиль можно в одном месте.
 */
function formatHeader() {
  const HEADER_COLOR = '#1155cc';
  const HEADER_FONT = 'IBM Plex Sans';
  const sheet = SpreadsheetApp.getActiveSheet();
  const header = sheet.getRange(1, 1, 1, sheet.getLastColumn());

  header
    .setBackground(HEADER_COLOR)
    .setFontFamily(HEADER_FONT)
    .setFontStyle('italic')
    .setHorizontalAlignment('center');
}
