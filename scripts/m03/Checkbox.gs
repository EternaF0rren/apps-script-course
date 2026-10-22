/**
 * Модуль 3, урок 4. Чекбокс-кнопка.
 *
 * Кнопки-рисунки не работают в мобильном приложении Таблиц, а флажки работают.
 * Ставим в каждой строке флажок «Закрыть»: отметили — заказ закрылся сам.
 */

const CLOSE_COL = 13; // M — колонка с флажками

/** Один раз: добавить колонку «Close» с флажками на лист «Orders». */
function addCloseCheckboxes() {
  const sheet = SpreadsheetApp.getActive().getSheetByName('Orders');
  const lastRow = sheet.getLastRow();
  sheet.getRange(1, CLOSE_COL).setValue('Close');
  sheet.getRange(2, CLOSE_COL, lastRow - 1, 1).insertCheckboxes();
  SpreadsheetApp.getActive().toast('Checkboxes added to column M', 'Done', 3);
}

/**
 * onEdit — ещё одно особое имя. Таблица запускает её при КАЖДОЙ ручной правке
 * и передаёт объект e: что, где и на что поменяли.
 *
 * Важно: функция onEdit в проекте может быть только одна. Поэтому сама она
 * ничего не делает, а только раздаёт событие обработчикам. В модуле 5
 * сюда добавятся новые обработчики — по строчке на каждый.
 */
function onEdit(e) {
  handleCloseCheckbox_(e);
}

function handleCloseCheckbox_(e) {
  const range = e.range;
  const sheet = range.getSheet();

  // Реагируем только на флажок в колонке M листа «Orders», и только когда его ставят
  if (sheet.getName() !== 'Orders') return;
  if (range.getColumn() !== CLOSE_COL || range.getRow() < 2) return;
  if (e.value !== 'TRUE') return;

  const row = range.getRow();
  sheet.getRange(row, 5).setValue('Done');      // E — Status
  sheet.getRange(row, 11).setValue(new Date());  // K — Delivered
}
