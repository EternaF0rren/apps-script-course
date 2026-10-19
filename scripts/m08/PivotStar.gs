/**
 * Модуль 8, задание со *. Сводная таблица скриптом.
 *
 * Создаёт лист «Сводная»: строки — типы заказов, колонки — статусы,
 * в ячейках — число заказов и сумма часов. Это обычная сводная таблица:
 * после создания её можно настраивать мышкой в редакторе сводных таблиц.
 */
function createOrdersPivot() {
  const ss = SpreadsheetApp.getActive();
  const orders = ss.getSheetByName('Заказы');
  const source = orders.getRange(1, 1, orders.getLastRow(), 12); // A1:L с шапкой

  const old = ss.getSheetByName('Сводная');
  if (old) ss.deleteSheet(old); // сводная пересоздаётся при каждом запуске
  const sheet = ss.insertSheet('Сводная');

  const pivot = sheet.getRange('A1').createPivotTable(source);
  // Номера колонок — внутри исходного диапазона, с 1: D — 4-я, E — 5-я, A — 1-я, I — 9-я
  pivot.addRowGroup(4);    // строки: тип
  pivot.addColumnGroup(5); // колонки: статус
  pivot.addPivotValue(1, SpreadsheetApp.PivotTableSummarizeFunction.COUNTA).setDisplayName('Заказов');
  pivot.addPivotValue(9, SpreadsheetApp.PivotTableSummarizeFunction.SUM).setDisplayName('Часов');
}
