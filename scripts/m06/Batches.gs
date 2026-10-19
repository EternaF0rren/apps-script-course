/**
 * Модуль 6, задание со *. Большие объёмы: обработка частями.
 *
 * Один запуск скрипта длится не больше 6 минут. Если строк очень много,
 * обрабатываем их порциями и запоминаем, где остановились, в свойствах скрипта.
 * Следующий запуск (вручную или триггером раз в несколько минут) продолжает с того же места.
 */
const BATCH_SIZE = 50;

function processInBatches() {
  const props = PropertiesService.getScriptProperties();
  const sheet = SpreadsheetApp.getActive().getSheetByName('Заказы');
  const lastRow = sheet.getLastRow();

  let startRow = Number(props.getProperty('NEXT_ROW') || 2);
  if (startRow > lastRow) {
    props.deleteProperty('NEXT_ROW');
    console.log('Все строки обработаны, начинаем сначала при следующем запуске');
    return;
  }

  const n = Math.min(BATCH_SIZE, lastRow - startRow + 1);
  const range = sheet.getRange(startRow, 1, n, 12);
  const values = range.getValues();

  // Здесь — любая тяжёлая обработка порции. Для примера: ID в верхний регистр.
  const processed = values.map(function (row) {
    row[0] = String(row[0]).toUpperCase();
    return row;
  });
  range.setValues(processed);

  props.setProperty('NEXT_ROW', String(startRow + n));
  console.log('Обработаны строки ' + startRow + '–' + (startRow + n - 1));
}
