/**
 * Модуль 8, урок 4. Графики скриптом.
 *
 * Графики строятся на листе «Dashboard» по данным листа «Summary».
 * Повторный запуск не плодит копии: если график с таким заголовком уже есть,
 * у него просто обновляется диапазон данных.
 */
function buildCharts() {
  const ss = SpreadsheetApp.getActive();
  const summary = ss.getSheetByName('Summary');
  const dash = ss.getSheetByName('Dashboard') || ss.insertSheet('Dashboard');

  const weeks = filledRows_(summary, 1);  // строк в блоке по неделям, вместе с шапкой
  const types = filledRows_(summary, 10); // строк в блоке по типам, вместе с шапкой

  // 1. Линейный: заказы по неделям и скользящее среднее
  upsertChart_(dash, 'Orders by Week', function (builder) {
    builder.setChartType(Charts.ChartType.LINE)
      .addRange(summary.getRange(1, 1, weeks, 2)) // A — недели, B — заказы
      .addRange(summary.getRange(1, 7, weeks, 1)) // G — скользящее среднее
      .setPosition(2, 7, 0, 0);
  });

  // 2. Столбчатый: заказы по типам
  upsertChart_(dash, 'Orders by Type', function (builder) {
    builder.setChartType(Charts.ChartType.COLUMN)
      .addRange(summary.getRange(1, 10, types, 2)) // J — тип, K — заказов
      .setOption('legend', { position: 'none' })
      .setPosition(20, 7, 0, 0);
  });

  // 3. Комбинированный: заказы столбиками, средние часы линией на правой оси
  upsertChart_(dash, 'Orders and Avg Hours', function (builder) {
    builder.setChartType(Charts.ChartType.COMBO)
      .addRange(summary.getRange(1, 1, weeks, 3)) // A — недели, B — заказы, C — средние часы
      .setOption('seriesType', 'bars')
      .setOption('series', { 1: { type: 'line', targetAxisIndex: 1 } })
      .setPosition(20, 14, 0, 0);
  });
}

/**
 * Находит график по заголовку и обновляет его, а если такого нет — создаёт.
 * configure получает «конструктор» графика и настраивает тип, данные и место.
 */
function upsertChart_(sheet, title, configure) {
  const existing = sheet.getCharts().filter(function (chart) {
    return chart.getOptions().get('title') === title;
  })[0];

  const builder = existing ? existing.modify().clearRanges() : sheet.newChart();
  configure(builder);
  builder.setNumHeaders(1) // первая строка диапазона — названия рядов
    .setOption('title', title);

  if (existing) sheet.updateChart(builder.build());
  else sheet.insertChart(builder.build());
}

/** Сколько заполненных ячеек подряд в колонке, начиная с первой строки. */
function filledRows_(sheet, column) {
  const values = sheet.getRange(1, column, sheet.getLastRow() || 1, 1).getValues();
  let count = 0;
  while (count < values.length && values[count][0] !== '') count++;
  return count;
}
