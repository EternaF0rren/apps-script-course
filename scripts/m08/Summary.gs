/**
 * Модуль 8, урок 2. Метрики по неделям и по типам заказов — на лист «Сводка».
 *
 * Лист «Сводка» после запуска:
 *   A–E — по неделям (только полные недели),
 *   F–G — тренды, их добавляет addTrendColumns из урока 3,
 *   J–P — по типам заказов.
 * Функция полностью перезаписывает лист, в том числе сводку из модуля 6.
 */
function buildSummary() {
  const orders = readOrders_();
  const sheet = SpreadsheetApp.getActive().getSheetByName('Сводка');
  sheet.clearContents();

  writeWeeks_(sheet, orders);
  writeTypes_(sheet, orders);
  SpreadsheetApp.getActive().toast('Сводка обновлена', 'Модуль 8');
}

/** Понедельник той недели, в которую попадает дата. */
function weekStart_(date) {
  const daysFromMonday = (date.getDay() + 6) % 7; // пн → 0, вс → 6
  return new Date(date.getFullYear(), date.getMonth(), date.getDate() - daysFromMonday);
}

/** Блок «по неделям» в A1:E. */
function writeWeeks_(sheet, orders) {
  const now = new Date();
  const byWeek = {};
  orders.forEach(function (order) {
    if (!(order.created instanceof Date)) return;
    const start = weekStart_(order.created);
    const key = start.getTime();
    if (!byWeek[key]) byWeek[key] = { start: start, hours: [], late: 0 };
    byWeek[key].hours.push(order.hours);
    if (order.late) byWeek[key].late++;
  });

  const weekMs = 7 * 24 * 60 * 60 * 1000;
  const rows = Object.keys(byWeek)
    .map(Number)
    .sort(function (a, b) { return a - b; })
    // Неполную текущую неделю не берём: в понедельник она всегда выглядит провалом
    .filter(function (key) { return key + weekMs <= now.getTime(); })
    .map(function (key) {
      const w = byWeek[key];
      return [w.start, w.hours.length, round1_(mean_(w.hours)), round1_(median_(w.hours)), w.late / w.hours.length];
    });

  sheet.getRange(1, 1, 1, 5)
    .setValues([['Неделя с', 'Заказов', 'Средние часы', 'Медиана часов', 'Доля просрочек']])
    .setFontWeight('bold');
  if (rows.length === 0) return;
  sheet.getRange(2, 1, rows.length, 5).setValues(rows);
  sheet.getRange(2, 1, rows.length, 1).setNumberFormat('dd.mm');
  sheet.getRange(2, 5, rows.length, 1).setNumberFormat('0%');
}

/** Блок «по типам заказов» в J1:P. */
function writeTypes_(sheet, orders) {
  const byType = {};
  orders.forEach(function (order) {
    if (!byType[order.type]) byType[order.type] = { hours: [], qc: [], late: 0 };
    const t = byType[order.type];
    t.hours.push(order.hours);
    if (typeof order.qc === 'number') t.qc.push(order.qc);
    if (order.late) t.late++;
  });

  const rows = Object.keys(byType).sort().map(function (type) {
    const t = byType[type];
    return [
      type,
      t.hours.length,
      round1_(mean_(t.hours)),
      round1_(median_(t.hours)),
      round1_(percentile_(t.hours, 0.9)),
      round1_(mean_(t.qc)),
      t.late / t.hours.length,
    ];
  });

  sheet.getRange(1, 10, 1, 7)
    .setValues([['Тип', 'Заказов', 'Средние часы', 'Медиана часов', '90-й перцентиль', 'Средняя оценка QC', 'Доля просрочек']])
    .setFontWeight('bold');
  sheet.getRange(2, 10, rows.length, 7).setValues(rows);
  sheet.getRange(2, 16, rows.length, 1).setNumberFormat('0%');
}
