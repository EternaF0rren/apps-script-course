/**
 * Модуль 8, урок 5. Дашборд: главные цифры недели на одном листе.
 *
 * Нужны файлы уроков 2–4 (Stats, Summary, Trends, Charts)
 * и функция removeTriggersFor_ из модуля 5 (файл «Triggers»).
 */
function buildDashboard() {
  buildSummary();
  addTrendColumns();

  const ss = SpreadsheetApp.getActive();
  const dash = ss.getSheetByName('Dashboard') || ss.insertSheet('Dashboard');
  dash.clearContents(); // значения стираем, графики остаются на месте

  const tz = ss.getSpreadsheetTimeZone();
  dash.getRange('A1').setValue('Orders Dashboard').setFontSize(16).setFontWeight('bold');
  dash.getRange('A2').setValue('Updated: ' + Utilities.formatDate(new Date(), tz, 'dd.MM.yyyy HH:mm'));

  // --- Главные цифры: последняя полная неделя против предыдущей ---
  const weeks = ss.getSheetByName('Summary').getRange('A2:E').getValues()
    .filter(function (row) { return row[0] !== ''; });
  const last = weeks[weeks.length - 1];
  const prev = weeks[weeks.length - 2];

  const kpi = [
    ['Metric', 'Week of ' + Utilities.formatDate(last[0], tz, 'dd.MM'), 'Week Before', 'Change'],
    kpiRow_('Orders', last[1], prev[1]),
    kpiRow_('Avg Hours', last[2], prev[2]),
    kpiRow_('Median Hours', last[3], prev[3]),
    kpiRow_('Overdue Share', last[4], prev[4], true),
  ];
  dash.getRange(4, 1, kpi.length, 4).setValues(kpi);
  dash.getRange(4, 1, 1, 4).setFontWeight('bold').setBackground('#1f2a44').setFontColor('#ffffff');
  dash.getRange(8, 2, 1, 2).setNumberFormat('0%');
  dash.getRange(5, 4, 3, 1).setNumberFormat('+0%;-0%;0%');
  dash.getRange(8, 4).setNumberFormat('+0" pp";-0" pp";0" pp"');

  // --- Что происходит прямо сейчас ---
  const orders = readOrders_();
  const open = orders.filter(function (o) { return o.status !== 'Done'; });
  const overdueNow = open.filter(function (o) { return o.late; });
  dash.getRange(11, 1, 3, 2).setValues([
    ['Right Now', ''],
    ['Open Orders', open.length],
    ['Of Them Overdue', overdueNow.length],
  ]);
  dash.getRange(11, 1).setFontWeight('bold');

  // --- Топ-5 исполнителей по числу просрочек ---
  const byWorker = {};
  orders.forEach(function (o) {
    if (!byWorker[o.worker]) byWorker[o.worker] = { orders: 0, late: 0 };
    byWorker[o.worker].orders++;
    if (o.late) byWorker[o.worker].late++;
  });
  const top = Object.keys(byWorker)
    .map(function (name) { return [name, byWorker[name].late, byWorker[name].orders]; })
    .sort(function (a, b) { return b[1] - a[1] || a[0].localeCompare(b[0]); })
    .slice(0, 5);
  dash.getRange(15, 1, 1, 3).setValues([['Most Overdue', 'Overdue', 'Orders']]).setFontWeight('bold');
  dash.getRange(16, 1, top.length, 3).setValues(top);

  buildCharts();
}

/**
 * Строка показателя: значение сейчас, неделей раньше и изменение.
 * Для долей (isShare = true) изменение — в процентных пунктах:
 * было 20%, стало 25% → «+5 pp», а не «+25%».
 */
function kpiRow_(label, current, previous, isShare) {
  let change = '';
  if (isShare) change = Math.round((current - previous) * 100);
  else if (previous) change = (current - previous) / previous;
  return [label, current, previous, change];
}

/** Обновлять дашборд каждый понедельник в 7 утра — к началу рабочей недели. */
function createDashboardTrigger() {
  removeTriggersFor_('buildDashboard');
  ScriptApp.newTrigger('buildDashboard')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(7)
    .create();
  SpreadsheetApp.getActive().toast('The dashboard will update on Mondays at 7:00');
}
