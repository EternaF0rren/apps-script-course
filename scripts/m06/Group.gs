/**
 * Модуль 6, урок 5. Группировка: объект как словарь.
 *
 * Считаем по каждому исполнителю: сколько заказов, сколько часов, сколько
 * просрочек. Это то, что делает сводная таблица, — только с любыми правилами.
 * Результат — на лист «Сводка». В модуле 8 к нему добавятся графики.
 */
function summarizeByWorker() {
  const ss = SpreadsheetApp.getActive();
  const values = ss.getSheetByName('Заказы').getDataRange().getValues().slice(1);
  const now = new Date();

  // Словарь: { 'Ольга Л.': { orders: 7, hours: 31.5, late: 1 }, ... }
  const byWorker = {};
  values.forEach(function (row) {
    const name = row[6];
    if (!byWorker[name]) byWorker[name] = { orders: 0, hours: 0, late: 0 };

    const stat = byWorker[name];
    stat.orders++;
    stat.hours += row[8];

    const deadline = row[2], done = row[10];
    const isLate = done instanceof Date ? done > deadline : deadline < now;
    if (isLate) stat.late++;
  });

  const rows = Object.keys(byWorker).sort().map(function (name) {
    const s = byWorker[name];
    return [name, s.orders, Math.round(s.hours * 10) / 10, s.late, s.late / s.orders];
  });

  const out = ss.getSheetByName('Сводка');
  out.clearContents();
  out.getRange(1, 1, 1, 5).setValues([['Исполнитель', 'Заказов', 'Часов', 'Просрочек', 'Доля просрочек']])
    .setFontWeight('bold');
  out.getRange(2, 1, rows.length, 5).setValues(rows);
  out.getRange(2, 5, rows.length, 1).setNumberFormat('0%');
}

/** Короче: сколько заказов в каждом статусе. */
function countByStatus() {
  const values = SpreadsheetApp.getActive().getSheetByName('Заказы').getDataRange().getValues().slice(1);
  const counts = {};
  values.forEach(function (row) { counts[row[4]] = (counts[row[4]] || 0) + 1; });
  console.log(counts); // { 'Готов': 170, 'В работе': 11, ... }
}
