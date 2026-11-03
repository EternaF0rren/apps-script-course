/**
 * Модуль 6, задание со * 1. Кто перегружен прямо сейчас — эталонное решение.
 *
 * Сначала попробуйте написать сами, потом сравните.
 *
 * buildWorkload собирает на лист «Workload» нагрузку по исполнителям:
 * сколько у каждого незакрытых заказов, сколько из них в статусе «Problem»
 * и когда ближайший дедлайн. Сверху — самые загруженные.
 */
const WORKLOAD_SHEET = 'Workload';
const WORKLOAD_LIMIT = 3;          // столько открытых заказов и больше — подсвечиваем
const WORKLOAD_HIGHLIGHT = '#fde2e1';

function buildWorkload() {
  const ss = SpreadsheetApp.getActive();
  const values = ss.getSheetByName('Orders').getDataRange().getValues().slice(1);

  // Словарь по ID исполнителя: { 'D-004': { name, supervisor, open, problem, nearest }, ... }
  const byWorker = {};
  values.forEach(function (row) {
    const id = row[5], status = row[4], deadline = row[2];
    if (!id || status === 'Done') return; // пустые строки и закрытые заказы не считаем

    if (!byWorker[id]) {
      byWorker[id] = { name: row[6], supervisor: row[7], open: 0, problem: 0, nearest: null };
    }
    const w = byWorker[id];
    w.open++;
    if (status === 'Problem') w.problem++;
    if (deadline instanceof Date && (!w.nearest || deadline < w.nearest)) w.nearest = deadline;
  });

  // Самые загруженные сверху; при равенстве — у кого дедлайн раньше
  const ids = Object.keys(byWorker).sort(function (a, b) {
    const x = byWorker[a], y = byWorker[b];
    if (y.open !== x.open) return y.open - x.open;
    return (x.nearest || Infinity) - (y.nearest || Infinity);
  });
  const rows = ids.map(function (id) {
    const w = byWorker[id];
    return [id, w.name, w.supervisor, w.open, w.problem, w.nearest || ''];
  });

  // Лист создаём, если его нет; старое содержимое и подсветку стираем
  const out = ss.getSheetByName(WORKLOAD_SHEET) || ss.insertSheet(WORKLOAD_SHEET);
  out.clear();
  out.getRange(1, 1, 1, 6)
    .setValues([['Assignee ID', 'Assignee', 'Supervisor', 'Open Orders', 'Problems', 'Nearest Deadline']])
    .setFontWeight('bold');
  out.setFrozenRows(1);

  if (rows.length === 0) {
    out.getRange(2, 1).setValue('No open orders');
    return;
  }
  out.getRange(2, 1, rows.length, 6).setValues(rows);
  out.getRange(2, 6, rows.length, 1).setNumberFormat('dd.MM HH:mm');

  let overloaded = 0;
  rows.forEach(function (row, i) {
    if (row[3] >= WORKLOAD_LIMIT) {
      out.getRange(i + 2, 1, 1, 6).setBackground(WORKLOAD_HIGHLIGHT);
      overloaded++;
    }
  });
  out.autoResizeColumns(1, 6);

  const totalOpen = rows.reduce(function (sum, row) { return sum + row[3]; }, 0);
  ss.toast('Open orders: ' + totalOpen + ', assignees with ' + WORKLOAD_LIMIT + '+ open: ' + overloaded,
    'Workload', 5);
}
