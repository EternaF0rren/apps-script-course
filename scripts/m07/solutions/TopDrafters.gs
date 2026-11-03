/**
 * Модуль 7, задание со * 2. Рейтинг исполнителей одной формулой — эталонное решение.
 *
 * Сначала попробуйте написать сами, потом сравните.
 *
 *   =TOP_DRAFTERS(Orders!A2:L)          → десять лучших, у кого не меньше 5 закрытых заказов
 *   =TOP_DRAFTERS(Orders!A2:L; 5; 10)   → пятеро лучших среди тех, у кого 10 заказов и больше
 *
 * Формула возвращает таблицу с шапкой: она сама занимает нужное число строк и колонок.
 * Диапазон заказов передаётся аргументом, поэтому рейтинг пересчитывается
 * при любой правке на листе «Orders» (урок 7.3).
 */

// Номера колонок листа «Orders», считая с 0: A — 0, C — 2, …
const TOP_COL = { DEADLINE: 2, STATUS: 4, ASSIGNEE: 6, QC: 9, DELIVERED: 10 };

/**
 * Рейтинг исполнителей по закрытым заказам: средняя оценка QC и доля сданных в срок.
 *
 * @param {Array} orders Диапазон листа «Orders», например Orders!A2:L.
 * @param {number} top Необязательно: сколько строк показать. По умолчанию 10.
 * @param {number} minOrders Необязательно: сколько закрытых заказов нужно, чтобы попасть в рейтинг. По умолчанию 5.
 * @return Таблица: Assignee, Done orders, Avg QC, On time.
 * @customfunction
 */
function TOP_DRAFTERS(orders, top, minOrders) {
  if (!Array.isArray(orders) || orders[0].length <= TOP_COL.DELIVERED) {
    throw new Error('TOP_DRAFTERS: pass the Orders range with columns A to K or wider, for example Orders!A2:L');
  }
  const limit = typeof top === 'number' && top > 0 ? top : 10;
  const minCount = typeof minOrders === 'number' ? minOrders : 5;

  // 1. Словарь по исполнителю: считаем только закрытые заказы
  const stats = {};
  orders.forEach(function (row) {
    const name = row[TOP_COL.ASSIGNEE];
    const deadline = row[TOP_COL.DEADLINE];
    const delivered = row[TOP_COL.DELIVERED];
    if (!name || row[TOP_COL.STATUS] !== 'Done') return; // пустые строки, шапка и открытые заказы

    const s = stats[name] || (stats[name] = { done: 0, onTime: 0, qcSum: 0, qcCount: 0 });
    s.done++;
    if (deadline instanceof Date && delivered instanceof Date && delivered <= deadline) s.onTime++;
    const qc = row[TOP_COL.QC];
    if (typeof qc === 'number') { // пустую оценку не считаем нулём
      s.qcSum += qc;
      s.qcCount++;
    }
  });

  // 2. Строки рейтинга: только те, у кого достаточно заказов и есть хоть одна оценка
  const rows = Object.keys(stats)
    .filter(function (name) { return stats[name].done >= minCount && stats[name].qcCount > 0; })
    .map(function (name) {
      const s = stats[name];
      return [name, s.done, Math.round(s.qcSum / s.qcCount * 100) / 100, Math.round(s.onTime / s.done * 100) / 100];
    });

  // 3. Сверху — лучшая средняя оценка, при равенстве — больше заказов
  rows.sort(function (a, b) { return b[2] - a[2] || b[1] - a[1]; });

  if (rows.length === 0) return [['No one has ' + minCount + ' or more done orders']];
  return [['Assignee', 'Done orders', 'Avg QC', 'On time']].concat(rows.slice(0, limit));
}
