/**
 * Модуль 8, уроки 1–2. Метрики: среднее, медиана, перцентили, разброс.
 *
 * Помощники с «_» в конце — общие для всего модуля: их используют
 * сводка, тренды, дашборд и отчёт.
 */

/** Среднее арифметическое. Как СРЗНАЧ. */
function mean_(numbers) {
  if (numbers.length === 0) return '';
  return numbers.reduce(function (sum, x) { return sum + x; }, 0) / numbers.length;
}

/**
 * Перцентиль: p = 0.5 — медиана, p = 0.9 — значение, ниже которого 90% данных.
 * Считает так же, как ПРОЦЕНТИЛЬ (PERCENTILE) в Таблицах.
 */
function percentile_(numbers, p) {
  if (numbers.length === 0) return '';
  const sorted = numbers.slice().sort(function (a, b) { return a - b; });
  const position = (sorted.length - 1) * p;
  const lower = Math.floor(position);
  const upper = Math.ceil(position);
  return sorted[lower] + (sorted[upper] - sorted[lower]) * (position - lower);
}

/** Медиана: середина ряда. Как МЕДИАНА. */
function median_(numbers) {
  return percentile_(numbers, 0.5);
}

/** Стандартное отклонение по выборке — насколько значения разбросаны. Как СТАНДОТКЛОН. */
function stdev_(numbers) {
  if (numbers.length < 2) return '';
  const m = mean_(numbers);
  const squares = numbers.reduce(function (sum, x) { return sum + (x - m) * (x - m); }, 0);
  return Math.sqrt(squares / (numbers.length - 1));
}

/** Округление до одного знака; пустое значение оставляем пустым. */
function round1_(x) {
  return x === '' ? '' : Math.round(x * 10) / 10;
}

/**
 * Читает лист «Orders» и возвращает список заказов-объектов.
 * С именами полей код читается проще, чем с row[8] и row[10].
 */
function readOrders_() {
  const values = SpreadsheetApp.getActive().getSheetByName('Orders').getDataRange().getValues().slice(1);
  const now = new Date();
  return values
    .filter(function (row) { return row[0] !== ''; })
    .map(function (row) {
      const deadline = row[2], done = row[10];
      return {
        id: row[0],
        created: row[1],
        deadline: deadline,
        type: row[3],
        status: row[4],
        worker: row[6],
        hours: row[8],
        qc: row[9],
        done: done,
        // Просрочен: сдан позже дедлайна или не сдан, а дедлайн уже прошёл (как в модуле 6)
        late: done instanceof Date ? done > deadline : deadline < now,
      };
    });
}

/**
 * Урок 1. Описание часов на заказ: запустите и посмотрите в журнал.
 * Сравните результат с формулами в любой пустой ячейке:
 *   =СРЗНАЧ(Orders!I2:I)  =МЕДИАНА(Orders!I2:I)
 *   =ПРОЦЕНТИЛЬ(Orders!I2:I; 0,9)  =СТАНДОТКЛОН(Orders!I2:I)
 */
function describeHours() {
  const hours = readOrders_()
    .map(function (order) { return order.hours; })
    .filter(function (h) { return typeof h === 'number'; });

  console.log('Orders:          ' + hours.length);
  console.log('Mean:            ' + round1_(mean_(hours)) + ' h');
  console.log('Median:          ' + round1_(median_(hours)) + ' h');
  console.log('90th percentile: ' + round1_(percentile_(hours, 0.9)) + ' h');
  console.log('Spread (σ):      ' + round1_(stdev_(hours)) + ' h');
  console.log('Min / max:       ' + Math.min.apply(null, hours) + ' / ' + Math.max.apply(null, hours) + ' h');
}
