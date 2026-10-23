/**
 * Модуль 1, задание со * 2. Сторож для Config.gs — эталонное решение.
 *
 * Сначала попробуйте написать сами, потом сравните.
 *
 * Нужен Config.gs из урока 1.4: объекты SHEETS и ORDER_COLS.
 * Ожидаемые заголовки храним рядом с номерами колонок — с теми же ключами.
 * Когда в Config.gs добавляется колонка, сюда добавляется её заголовок.
 */
const ORDER_HEADERS = {
  ID: 'Order ID',
  CREATED: 'Created',
  DEADLINE: 'Deadline',
  TYPE: 'Type',
  STATUS: 'Status',
  EMPLOYEE_ID: 'Assignee ID',
  EMPLOYEE: 'Assignee',
  SUPERVISOR: 'Supervisor',
  HOURS: 'Hours',
  QC: 'QC Score',
  DONE: 'Delivered',
  COMMENT: 'Comment',
};

/**
 * Проверяет, что таблица всё ещё такая, какой её ждёт код:
 * все листы из SHEETS на месте, а колонки листа «Orders» стоят там, где указано в ORDER_COLS.
 * Не останавливается на первой ошибке — собирает все и печатает списком.
 */
function checkConfig() {
  const ss = SpreadsheetApp.getActive();
  const problems = [];

  // 1. Листы. getSheet из урока остановил бы скрипт на первой ошибке,
  //    поэтому здесь getSheetByName и проверка на null
  Object.keys(SHEETS).forEach(function (key) {
    if (!ss.getSheetByName(SHEETS[key])) {
      problems.push('Sheet "' + SHEETS[key] + '" (SHEETS.' + key + ') not found');
    }
  });

  // 2. Колонки листа «Orders»
  const orders = ss.getSheetByName(SHEETS.ORDERS);
  if (orders) {
    const lastColumn = orders.getLastColumn();
    const headers = lastColumn > 0
      ? orders.getRange(1, 1, 1, lastColumn).getValues()[0]
      : [];

    Object.keys(ORDER_COLS).forEach(function (key) {
      const column = ORDER_COLS[key];
      const expected = ORDER_HEADERS[key];
      if (expected === undefined) {
        problems.push('ORDER_COLS.' + key + ': no expected header in ORDER_HEADERS');
        return;
      }
      const actual = column <= headers.length ? String(headers[column - 1]).trim() : '';
      if (actual !== expected) {
        problems.push('Orders, column ' + column + ' (ORDER_COLS.' + key + '): expected "' +
          expected + '", found "' + (actual || 'empty') + '"');
      }
    });
  }

  if (problems.length === 0) {
    console.log('Config OK');
  } else {
    console.log('Config problems: ' + problems.length + '\n- ' + problems.join('\n- '));
  }
  return problems; // пригодится в модуле 3, чтобы показать итог во всплывающем окне
}
