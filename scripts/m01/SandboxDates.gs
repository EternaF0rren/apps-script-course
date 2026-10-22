/**
 * Служебный скрипт песочницы.
 *
 * Данные песочницы «застыли» на дате выгрузки. Чтобы дедлайны и просрочки
 * выглядели живыми, сдвигаем все даты так, чтобы последний заказ
 * был создан час назад. Запускайте один раз после копирования таблицы.
 */
function refreshSandboxDates() {
  const ss = SpreadsheetApp.getActive();

  // Ячейки показывают время в часовом поясе ТАБЛИЦЫ (Файл → Настройки),
  // а это отдельная настройка от часового пояса проекта скриптов.
  // У таблицы, созданной из xlsx, по умолчанию стоит GMT — тогда время в ячейках
  // будет сдвинуто на несколько часов. Поэтому сначала сверяем два пояса.
  const sheetZone = ss.getSpreadsheetTimeZone();
  const scriptZone = Session.getScriptTimeZone();
  if (sheetZone !== scriptZone) {
    throw new Error('Time zones differ: spreadsheet — ' + sheetZone + ', script project — ' + scriptZone +
      '. Set the same time zone in File → Settings and in Project Settings, then run again.');
  }
  const orders = ss.getSheetByName('Orders');
  const hours = ss.getSheetByName('Hours');

  const data = orders.getRange(2, 1, orders.getLastRow() - 1, orders.getLastColumn()).getValues();

  // Самая поздняя дата создания заказа (колонка B, индекс 1 в массиве)
  let latest = 0;
  data.forEach(function (row) {
    if (row[1] instanceof Date && row[1].getTime() > latest) latest = row[1].getTime();
  });
  if (!latest) throw new Error('No dates in the Created column');

  const shiftMs = (Date.now() - 60 * 60 * 1000) - latest;
  const shift = function (value) {
    return value instanceof Date ? new Date(value.getTime() + shiftMs) : value;
  };

  // Колонки B (Created), C (Deadline), K (Delivered) → индексы 1, 2, 10
  const shifted = data.map(function (row) {
    return [shift(row[1]), shift(row[2])];
  });
  const shiftedDone = data.map(function (row) {
    return [shift(row[10])];
  });
  orders.getRange(2, 2, shifted.length, 2).setValues(shifted);
  orders.getRange(2, 11, shiftedDone.length, 1).setValues(shiftedDone);

  // Лист «Hours»: недели сдвигаем на целое число недель, чтобы понедельники остались понедельниками
  const weekMs = 7 * 24 * 60 * 60 * 1000;
  const weekShift = Math.round(shiftMs / weekMs) * weekMs;
  const weeks = hours.getRange(2, 1, hours.getLastRow() - 1, 1).getValues().map(function (row) {
    return [row[0] instanceof Date ? new Date(row[0].getTime() + weekShift) : row[0]];
  });
  hours.getRange(2, 1, weeks.length, 1).setValues(weeks);

  ss.toast('Dates shifted by ' + Math.round(shiftMs / 86400000) + ' days', 'Sandbox', 5);
}
