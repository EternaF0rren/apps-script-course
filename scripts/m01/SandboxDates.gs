/**
 * Служебный скрипт песочницы.
 *
 * Данные песочницы «застыли» на дате выгрузки. Чтобы дедлайны и просрочки
 * выглядели живыми, сдвигаем все даты так, чтобы последний заказ
 * был создан час назад. Запускайте один раз после копирования таблицы.
 */
function refreshSandboxDates() {
  const ss = SpreadsheetApp.getActive();
  const orders = ss.getSheetByName('Заказы');
  const hours = ss.getSheetByName('Часы');

  const data = orders.getRange(2, 1, orders.getLastRow() - 1, orders.getLastColumn()).getValues();

  // Самая поздняя дата создания заказа (колонка B, индекс 1 в массиве)
  let latest = 0;
  data.forEach(function (row) {
    if (row[1] instanceof Date && row[1].getTime() > latest) latest = row[1].getTime();
  });
  if (!latest) throw new Error('В колонке «Создан» нет дат');

  const shiftMs = (Date.now() - 60 * 60 * 1000) - latest;
  const shift = function (value) {
    return value instanceof Date ? new Date(value.getTime() + shiftMs) : value;
  };

  // Колонки B (Создан), C (Дедлайн), K (Сдан) → индексы 1, 2, 10
  const shifted = data.map(function (row) {
    return [shift(row[1]), shift(row[2])];
  });
  const shiftedDone = data.map(function (row) {
    return [shift(row[10])];
  });
  orders.getRange(2, 2, shifted.length, 2).setValues(shifted);
  orders.getRange(2, 11, shiftedDone.length, 1).setValues(shiftedDone);

  // Лист «Часы»: недели сдвигаем на целое число недель, чтобы понедельники остались понедельниками
  const weekMs = 7 * 24 * 60 * 60 * 1000;
  const weekShift = Math.round(shiftMs / weekMs) * weekMs;
  const weeks = hours.getRange(2, 1, hours.getLastRow() - 1, 1).getValues().map(function (row) {
    return [row[0] instanceof Date ? new Date(row[0].getTime() + weekShift) : row[0]];
  });
  hours.getRange(2, 1, weeks.length, 1).setValues(weeks);

  ss.toast('Даты сдвинуты на ' + Math.round(shiftMs / 86400000) + ' дн.', 'Песочница', 5);
}
