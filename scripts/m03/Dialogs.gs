/**
 * Модуль 3, урок 5. Диалоги и боковая панель.
 *
 * Окна работают, только когда функцию запустили из таблицы (меню, кнопка).
 * При запуске из редактора будет ошибка «Cannot call SpreadsheetApp.getUi()
 * from this context» — это нормально.
 */

/** Спросить ID исполнителя и посчитать его открытые заказы. */
function askOpenOrdersByWorker() {
  const ui = SpreadsheetApp.getUi();
  const answer = ui.prompt('Открытые заказы', 'Введите ID исполнителя, например D-004', ui.ButtonSet.OK_CANCEL);
  if (answer.getSelectedButton() !== ui.Button.OK) return; // нажали «Отмена»

  const workerId = answer.getResponseText().trim().toUpperCase();
  const values = SpreadsheetApp.getActive().getSheetByName('Заказы').getDataRange().getValues();

  // Колонка F (индекс 5) — ID исполнителя, E (индекс 4) — статус
  const open = values.slice(1).filter(function (row) {
    return row[5] === workerId && row[4] !== 'Готов';
  });

  if (open.length === 0) {
    ui.alert('У ' + workerId + ' нет открытых заказов (или такого ID нет).');
    return;
  }
  const ids = open.map(function (row) { return row[0] + ' — ' + row[4]; }).join('\n');
  ui.alert('Открытых заказов у ' + workerId + ': ' + open.length, ids, ui.ButtonSet.OK);
}

/** Подтверждение перед опасным действием: очистить архив. */
function clearArchiveWithConfirm() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActive().getSheetByName('Архив');
  const rows = sheet.getLastRow() - 1;

  if (rows < 1) {
    ui.alert('Архив и так пуст.');
    return;
  }
  const ok = ui.alert('Очистить архив?', 'Будет удалено строк: ' + rows + '. Отменить это можно только через историю версий.', ui.ButtonSet.YES_NO);
  if (ok !== ui.Button.YES) return;

  sheet.getRange(2, 1, rows, sheet.getLastColumn()).clearContent();
  SpreadsheetApp.getActive().toast('Архив очищен', 'Готово', 3);
}

/* ---------- Задание со *: боковая панель с формой ---------- */

/** Открыть боковую панель. Её разметка — в файле AddOrder.html этого же проекта. */
function showAddOrderSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('AddOrder').setTitle('Добавить заказ');
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Эту функцию вызывает форма из панели. Она получает объект с полями формы
 * и дописывает строку в конец листа «Заказы».
 */
function addOrderFromForm(form) {
  const sheet = SpreadsheetApp.getActive().getSheetByName('Заказы');
  const hours = { FP: 24, ESX: 48, FF: 72 }[form.type] || 24;
  const created = new Date();
  const deadline = new Date(created.getTime() + hours * 60 * 60 * 1000);
  const nextId = 'ORD-' + (10000 + sheet.getLastRow()); // простая нумерация для учебного примера

  sheet.appendRow([nextId, created, deadline, form.type, 'Новый', form.workerId.trim().toUpperCase(),
    form.workerName.trim(), form.supervisor.trim(), '', '', '', form.comment]);
  return nextId; // это значение получит форма и покажет пользователю
}
