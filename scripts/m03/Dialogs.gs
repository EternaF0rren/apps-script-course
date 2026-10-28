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
  const answer = ui.prompt('Open orders', 'Enter an assignee ID, for example D-004', ui.ButtonSet.OK_CANCEL);
  if (answer.getSelectedButton() !== ui.Button.OK) return; // нажали «Отмена»

  const workerId = answer.getResponseText().trim().toUpperCase();
  const values = SpreadsheetApp.getActive().getSheetByName('Orders').getDataRange().getValues();

  // Колонка F (индекс 5) — ID исполнителя, E (индекс 4) — статус
  const open = values.slice(1).filter(function (row) {
    return row[5] === workerId && row[4] !== 'Done';
  });

  if (open.length === 0) {
    ui.alert(workerId + ' has no open orders (or there is no such ID).');
    return;
  }
  const ids = open.map(function (row) { return row[0] + ' — ' + row[4]; }).join('\n');
  ui.alert('Open orders for ' + workerId + ': ' + open.length, ids, ui.ButtonSet.OK);
}

/** Подтверждение перед опасным действием: очистить архив. */
function clearArchiveWithConfirm() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActive().getSheetByName('Archive');
  const rows = sheet.getLastRow() - 1;

  if (rows < 1) {
    ui.alert('The archive is already empty.');
    return;
  }
  const ok = ui.alert('Clear the archive?', 'Rows to delete: ' + rows + '. You can only undo this through version history.', ui.ButtonSet.YES_NO);
  if (ok !== ui.Button.YES) return;

  sheet.getRange(2, 1, rows, sheet.getLastColumn()).clearContent();
  SpreadsheetApp.getActive().toast('Archive cleared', 'Done', 3);
}

/* ---------- Дополнительно: боковая панель с формой ---------- */

/** Открыть боковую панель. Её разметка — в файле AddOrder.html этого же проекта. */
function showAddOrderSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('AddOrder').setTitle('Add order');
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Эту функцию вызывает форма из панели. Она получает объект с полями формы
 * и дописывает строку в конец листа «Orders».
 */
function addOrderFromForm(form) {
  const sheet = SpreadsheetApp.getActive().getSheetByName('Orders');
  const hours = { FP: 24, ESX: 48, FF: 72 }[form.type] || 24;
  const created = new Date();
  const deadline = new Date(created.getTime() + hours * 60 * 60 * 1000);
  const nextId = 'ORD-' + (10000 + sheet.getLastRow()); // простая нумерация для учебного примера

  sheet.appendRow([nextId, created, deadline, form.type, 'New', form.workerId.trim().toUpperCase(),
    form.workerName.trim(), form.supervisor.trim(), '', '', '', form.comment]);
  return nextId; // это значение получит форма и покажет пользователю
}
