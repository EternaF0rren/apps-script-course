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

/** Пункт меню «Add order form»: открыть боковую панель. Разметка — в файле AddOrder.html. */
function showAddOrderSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('AddOrder').setTitle('Add order');
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Её вызывает форма из панели через google.script.run.
 * Получает поля формы, дописывает заказ в конец «Orders» и возвращает его номер.
 */
function addOrderFromForm(form) {
  const ss = SpreadsheetApp.getActive();
  const orders = ss.getSheetByName('Orders');

  // Имя и супервайзера берём из справочника, а не просим набирать руками
  const workerId = form.workerId.trim().toUpperCase();
  const employee = ss.getSheetByName('Employees').getDataRange().getValues()
    .filter(function (row) { return row[0] === workerId; })[0];
  if (!employee) throw new Error('No assignee with ID ' + workerId); // этот текст покажет панель

  // Новый номер — на 1 больше самого большого на «Orders» и в «Archive»
  let maxNumber = 10000;
  ['Orders', 'Archive'].forEach(function (name) {
    ss.getSheetByName(name).getRange('A2:A').getValues().forEach(function (row) {
      const n = Number(String(row[0]).replace(/\D/g, ''));
      if (n > maxNumber) maxNumber = n;
    });
  });
  const id = 'ORD-' + (maxNumber + 1);

  const DEADLINE_HOURS = { FP: 24, ESX: 48, FF: 72 }; // срок по типу заказа
  const created = new Date();
  const deadline = new Date(created.getTime() + DEADLINE_HOURS[form.type] * 60 * 60 * 1000);

  orders.appendRow([id, created, deadline, form.type, 'New', workerId,
    employee[1], employee[5], '', '', '', form.comment]);
  return id; // это значение получит панель
}
