/**
 * Модуль 3, задание со *. Найти заказ по ID — эталонное решение.
 *
 * Сначала попробуйте написать сами, потом сравните.
 *
 * findOrderById спрашивает номер заказа, переходит к его строке
 * и показывает короткую карточку. Ищет на листе «Orders», затем в «Archive».
 * Запускать из меню «Course» или по горячим клавишам (урок 3.3):
 * окна работают только при запуске из таблицы.
 */
const FIND_SHEETS = ['Orders', 'Archive']; // где искать и в каком порядке

function findOrderById() {
  const ui = SpreadsheetApp.getUi();
  const answer = ui.prompt('Find order', 'Order ID, for example ORD-10145 or just 10145', ui.ButtonSet.OK_CANCEL);
  if (answer.getSelectedButton() !== ui.Button.OK) return; // нажали «Отмена» или закрыли окно

  // «ord-10145», « 10145 », «ORD 10145» → «ORD-10145»
  const digits = answer.getResponseText().replace(/\D/g, '');
  if (!digits) {
    ui.alert('Enter an order number, for example ORD-10145.');
    return;
  }
  const id = 'ORD-' + digits;

  const ss = SpreadsheetApp.getActive();
  for (let s = 0; s < FIND_SHEETS.length; s++) {
    const sheet = ss.getSheetByName(FIND_SHEETS[s]);
    if (!sheet) continue; // листа нет — ищем дальше, а не падаем

    const values = sheet.getDataRange().getValues();
    const index = values.findIndex(function (row) {
      return String(row[0]).trim().toUpperCase() === id;
    });
    if (index < 1) continue; // -1 — не нашли; 0 — это шапка

    // Переходим к строке: открываем лист и выделяем строку заказа целиком
    sheet.activate();
    sheet.getRange(index + 1, 1, 1, sheet.getLastColumn()).activate();

    ui.alert(id, orderCard_(values[index], sheet.getName(), ss.getSpreadsheetTimeZone()), ui.ButtonSet.OK);
    return;
  }

  ui.alert(id + ' not found on sheets: ' + FIND_SHEETS.join(', ') + '.');
}

/** Текст карточки: статус, исполнитель, дедлайн и пометка о просрочке. */
function orderCard_(row, sheetName, tz) {
  const status = row[4];
  const deadline = row[2];
  const lines = [
    'Status: ' + status,
    'Assignee: ' + row[6] + ' (' + row[5] + ')',
    'Deadline: ' + (deadline instanceof Date ? Utilities.formatDate(deadline, tz, 'dd.MM.yyyy HH:mm') : deadline),
  ];
  if (deadline instanceof Date && deadline < new Date() && status !== 'Done') {
    lines.push('⚠ Overdue');
  }
  if (sheetName !== 'Orders') {
    lines.push('Found on the ' + sheetName + ' sheet');
  }
  return lines.join('\n');
}
