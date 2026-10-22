/**
 * Модуль 3, уроки 2–3. Кнопка и горячая клавиша.
 *
 * Функция закрывает заказ в строке, где сейчас стоит курсор:
 * статус → «Done», в «Delivered» — текущие дата и время.
 *
 * Как повесить на кнопку: Вставка → Рисунок (или изображение поверх ячеек),
 * затем три точки на рисунке → «Назначить скрипт» → markActiveOrderDone.
 *
 * Как повесить на клавиши: Расширения → Макросы → Импортировать макрос →
 * markActiveOrderDone, затем Расширения → Макросы → Управление макросами →
 * задать сочетание Ctrl+Alt+Shift+цифра.
 */
function markActiveOrderDone() {
  const sheet = SpreadsheetApp.getActiveSheet();
  if (sheet.getName() !== 'Orders') {
    SpreadsheetApp.getActive().toast('Open the Orders sheet and select an order row', 'Wrong sheet', 5);
    return;
  }

  const row = sheet.getActiveRange().getRow();
  if (row < 2) {
    SpreadsheetApp.getActive().toast('Select an order row, not the header', 'Nothing to close', 5);
    return;
  }

  const STATUS_COL = 5; // E — Status
  const DONE_COL = 11;  // K — Delivered
  const id = sheet.getRange(row, 1).getValue();

  sheet.getRange(row, STATUS_COL).setValue('Done');
  sheet.getRange(row, DONE_COL).setValue(new Date());
  SpreadsheetApp.getActive().toast('Order ' + id + ' closed', 'Done', 3);
}
