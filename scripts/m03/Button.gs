/**
 * Модуль 3, уроки 3.2–3.3. Кнопка и горячая клавиша.
 *
 * Функция закрывает заказ в строке, где сейчас стоит курсор:
 * статус → «Готов», в «Сдан» — текущие дата и время.
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
  if (sheet.getName() !== 'Заказы') {
    SpreadsheetApp.getActive().toast('Откройте лист «Заказы» и встаньте на строку заказа', 'Не тот лист', 5);
    return;
  }

  const row = sheet.getActiveRange().getRow();
  if (row < 2) {
    SpreadsheetApp.getActive().toast('Встаньте на строку заказа, а не на шапку', 'Нечего закрывать', 5);
    return;
  }

  const STATUS_COL = 5; // E — Статус
  const DONE_COL = 11;  // K — Сдан
  const id = sheet.getRange(row, 1).getValue();

  sheet.getRange(row, STATUS_COL).setValue('Готов');
  sheet.getRange(row, DONE_COL).setValue(new Date());
  SpreadsheetApp.getActive().toast('Заказ ' + id + ' закрыт', 'Готово', 3);
}
