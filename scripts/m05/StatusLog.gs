/**
 * Модуль 5, урок 5.2. Метка времени и автор правки.
 *
 * Когда кто-то меняет статус заказа (колонка E), скрипт записывает рядом,
 * когда это случилось и кто это сделал: колонки N «Статус изменён» и O «Кто изменил».
 *
 * Это ещё один обработчик для диспетчера onEdit из модуля 3.
 * В файле «Флажки» замените функцию onEdit на версию ниже — с двумя строками.
 */

const STATUS_COL = 5;        // E — Статус
const STATUS_TIME_COL = 14;  // N — Статус изменён
const STATUS_WHO_COL = 15;   // O — Кто изменил

/* Обновлённый диспетчер: по строке на каждый обработчик.
function onEdit(e) {
  handleCloseCheckbox_(e);   // модуль 3: флажок «Закрыть»
  handleStatusChange_(e);    // модуль 5: метка времени и автор
}
*/

function handleStatusChange_(e) {
  const range = e.range;
  const sheet = range.getSheet();
  if (sheet.getName() !== 'Заказы') return;
  if (range.getColumn() !== STATUS_COL || range.getRow() < 2) return;
  if (range.getNumRows() > 1) return; // вставили сразу много строк — не наш случай

  const row = range.getRow();
  // Email виден, если правку сделал сотрудник того же Google Workspace.
  // Для внешних пользователей Google его не показывает — тогда пишем «—».
  const who = Session.getActiveUser().getEmail() || '—';
  sheet.getRange(row, STATUS_TIME_COL, 1, 2).setValues([[new Date(), who]]);

  // Автостатус: если поставили «Готов», а дата сдачи пустая — ставим её
  if (e.value === 'Готов' && sheet.getRange(row, 11).isBlank()) {
    sheet.getRange(row, 11).setValue(new Date());
  }
}

/** Один раз: подписать новые колонки N и O. */
function addStatusLogHeaders() {
  const sheet = SpreadsheetApp.getActive().getSheetByName('Заказы');
  sheet.getRange(1, STATUS_TIME_COL, 1, 2).setValues([['Статус изменён', 'Кто изменил']]);
  sheet.getRange(2, STATUS_TIME_COL, sheet.getMaxRows() - 1, 1).setNumberFormat('dd.MM.yyyy HH:mm');
}
