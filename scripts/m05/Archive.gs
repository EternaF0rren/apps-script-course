/**
 * Модуль 5, урок 3. Архив закрытых заказов.
 *
 * Переносит на лист «Archive» заказы со статусом «Done», сданные больше
 * ARCHIVE_AFTER_DAYS дней назад. Рабочий лист остаётся коротким и быстрым.
 * Запуск: вручную из меню или раз в неделю триггером (Triggers.gs).
 */

const ARCHIVE_AFTER_DAYS = 7;
const ORDER_COLUMNS = 12; // A–L: колонки заказа, которые переносим в архив

function archiveDoneOrders() {
  const ss = SpreadsheetApp.getActive();
  const orders = ss.getSheetByName('Orders');
  const archive = ss.getSheetByName('Archive');

  const values = orders.getRange(2, 1, orders.getLastRow() - 1, ORDER_COLUMNS).getValues();
  const border = new Date(Date.now() - ARCHIVE_AFTER_DAYS * 24 * 60 * 60 * 1000);
  const now = new Date();

  // Номера строк в таблице (с 1) и сами строки, которые уезжают в архив
  const toMove = [];
  values.forEach(function (row, i) {
    const status = row[4];
    const done = row[10];
    if (status === 'Done' && done instanceof Date && done < border) {
      toMove.push({ sheetRow: i + 2, data: row.concat([now]) }); // + «Archived At»
    }
  });

  if (toMove.length === 0) {
    console.log('Nothing to archive');
    return;
  }

  // 1. Дописываем в архив одним вызовом
  const archiveRows = toMove.map(function (item) { return item.data; });
  archive.getRange(archive.getLastRow() + 1, 1, archiveRows.length, archiveRows[0].length)
    .setValues(archiveRows);

  // 2. Удаляем из «Orders» снизу вверх, чтобы номера строк не съезжали
  toMove.slice().reverse().forEach(function (item) {
    orders.deleteRow(item.sheetRow);
  });

  console.log('Moved to archive: ' + toMove.length);
}
