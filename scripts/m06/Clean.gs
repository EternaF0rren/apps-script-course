/**
 * Модуль 6, урок 2. Чистка данных.
 *
 * В модуле 2 мы удаляли строки-близнецы целиком. В жизни чаще бывает иначе:
 * один и тот же заказ выгрузили дважды, но с разным статусом. Нужно оставить
 * одну строку на каждый ID — самую свежую (ту, что ниже в выгрузке).
 * Функции toNameCase, capitalizeFirst, parseDate, parseNumber — из модуля 2.
 */
function dedupeExportById() {
  const sheet = SpreadsheetApp.getActive().getSheetByName('Выгрузка');
  const values = sheet.getDataRange().getValues();
  const header = values[0];

  // Объект-словарь: ключ — ID заказа, значение — последняя встреченная строка
  const latest = {};
  values.slice(1).forEach(function (row) {
    const id = String(row[0]).trim().toUpperCase();
    if (!id) return; // пустые строки пропускаем
    latest[id] = [id, toNameCase(row[1]), capitalizeFirst(row[2]), parseDate(row[3]), parseNumber(row[4])];
  });

  const rows = Object.keys(latest).sort().map(function (id) { return latest[id]; });

  sheet.getRange(2, 1, sheet.getMaxRows() - 1, header.length).clearContent();
  sheet.getRange(2, 1, rows.length, header.length).setValues(rows);
  SpreadsheetApp.getActive().toast('Было ' + (values.length - 1) + ' строк, осталось ' + rows.length + ' заказов', 'Готово', 5);
}
