/**
 * Модуль 6, урок 6.4. Проверка данных.
 *
 * validateOrders проходит по листу «Заказы» и проверяет правила.
 * Ошибочные ячейки подсвечивает, а список ошибок выводит на лист «Ошибки»:
 * строка, колонка, значение, что не так.
 */

const VALID_STATUSES = ['Новый', 'В работе', 'На проверке', 'Готов', 'Проблема'];
const ERROR_COLOR = '#fff3cd';

function validateOrders() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName('Заказы');
  const range = sheet.getRange(2, 1, sheet.getLastRow() - 1, 12);
  const values = range.getValues();
  const headers = sheet.getRange(1, 1, 1, 12).getValues()[0];

  const errors = [['Строка', 'Колонка', 'Значение', 'Что не так']];
  const colors = values.map(function (row) { return row.map(function () { return null; }); });

  function flag(i, col, problem) {
    errors.push([i + 2, headers[col], values[i][col], problem]);
    colors[i][col] = ERROR_COLOR;
  }

  values.forEach(function (row, i) {
    const created = row[1], deadline = row[2], status = row[4], hours = row[8], qc = row[9], done = row[10];

    if (!(created instanceof Date)) flag(i, 1, 'Дата создания не распознана как дата');
    if (deadline instanceof Date && created instanceof Date && deadline <= created) flag(i, 2, 'Дедлайн раньше создания');
    if (VALID_STATUSES.indexOf(status) === -1) flag(i, 4, 'Неизвестный статус');
    if (typeof hours !== 'number' || hours <= 0 || hours > 24) flag(i, 8, 'Часы должны быть числом от 0 до 24');
    if (qc !== '' && (qc < 1 || qc > 5)) flag(i, 9, 'Оценка QC — от 1 до 5');
    if (status === 'Готов' && !(done instanceof Date)) flag(i, 10, 'Заказ готов, но нет даты сдачи');
    if (status !== 'Готов' && done instanceof Date) flag(i, 10, 'Есть дата сдачи, но статус не «Готов»');
  });

  range.setBackgrounds(colors);

  const out = ss.getSheetByName('Ошибки') || ss.insertSheet('Ошибки');
  out.clearContents();
  out.getRange(1, 1, errors.length, 4).setValues(errors);
  ss.toast('Найдено ошибок: ' + (errors.length - 1), 'Проверка данных', 5);
}

/** Выпадающий список статусов: чтобы ошибкам вообще неоткуда было взяться. */
function addStatusDropdown() {
  const sheet = SpreadsheetApp.getActive().getSheetByName('Заказы');
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(VALID_STATUSES, true) // true — показывать стрелку списка
    .setAllowInvalid(false)                   // другие значения ввести нельзя
    .build();
  sheet.getRange(2, 5, sheet.getMaxRows() - 1, 1).setDataValidation(rule);
}
