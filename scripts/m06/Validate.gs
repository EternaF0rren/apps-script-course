/**
 * Модуль 6, урок 4. Проверка данных.
 *
 * validateOrders проходит по листу «Orders» и проверяет правила.
 * Ошибочные ячейки подсвечивает, а список ошибок выводит на лист «Errors»:
 * строка, колонка, значение, что не так.
 */

const VALID_STATUSES = ['New', 'In Progress', 'In Review', 'Done', 'Problem'];
const ERROR_COLOR = '#fff3cd';

function validateOrders() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName('Orders');
  const range = sheet.getRange(2, 1, sheet.getLastRow() - 1, 12);
  const values = range.getValues();
  const headers = sheet.getRange(1, 1, 1, 12).getValues()[0];

  const errors = [['Row', 'Column', 'Value', 'Issue']];
  const colors = values.map(function (row) { return row.map(function () { return null; }); });

  function flag(i, col, problem) {
    errors.push([i + 2, headers[col], values[i][col], problem]);
    colors[i][col] = ERROR_COLOR;
  }

  values.forEach(function (row, i) {
    const created = row[1], deadline = row[2], status = row[4], hours = row[8], qc = row[9], done = row[10];

    if (!(created instanceof Date)) flag(i, 1, 'Created is not a valid date');
    if (deadline instanceof Date && created instanceof Date && deadline <= created) flag(i, 2, 'Deadline is before Created');
    if (VALID_STATUSES.indexOf(status) === -1) flag(i, 4, 'Unknown status');
    if (typeof hours !== 'number' || hours <= 0 || hours > 24) flag(i, 8, 'Hours must be a number from 0 to 24');
    if (qc !== '' && (qc < 1 || qc > 5)) flag(i, 9, 'QC Score must be from 1 to 5');
    if (status === 'Done' && !(done instanceof Date)) flag(i, 10, 'Status is Done but Delivered is empty');
    if (status !== 'Done' && done instanceof Date) flag(i, 10, 'Delivered is filled but status is not Done');
  });

  range.setBackgrounds(colors);

  const out = ss.getSheetByName('Errors') || ss.insertSheet('Errors');
  out.clearContents();
  out.getRange(1, 1, errors.length, 4).setValues(errors);
  ss.toast('Errors found: ' + (errors.length - 1), 'Data validation', 5);
}

/** Выпадающий список статусов: чтобы ошибкам вообще неоткуда было взяться. */
function addStatusDropdown() {
  const sheet = SpreadsheetApp.getActive().getSheetByName('Orders');
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(VALID_STATUSES, true) // true — показывать стрелку списка
    .setAllowInvalid(false)                   // другие значения ввести нельзя
    .build();
  sheet.getRange(2, 5, sheet.getMaxRows() - 1, 1).setDataValidation(rule);
}
