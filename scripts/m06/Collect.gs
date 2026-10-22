/**
 * Модуль 6, урок 3. Сбор из нескольких таблиц и сверка списков.
 */

/**
 * Собрать данные из таблиц команд на один лист «Collected».
 * Список источников — на листе «Sources»: A — название, B — ссылка на таблицу,
 * C — имя листа в ней. Если листа «Sources» нет, функция создаст образец.
 * У запускающего должен быть доступ на просмотр ко всем таблицам из списка.
 */
function collectFromSources() {
  const ss = SpreadsheetApp.getActive();
  let sources = ss.getSheetByName('Sources');
  if (!sources) {
    sources = ss.insertSheet('Sources');
    sources.getRange('A1:C2').setValues([
      ['Name', 'Spreadsheet URL', 'Sheet'],
      ['My sandbox', ss.getUrl(), 'Orders'],
    ]);
    ss.toast('Sources sheet created. Add links and run again', 'Sources needed', 8);
    return;
  }

  const list = sources.getDataRange().getValues().slice(1).filter(function (r) { return r[1]; });
  let header = null;
  const all = [];
  const errors = [];

  list.forEach(function (src) {
    try {
      const sheet = SpreadsheetApp.openByUrl(src[1]).getSheetByName(src[2]);
      if (!sheet) throw new Error('no sheet "' + src[2] + '"');
      const values = sheet.getDataRange().getValues();
      if (!header) header = ['Source'].concat(values[0]);
      values.slice(1).forEach(function (row) { all.push([src[0]].concat(row)); });
    } catch (err) {
      errors.push(src[0] + ': ' + err.message); // одна сломанная ссылка не роняет весь сбор
    }
  });

  if (!header) {
    ss.toast('No source could be read', 'Error', 8);
    return;
  }

  // Строки разной длины (в таблицах разное число колонок) выравниваем по шапке
  const width = header.length;
  const rows = all.map(function (r) { return r.concat(new Array(width).fill('')).slice(0, width); });

  const out = ss.getSheetByName('Collected') || ss.insertSheet('Collected');
  out.clearContents();
  out.getRange(1, 1, rows.length + 1, width).setValues([header].concat(rows));

  const msg = 'Rows collected: ' + rows.length + (errors.length ? '. Errors: ' + errors.join('; ') : '');
  ss.toast(msg, 'Data collection', 10);
}

/**
 * Сверка двух списков: кто есть в «Сотрудниках», но не логировал часы
 * на последней неделе, и наоборот — чьих ID нет в справочнике.
 */
function compareEmployeesAndHours() {
  const ss = SpreadsheetApp.getActive();
  const employees = ss.getSheetByName('Employees').getDataRange().getValues().slice(1);
  const hours = ss.getSheetByName('Hours').getDataRange().getValues().slice(1);

  // Последняя неделя в листе «Hours»
  const lastWeek = Math.max.apply(null, hours.map(function (r) { return r[0].getTime(); }));
  const loggedIds = new Set(hours.filter(function (r) { return r[0].getTime() === lastWeek; })
    .map(function (r) { return r[1]; }));
  const knownIds = new Set(employees.map(function (r) { return r[0]; }));

  const result = [['ID', 'Name', 'Issue']];
  employees.forEach(function (r) {
    if (!loggedIds.has(r[0])) result.push([r[0], r[1], 'No hours logged last week']);
  });
  loggedIds.forEach(function (id) {
    if (!knownIds.has(id)) result.push([id, '', 'Not in the employee list']);
  });

  const out = ss.getSheetByName('Reconciliation') || ss.insertSheet('Reconciliation');
  out.clearContents();
  out.getRange(1, 1, result.length, 3).setValues(result);
  ss.toast('Mismatches found: ' + (result.length - 1), 'Reconciliation', 5);
}
