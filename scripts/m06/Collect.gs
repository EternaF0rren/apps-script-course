/**
 * Модуль 6, урок 6.3. Сбор из нескольких таблиц и сверка списков.
 */

/**
 * Собрать данные из таблиц команд на один лист «Сбор».
 * Список источников — на листе «Источники»: A — название, B — ссылка на таблицу,
 * C — имя листа в ней. Если листа «Источники» нет, функция создаст образец.
 * У запускающего должен быть доступ на просмотр ко всем таблицам из списка.
 */
function collectFromSources() {
  const ss = SpreadsheetApp.getActive();
  let sources = ss.getSheetByName('Источники');
  if (!sources) {
    sources = ss.insertSheet('Источники');
    sources.getRange('A1:C2').setValues([
      ['Название', 'Ссылка на таблицу', 'Лист'],
      ['Моя песочница', ss.getUrl(), 'Заказы'],
    ]);
    ss.toast('Создан лист «Источники». Добавьте ссылки и запустите снова', 'Нужны источники', 8);
    return;
  }

  const list = sources.getDataRange().getValues().slice(1).filter(function (r) { return r[1]; });
  let header = null;
  const all = [];
  const errors = [];

  list.forEach(function (src) {
    try {
      const sheet = SpreadsheetApp.openByUrl(src[1]).getSheetByName(src[2]);
      if (!sheet) throw new Error('нет листа «' + src[2] + '»');
      const values = sheet.getDataRange().getValues();
      if (!header) header = ['Источник'].concat(values[0]);
      values.slice(1).forEach(function (row) { all.push([src[0]].concat(row)); });
    } catch (err) {
      errors.push(src[0] + ': ' + err.message); // одна сломанная ссылка не роняет весь сбор
    }
  });

  if (!header) {
    ss.toast('Ни один источник не прочитался', 'Ошибка', 8);
    return;
  }

  // Строки разной длины (в таблицах разное число колонок) выравниваем по шапке
  const width = header.length;
  const rows = all.map(function (r) { return r.concat(new Array(width).fill('')).slice(0, width); });

  const out = ss.getSheetByName('Сбор') || ss.insertSheet('Сбор');
  out.clearContents();
  out.getRange(1, 1, rows.length + 1, width).setValues([header].concat(rows));

  const msg = 'Собрано строк: ' + rows.length + (errors.length ? '. Ошибки: ' + errors.join('; ') : '');
  ss.toast(msg, 'Сбор данных', 10);
}

/**
 * Сверка двух списков: кто есть в «Сотрудниках», но не логировал часы
 * на последней неделе, и наоборот — чьих ID нет в справочнике.
 */
function compareEmployeesAndHours() {
  const ss = SpreadsheetApp.getActive();
  const employees = ss.getSheetByName('Сотрудники').getDataRange().getValues().slice(1);
  const hours = ss.getSheetByName('Часы').getDataRange().getValues().slice(1);

  // Последняя неделя в листе «Часы»
  const lastWeek = Math.max.apply(null, hours.map(function (r) { return r[0].getTime(); }));
  const loggedIds = new Set(hours.filter(function (r) { return r[0].getTime() === lastWeek; })
    .map(function (r) { return r[1]; }));
  const knownIds = new Set(employees.map(function (r) { return r[0]; }));

  const result = [['ID', 'Имя', 'Проблема']];
  employees.forEach(function (r) {
    if (!loggedIds.has(r[0])) result.push([r[0], r[1], 'Нет часов за последнюю неделю']);
  });
  loggedIds.forEach(function (id) {
    if (!knownIds.has(id)) result.push([id, '', 'Нет в справочнике сотрудников']);
  });

  const out = ss.getSheetByName('Сверка') || ss.insertSheet('Сверка');
  out.clearContents();
  out.getRange(1, 1, result.length, 3).setValues(result);
  ss.toast('Найдено расхождений: ' + (result.length - 1), 'Сверка', 5);
}
