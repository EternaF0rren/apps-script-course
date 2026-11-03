/**
 * Модуль 6, задание со * 2. Что изменилось с прошлой копии — эталонное решение.
 *
 * Сначала попробуйте написать сами, потом сравните.
 *
 * showChangesSinceBackup сравнивает лист «Orders» с последней страховочной копией
 * «Orders · copy …» (backupSheet_ из модуля 4) и пишет на лист «Changes»:
 * какие заказы появились, какие пропали и что поменялось в отслеживаемых колонках.
 */
const CHANGES_SHEET = 'Changes';
const WATCH_COLUMNS = ['Status', 'Assignee', 'Deadline', 'Hours']; // какие колонки сравниваем — по заголовкам

function showChangesSinceBackup() {
  const ss = SpreadsheetApp.getActive();
  const tz = ss.getSpreadsheetTimeZone();

  // Последняя копия — самая правая вкладка с таким началом имени
  const copies = ss.getSheets().filter(function (s) {
    return s.getName().indexOf('Orders · copy ') === 0;
  });
  if (copies.length === 0) {
    ss.toast('No backup copy found. Run backupSheet_(\'Orders\') first.', 'Changes', 8);
    return;
  }
  const base = copies[copies.length - 1];

  const before = byId_(base.getDataRange().getValues());
  const after = byId_(ss.getSheetByName('Orders').getDataRange().getValues());

  const changes = [];
  // Новые заказы и изменения в старых
  Object.keys(after.rows).forEach(function (id) {
    const now = after.rows[id];
    const was = before.rows[id];
    if (!was) {
      changes.push([id, 'New order', '', '']);
      return;
    }
    WATCH_COLUMNS.forEach(function (col) {
      const a = was[before.col[col]], b = now[after.col[col]];
      if (!sameValue_(a, b)) changes.push([id, col, show_(a, tz), show_(b, tz)]);
    });
  });
  // Заказы, которых больше нет на листе: удалены или ушли в архив
  Object.keys(before.rows).forEach(function (id) {
    if (!after.rows[id]) changes.push([id, 'Removed', '', '']);
  });
  changes.sort(function (x, y) { return x[0] < y[0] ? -1 : x[0] > y[0] ? 1 : 0; });

  const out = ss.getSheetByName(CHANGES_SHEET) || ss.insertSheet(CHANGES_SHEET);
  out.clear();
  out.getRange(1, 1).setValue('Compared with: ' + base.getName()).setFontStyle('italic');
  out.getRange(2, 1, 1, 4).setValues([['Order ID', 'Change', 'Was', 'Now']]).setFontWeight('bold');
  if (changes.length > 0) {
    out.getRange(3, 1, changes.length, 4).setValues(changes);
  } else {
    out.getRange(3, 1).setValue('No changes');
  }
  out.autoResizeColumns(1, 4);
  ss.toast('Changes: ' + changes.length + ' (compared with ' + base.getName() + ')', 'Changes', 5);
}

/**
 * Превращает данные листа в словарь по ID заказа и запоминает,
 * в какой колонке какой заголовок. Колонки ищем по названию, а не по номеру:
 * если в копии колонок меньше или их порядок другой, сравнение не сломается.
 */
function byId_(values) {
  const header = values[0];
  const col = {};
  WATCH_COLUMNS.forEach(function (name) {
    const i = header.indexOf(name);
    if (i === -1) throw new Error('Column "' + name + '" not found');
    col[name] = i;
  });
  const rows = {};
  values.slice(1).forEach(function (row) {
    const id = String(row[0]).trim();
    if (id) rows[id] = row;
  });
  return { col: col, rows: rows };
}

/** Даты сравниваем по времени, остальное — как текст: 4 и «4» считаем одинаковыми. */
function sameValue_(a, b) {
  if (a instanceof Date && b instanceof Date) return a.getTime() === b.getTime();
  return String(a) === String(b);
}

/** Как показать значение в отчёте. */
function show_(value, tz) {
  return value instanceof Date ? Utilities.formatDate(value, tz, 'dd.MM.yyyy HH:mm') : value;
}
