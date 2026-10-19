/**
 * Модуль 5, урок 5.4. Листы и файлы по шаблону.
 *
 * Создаёт лист отчёта на текущий месяц, например «Отчёт 10.2026»,
 * по шаблону «Шаблон отчёта». Если шаблона нет — сначала создаёт его.
 * Повесьте функцию на триггер 1-го числа (onMonthDay(1)) — и новый месяц
 * будет начинаться сам.
 */
function createMonthReportSheet() {
  const ss = SpreadsheetApp.getActive();
  const name = 'Отчёт ' + Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'MM.yyyy');

  if (ss.getSheetByName(name)) {
    ss.toast('Лист «' + name + '» уже есть', 'Ничего не делаю', 5);
    return;
  }

  const template = ss.getSheetByName('Шаблон отчёта') || createReportTemplate_();
  const sheet = ss.insertSheet(name, { template: template });
  sheet.getRange('B1').setValue(name);
  ss.setActiveSheet(sheet);
}

/** Простой шаблон: заголовок, шапка таблицы, оформление. */
function createReportTemplate_() {
  const sheet = SpreadsheetApp.getActive().insertSheet('Шаблон отчёта');
  sheet.getRange('A1:B1').setValues([['Отчёт за', '']]).setFontWeight('bold').setFontSize(14);
  sheet.getRange('A3:E3')
    .setValues([['Исполнитель', 'Заказов', 'Часы', 'Средняя оценка QC', 'Просрочено']])
    .setBackground('#1f2a44').setFontColor('#ffffff').setFontWeight('bold');
  sheet.setFrozenRows(3);
  sheet.setColumnWidths(1, 5, 150);
  sheet.setTabColor('#e8e13a');
  return sheet;
}

/**
 * Задание со *: копия всего файла таблицы — например, личная таблица
 * для нового сотрудника. Нужен доступ к Drive: скрипт попросит разрешение.
 * Копия ложится в ту же папку, что и оригинал.
 */
function copySpreadsheetForEmployee() {
  const ui = SpreadsheetApp.getUi();
  const answer = ui.prompt('Копия для сотрудника', 'Имя сотрудника:', ui.ButtonSet.OK_CANCEL);
  if (answer.getSelectedButton() !== ui.Button.OK) return;

  const ss = SpreadsheetApp.getActive();
  const file = DriveApp.getFileById(ss.getId());
  const folder = file.getParents().hasNext() ? file.getParents().next() : DriveApp.getRootFolder();
  const copy = file.makeCopy(ss.getName() + ' — ' + answer.getResponseText().trim(), folder);
  ui.alert('Копия создана', copy.getUrl(), ui.ButtonSet.OK);
}
