/**
 * Модуль 8, урок 8.6. Отчёт по расписанию: дашборд в PDF и на почту.
 *
 * Нужны: Dashboard.gs (урок 8.5), getNotifyEmail_ из модуля 5 (файл «Уведомления»)
 * и removeTriggersFor_ из модуля 5 (файл «Триггеры»).
 */
function sendWeeklyReport() {
  const ss = SpreadsheetApp.getActive();
  buildDashboard();

  // Все изменения — в таблицу, и пара секунд, чтобы графики успели перерисоваться
  SpreadsheetApp.flush();
  Utilities.sleep(3000);

  const tz = ss.getSpreadsheetTimeZone();
  const today = Utilities.formatDate(new Date(), tz, 'dd.MM.yyyy');
  const pdf = exportSheetToPdf_('Дашборд').setName('Отчёт по заказам ' + today + '.pdf');

  // Короткая выжимка прямо в тексте письма: заказов и доля просрочек за неделю
  const kpi = ss.getSheetByName('Дашборд').getRange('A5:D8').getDisplayValues();
  const lines = kpi.map(function (row) {
    return '<li>' + row[0] + ': <b>' + row[1] + '</b> (неделей раньше ' + row[2] + ')</li>';
  }).join('');

  MailApp.sendEmail({
    to: getNotifyEmail_(),
    subject: 'Отчёт по заказам · ' + today,
    htmlBody: '<p>Итоги последней полной недели:</p><ul>' + lines + '</ul>' +
      '<p>Дашборд — во вложении. <a href="' + ss.getUrl() + '">Открыть таблицу</a></p>',
    attachments: [pdf],
  });
}

/**
 * Один лист таблицы → PDF-файл (blob), который можно приложить к письму
 * или сохранить на Диск: DriveApp.createFile(blob).
 */
function exportSheetToPdf_(sheetName) {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(sheetName);

  // Выгрузка идёт через Диск. Эта строка нужна, чтобы при первом запуске
  // Google запросил доступ к файлам на Диске — иначе адрес выгрузки может ответить ошибкой доступа.
  DriveApp.getFileById(ss.getId());

  const url = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export' +
    '?format=pdf' +
    '&gid=' + sheet.getSheetId() + // только этот лист
    '&size=A4&portrait=false' +    // A4, альбомная ориентация
    '&fitw=true' +                 // уместить по ширине страницы
    '&gridlines=false&sheetnames=false&printtitle=false&pagenumbers=false';

  const response = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
  });
  return response.getBlob();
}

/** Отчёт каждый понедельник в 8 утра — через час после обновления дашборда. */
function createWeeklyReportTrigger() {
  removeTriggersFor_('sendWeeklyReport');
  ScriptApp.newTrigger('sendWeeklyReport')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(8)
    .create();
  SpreadsheetApp.getActive().toast('Отчёт будет приходить по понедельникам в 8:00');
}
