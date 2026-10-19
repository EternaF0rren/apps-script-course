/**
 * Модуль 5, урок 5.5. Уведомления: письмо и Slack.
 */

/**
 * Сводка просрочек на почту. Адрес берётся с листа «Настройки» (EMAIL_УВЕДОМЛЕНИЙ),
 * а если там пусто — письмо уходит вам. Нет просрочек — письма нет.
 */
function sendOverdueDigest() {
  const ss = SpreadsheetApp.getActive();
  const values = ss.getSheetByName('Заказы').getDataRange().getValues();
  const now = new Date();

  const overdue = values.slice(1).filter(function (row) {
    return row[2] instanceof Date && row[2] < now && row[4] !== 'Готов';
  });
  if (overdue.length === 0) return;

  const tz = ss.getSpreadsheetTimeZone();
  const rowsHtml = overdue.map(function (row) {
    return '<tr><td>' + row[0] + '</td><td>' + row[6] + '</td><td>' + row[4] + '</td><td>' +
      Utilities.formatDate(row[2], tz, 'dd.MM HH:mm') + '</td></tr>';
  }).join('');

  const html = '<p>Просроченных заказов: <b>' + overdue.length + '</b></p>' +
    '<table border="1" cellpadding="4" style="border-collapse:collapse">' +
    '<tr><th>ID</th><th>Исполнитель</th><th>Статус</th><th>Дедлайн</th></tr>' + rowsHtml + '</table>' +
    '<p><a href="' + ss.getUrl() + '">Открыть таблицу</a></p>';

  MailApp.sendEmail({
    to: getNotifyEmail_(),
    subject: 'Просрочки: ' + overdue.length + ' заказов',
    htmlBody: html,
  });
}

/** Адрес для уведомлений: с листа «Настройки» или ваш собственный. */
function getNotifyEmail_() {
  const rows = SpreadsheetApp.getActive().getSheetByName('Настройки').getDataRange().getValues();
  const found = rows.filter(function (row) { return row[0] === 'EMAIL_УВЕДОМЛЕНИЙ'; })[0];
  const email = found ? String(found[1]).trim() : '';
  return email || Session.getEffectiveUser().getEmail();
}

/* ---------- Slack ---------- */

/**
 * Один раз: сохранить адрес входящего вебхука Slack в свойствах скрипта.
 * Не в коде и не на листе: свойства не видны тем, кто просто открывает таблицу.
 * Адрес вебхука создаёт администратор Slack или владелец канала.
 */
function saveSlackWebhook() {
  const ui = SpreadsheetApp.getUi();
  const answer = ui.prompt('Slack', 'Вставьте адрес входящего вебхука (https://hooks.slack.com/...)', ui.ButtonSet.OK_CANCEL);
  if (answer.getSelectedButton() !== ui.Button.OK) return;
  PropertiesService.getScriptProperties().setProperty('SLACK_WEBHOOK', answer.getResponseText().trim());
  ui.alert('Адрес сохранён в свойствах скрипта.');
}

/** Отправить текст в канал Slack. */
function sendSlackMessage_(text) {
  const url = PropertiesService.getScriptProperties().getProperty('SLACK_WEBHOOK');
  if (!url) {
    console.log('Вебхук Slack не настроен — сообщение не отправлено: ' + text);
    return;
  }
  UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({ text: text }),
  });
}

/**
 * Запускается устанавливаемым триггером «при правке» (Triggers.gs).
 * Простой onEdit так не умеет: ему запрещено ходить во внешние сервисы.
 */
function onProblemStatus(e) {
  const range = e.range;
  if (range.getSheet().getName() !== 'Заказы') return;
  if (range.getColumn() !== 5 || range.getRow() < 2 || range.getNumRows() > 1) return;
  if (e.value !== 'Проблема') return;

  const row = range.getSheet().getRange(range.getRow(), 1, 1, 8).getValues()[0];
  sendSlackMessage_('⚠️ Заказ ' + row[0] + ' в статусе «Проблема». Исполнитель: ' + row[6] + ', супервайзер: ' + row[7]);
}

/** Сколько писем ещё можно отправить сегодня. */
function checkMailQuota() {
  console.log('Осталось писем на сегодня: ' + MailApp.getRemainingDailyQuota());
}
