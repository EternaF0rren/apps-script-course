/**
 * Модуль 5, урок 1. Триггеры.
 *
 * Простые триггеры — функции с особыми именами onOpen и onEdit (модуль 3).
 * Устанавливаемые триггеры создаются отдельно: в разделе «Триггеры» редактора
 * или кодом ниже. Им можно больше: работать по расписанию, отправлять письма,
 * ходить в Slack. Выполняются они с правами того, кто их создал.
 */

/** Каждое утро в 9:00–10:00 присылать сводку просрочек (функция из Notify.gs). */
function createMorningDigestTrigger() {
  removeTriggersFor_('sendOverdueDigest'); // не плодим дубли при повторном запуске
  ScriptApp.newTrigger('sendOverdueDigest')
    .timeBased()
    .everyDays(1)
    .atHour(9)
    .create();
  SpreadsheetApp.getActive().toast('Сводка будет приходить каждый день около 9:00', 'Триггер создан', 5);
}

/** Раз в неделю, в понедельник утром, переносить закрытые заказы в архив (Archive.gs). */
function createWeeklyArchiveTrigger() {
  removeTriggersFor_('archiveDoneOrders');
  ScriptApp.newTrigger('archiveDoneOrders')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(7)
    .create();
  SpreadsheetApp.getActive().toast('Архив будет обновляться по понедельникам около 7:00', 'Триггер создан', 5);
}

/**
 * Устанавливаемый триггер «при правке». Нужен там, где простой onEdit не справится:
 * например, отправить сообщение в Slack, когда статус стал «Проблема» (Notify.gs).
 */
function createProblemAlertTrigger() {
  removeTriggersFor_('onProblemStatus');
  ScriptApp.newTrigger('onProblemStatus')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onEdit()
    .create();
  SpreadsheetApp.getActive().toast('Статус «Проблема» будет уходить в Slack', 'Триггер создан', 5);
}

/** Показать в журнале все триггеры проекта: какая функция и по какому событию. */
function listTriggers() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    console.log(t.getHandlerFunction() + ' — ' + t.getEventType());
  });
}

/** Удалить все триггеры этого проекта, созданные вами. Полезно перед отпуском. */
function deleteAllMyTriggers() {
  ScriptApp.getProjectTriggers().forEach(function (t) { ScriptApp.deleteTrigger(t); });
  SpreadsheetApp.getActive().toast('Все триггеры удалены', 'Готово', 3);
}

/** Служебная: удалить триггеры, которые запускают функцию с этим именем. */
function removeTriggersFor_(functionName) {
  ScriptApp.getProjectTriggers()
    .filter(function (t) { return t.getHandlerFunction() === functionName; })
    .forEach(function (t) { ScriptApp.deleteTrigger(t); });
}
