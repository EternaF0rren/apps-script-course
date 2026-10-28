/**
 * Модуль 5, задание со *. Предупреждение о горящих заказах — эталонное решение.
 *
 * Сначала попробуйте написать сами, потом сравните.
 *
 * Раз в час warnHotOrders ищет незакрытые заказы, до дедлайна которых осталось
 * меньше HOT_HOURS часов (лист «Settings»), и присылает одно сообщение со списком:
 * в Slack, если вебхук настроен, иначе письмом. О каждом заказе предупреждает один раз.
 *
 * Нужны функции из файлов урока 5.5 и 5.1: sendSlackMessage_, getNotifyEmail_ (Notify),
 * removeTriggersFor_ (Triggers).
 */
const HOT_WARNED_KEY = 'HOT_WARNED_IDS'; // где в свойствах скрипта храним, о ком уже предупредили
const HOT_QUIET_FROM = 21;               // бонус: с 21:00 до 9:00 не беспокоим
const HOT_QUIET_TO = 9;

function warnHotOrders() {
  const ss = SpreadsheetApp.getActive();
  const tz = ss.getSpreadsheetTimeZone();
  const now = new Date();

  // Бонус «не беспокоить ночью». Ничего не помечаем — утром предупредим.
  const hour = Number(Utilities.formatDate(now, tz, 'H'));
  if (hour >= HOT_QUIET_FROM || hour < HOT_QUIET_TO) return;

  const limitMs = hotHours_() * 60 * 60 * 1000;
  const values = ss.getSheetByName('Orders').getDataRange().getValues().slice(1);

  // Горящие: не Done, дедлайн ещё не прошёл, но до него меньше limitMs
  const hot = values.filter(function (row) {
    const deadline = row[2];
    return deadline instanceof Date && row[4] !== 'Done' &&
      deadline > now && deadline - now <= limitMs;
  });

  // О ком уже предупреждали. Оставляем в памяти только тех, кто всё ещё горит:
  // так список не растёт бесконечно.
  const props = PropertiesService.getScriptProperties();
  const warned = JSON.parse(props.getProperty(HOT_WARNED_KEY) || '[]');
  const hotIds = hot.map(function (row) { return row[0]; });
  const fresh = hot.filter(function (row) { return warned.indexOf(row[0]) === -1; });

  props.setProperty(HOT_WARNED_KEY, JSON.stringify(
    warned.filter(function (id) { return hotIds.indexOf(id) !== -1; })
      .concat(fresh.map(function (row) { return row[0]; }))));

  if (fresh.length === 0) return; // новостей нет — сообщения нет

  const lines = fresh.map(function (row) {
    const hoursLeft = Math.round((row[2] - now) / 36e5 * 10) / 10;
    return row[0] + ' — ' + row[6] + ', ' + row[4] + ', ' + hoursLeft + ' h left (' +
      Utilities.formatDate(row[2], tz, 'dd.MM HH:mm') + ')';
  });
  const text = '🔥 Deadline soon: ' + fresh.length + ' order(s)\n' + lines.join('\n');

  if (PropertiesService.getScriptProperties().getProperty('SLACK_WEBHOOK')) {
    sendSlackMessage_(text);
  } else {
    MailApp.sendEmail(getNotifyEmail_(), 'Deadline soon: ' + fresh.length + ' order(s)',
      text + '\n\n' + ss.getUrl());
  }
}

/** Порог из строки HOT_HOURS листа «Settings». Пусто или не число — 4 часа. */
function hotHours_() {
  const rows = SpreadsheetApp.getActive().getSheetByName('Settings').getDataRange().getValues();
  const found = rows.filter(function (row) { return row[0] === 'HOT_HOURS'; })[0];
  const value = found ? Number(found[1]) : NaN;
  return value > 0 ? value : 4;
}

/** Запускать warnHotOrders каждый час. Повторный запуск не создаёт дублей. */
function createHotOrdersTrigger() {
  removeTriggersFor_('warnHotOrders');
  ScriptApp.newTrigger('warnHotOrders').timeBased().everyHours(1).create();
  SpreadsheetApp.getActive().toast('Hot orders are checked every hour', 'Trigger created', 5);
}

/** Для проверки: забыть, о ком уже предупреждали, — следующий запуск пришлёт всех заново. */
function resetHotWarnings() {
  PropertiesService.getScriptProperties().deleteProperty(HOT_WARNED_KEY);
}
