/**
 * Модуль 8, задание со *. Сигнал о необычной неделе — эталонное решение.
 *
 * Сначала попробуйте написать сами, потом сравните.
 *
 * checkWeekAlerts сравнивает последнюю полную неделю с предыдущими
 * и присылает сообщение, только если какая-то метрика вышла за обычные рамки.
 * Обычные рамки — среднее за прошлые недели ± 2 стандартных отклонения (σ).
 *
 * Нужны файлы урока 8.1–8.2 (Stats, Summary) и модуля 5:
 * sendSlackMessage_, getNotifyEmail_ (Notify), removeTriggersFor_ (Triggers).
 */
const ALERT = {
  HISTORY_WEEKS: 6, // с какими неделями сравниваем: шесть предыдущих
  MIN_HISTORY: 4,   // меньше четырёх недель истории — сравнивать рано
  SIGMA: 2,         // насколько далеко от среднего — уже «необычно»
};

function checkWeekAlerts() {
  buildSummary(); // свежие метрики на листе «Summary» (урок 8.2)

  const ss = SpreadsheetApp.getActive();
  const tz = ss.getSpreadsheetTimeZone();
  // Блок по неделям — колонки A–E. Колонки берём по шапке: добавите метрику — она тоже проверится
  const values = ss.getSheetByName('Summary').getRange('A1:E').getValues()
    .filter(function (row) { return row[0] !== ''; });
  const header = values[0];
  const weeks = values.slice(1);

  if (weeks.length < ALERT.MIN_HISTORY + 1) {
    console.log('Not enough history: ' + weeks.length + ' full week(s), need ' + (ALERT.MIN_HISTORY + 1));
    return;
  }

  const last = weeks[weeks.length - 1];
  const history = weeks.slice(-1 - ALERT.HISTORY_WEEKS, -1); // до 6 недель перед последней

  const alerts = [];
  for (let col = 1; col < header.length; col++) {
    const value = last[col];
    const past = history
      .map(function (row) { return row[col]; })
      .filter(function (x) { return typeof x === 'number'; });
    if (typeof value !== 'number' || past.length < ALERT.MIN_HISTORY) continue;

    const m = mean_(past);
    const s = stdev_(past);
    const low = m - ALERT.SIGMA * s;
    const high = m + ALERT.SIGMA * s;
    if (value >= low && value <= high) continue; // обычная неделя для этой метрики

    alerts.push(header[col] + ': ' + formatMetric_(header[col], value) + (value > high ? ' ↑' : ' ↓') +
      ' (usual: ' + formatMetric_(header[col], Math.max(low, 0)) + '–' + formatMetric_(header[col], high) + ')');
  }

  const week = Utilities.formatDate(last[0], tz, 'dd.MM');
  if (alerts.length === 0) {
    console.log('Week of ' + week + ': all metrics are within the usual range');
    return; // ничего необычного — не беспокоим
  }

  const text = '📈 Unusual week of ' + week + '\n' + alerts.join('\n') + '\n' + ss.getUrl();
  if (PropertiesService.getScriptProperties().getProperty('SLACK_WEBHOOK')) {
    sendSlackMessage_(text);
  } else {
    MailApp.sendEmail(getNotifyEmail_(), 'Unusual week of ' + week, text);
  }
}

/** Доли — в процентах, остальное — с одним знаком после запятой. */
function formatMetric_(name, value) {
  return name.indexOf('Share') !== -1 ? Math.round(value * 100) + '%' : String(round1_(value));
}

/** Проверка каждый понедельник в 9 утра — после дашборда и отчёта. */
function createWeekAlertsTrigger() {
  removeTriggersFor_('checkWeekAlerts');
  ScriptApp.newTrigger('checkWeekAlerts')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(9)
    .create();
  SpreadsheetApp.getActive().toast('Unusual weeks are checked on Mondays at 9:00');
}
