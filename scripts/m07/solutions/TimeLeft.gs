/**
 * Модуль 7, задание со * 1. Сколько осталось до дедлайна — эталонное решение.
 *
 * Сначала попробуйте написать сами, потом сравните.
 *
 *   =TIME_LEFT(C2; E2)                       → «1 d 5 h»
 *   =TIME_LEFT(C2:C201; E2:E201; Settings!B8) → вся колонка, обновляется каждый час
 *
 * Третий аргумент — время из строки NOW листа «Settings» (урок 7.3, функция stampNow).
 * Своя формула не пересчитывается сама по себе, а с этим аргументом пересчитывается
 * каждый раз, когда триггер обновляет время.
 */

/**
 * Сколько осталось до дедлайна, текстом: «1 d 5 h», «3 h 20 min», «Overdue: 2 h».
 *
 * @param {Date} deadline Дедлайн (колонка C) или диапазон дедлайнов.
 * @param {string} status Статус (колонка E) или диапазон того же размера.
 * @param {Date} now Необязательно: от какого времени считать, например Settings!B8. По умолчанию — текущее время.
 * @return Текст с оставшимся временем, «Done» для закрытых заказов.
 * @customfunction
 */
function TIME_LEFT(deadline, status, now) {
  // Два диапазона одинакового размера: идём по строкам, i-я строка из каждого
  if (Array.isArray(deadline)) {
    return deadline.map(function (row, i) {
      return [TIME_LEFT(row[0], status[i][0], now)];
    });
  }
  if (status === 'Done') return 'Done';
  if (!(deadline instanceof Date)) return '';

  const from = now instanceof Date ? now : new Date();
  const ms = deadline - from;
  if (ms < 0) return 'Overdue: ' + durationText_(-ms);
  return durationText_(ms);
}

/** Миллисекунды → «2 d 5 h», «3 h 20 min» или «40 min». Минуты показываем только для срока меньше суток. */
function durationText_(ms) {
  const totalMinutes = Math.floor(ms / 60000);
  const days = Math.floor(totalMinutes / (24 * 60));
  const hours = Math.floor(totalMinutes / 60) % 24;
  const minutes = totalMinutes % 60;

  if (days > 0) return hours > 0 ? days + ' d ' + hours + ' h' : days + ' d';
  if (hours > 0) return minutes > 0 ? hours + ' h ' + minutes + ' min' : hours + ' h';
  return minutes + ' min';
}
