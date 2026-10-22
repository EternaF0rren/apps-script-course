/**
 * Модуль 7, уроки 1–2. Свои формулы.
 *
 * Функцию из этого файла можно писать прямо в ячейке, как СУММ:
 *   =HOURS_TEXT(I2)            → «4 ч 36 мин»
 *   =HOURS_TEXT(I2:I201)       → вся колонка одной формулой
 *   =ORDER_STATUS(C2; K2; E2)  → «Сдан вовремя»
 *
 * Имена латиницей и заглавными: так их легко отличить от встроенных формул,
 * и они одинаково работают при любом языке таблицы.
 * Комментарий с @customfunction включает подсказку при вводе формулы.
 */

/**
 * Переводит часы из числа в текст: 2.25 → «2 ч 15 мин».
 *
 * @param {number} hours Часы числом или диапазон с часами, например I2:I201.
 * @return Текст вида «2 ч 15 мин».
 * @customfunction
 */
function HOURS_TEXT(hours) {
  // Диапазон приходит двумерным массивом: обрабатываем каждую ячейку
  if (Array.isArray(hours)) {
    return hours.map(function (row) {
      return row.map(function (cell) { return HOURS_TEXT(cell); });
    });
  }
  if (hours === '' || hours === null) return '';
  if (typeof hours !== 'number') throw new Error('HOURS_TEXT: нужно число часов, а пришло «' + hours + '»');

  const totalMinutes = Math.round(hours * 60);
  const h = Math.floor(totalMinutes / 60);
  const m = totalMinutes % 60;
  if (h === 0) return m + ' мин';
  if (m === 0) return h + ' ч';
  return h + ' ч ' + m + ' мин';
}

/**
 * Достаёт ID заказа вида ORD-12345 из ссылки, письма или комментария.
 *
 * @param {string} text Текст или ссылка, например «https://crm.example/order/ORD-10083?tab=1».
 * @return ID заказа заглавными буквами или пустая строка, если ID нет.
 * @customfunction
 */
function EXTRACT_ORDER_ID(text) {
  if (Array.isArray(text)) {
    return text.map(function (row) {
      return row.map(function (cell) { return EXTRACT_ORDER_ID(cell); });
    });
  }
  const found = String(text).match(/ORD-\d+/i);
  return found ? found[0].toUpperCase() : '';
}

/**
 * Статус заказа по дедлайну и дате сдачи — вместо нескольких вложенных ЕСЛИ.
 *
 * @param {Date} deadline Дедлайн (колонка C) или диапазон дедлайнов.
 * @param {Date} done Дата сдачи (колонка K) или диапазон того же размера.
 * @param {string} status Статус (колонка E) или диапазон того же размера.
 * @param {number} hotHours Необязательно: за сколько часов до дедлайна заказ «горит». По умолчанию 4.
 * @return «Сдан вовремя», «Сдан с опозданием», «Просрочен», «Горит» или «В срок».
 * @customfunction
 */
function ORDER_STATUS(deadline, done, status, hotHours) {
  // Три диапазона одинакового размера: идём по строкам и берём i-ю строку из каждого
  if (Array.isArray(deadline)) {
    return deadline.map(function (row, i) {
      return [ORDER_STATUS(row[0], done[i][0], status[i][0], hotHours)];
    });
  }
  if (!(deadline instanceof Date)) return '';
  const limit = (typeof hotHours === 'number' ? hotHours : 4) * 60 * 60 * 1000;
  const now = new Date();

  if (done instanceof Date) return done <= deadline ? 'Сдан вовремя' : 'Сдан с опозданием';
  if (status === 'Готов') return 'Сдан вовремя';
  if (deadline < now) return 'Просрочен';
  if (deadline - now <= limit) return 'Горит';
  return 'В срок';
}

/**
 * Рабочие дни между двумя датами с учётом выходных и праздников страны.
 * Праздники берутся из диапазона: первая колонка — страна, вторая — дата.
 *
 * @param {Date} start Дата начала (включительно).
 * @param {Date} end Дата конца (включительно).
 * @param {string} country Страна так же, как она записана в списке праздников, например «Грузия».
 * @param {Array} holidays Диапазон с праздниками, например Праздники!A2:B.
 * @return Количество рабочих дней.
 * @customfunction
 */
function WORKDAYS_BY_COUNTRY(start, end, country, holidays) {
  if (!(start instanceof Date) || !(end instanceof Date)) return '';

  // Праздники нужной страны — в набор, чтобы быстро проверять «есть или нет»
  const daysOff = new Set();
  (Array.isArray(holidays) ? holidays : []).forEach(function (row) {
    if (String(row[0]).trim() === String(country).trim() && row[1] instanceof Date) {
      daysOff.add(dayKey_(row[1]));
    }
  });

  let count = 0;
  const day = new Date(start.getFullYear(), start.getMonth(), start.getDate());
  const last = new Date(end.getFullYear(), end.getMonth(), end.getDate());
  while (day <= last) {
    const weekDay = day.getDay(); // 0 — воскресенье, 6 — суббота
    if (weekDay !== 0 && weekDay !== 6 && !daysOff.has(dayKey_(day))) count++;
    day.setDate(day.getDate() + 1);
  }
  return count;
}

/** Дата → «2026-10-14», чтобы сравнивать дни без учёта времени. */
function dayKey_(date) {
  return date.getFullYear() + '-' + (date.getMonth() + 1) + '-' + date.getDate();
}

/**
 * Запустите один раз: создаёт лист «Праздники» с примером.
 * Даты — только для тренировки. Для работы заполните лист по официальному
 * календарю своей страны: переносы выходных и плавающие праздники
 * (например, Пасху) нужно добавить самим.
 */
function createHolidaysSheet() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName('Праздники') || ss.insertSheet('Праздники');
  const rows = [
    ['Страна', 'Дата', 'Праздник'],
    ['Грузия', new Date(2026, 9, 14), 'Светицховлоба'],
    ['Грузия', new Date(2026, 10, 23), 'Гиоргоба'],
    ['Грузия', new Date(2027, 0, 1), 'Новый год'],
    ['Грузия', new Date(2027, 0, 2), 'Новый год'],
    ['Грузия', new Date(2027, 0, 7), 'Рождество'],
    ['Польша', new Date(2026, 10, 11), 'День независимости'],
    ['Польша', new Date(2026, 11, 24), 'Сочельник'],
    ['Польша', new Date(2026, 11, 25), 'Рождество'],
    ['Польша', new Date(2027, 0, 1), 'Новый год'],
    ['Польша', new Date(2027, 0, 6), 'Богоявление'],
    ['Сербия', new Date(2026, 10, 11), 'День перемирия'],
    ['Сербия', new Date(2027, 0, 1), 'Новый год'],
    ['Сербия', new Date(2027, 0, 2), 'Новый год'],
    ['Сербия', new Date(2027, 0, 7), 'Рождество'],
  ];
  sheet.clearContents();
  sheet.getRange(1, 1, rows.length, 3).setValues(rows);
  sheet.getRange(1, 1, 1, 3).setFontWeight('bold');
  sheet.getRange(2, 2, rows.length - 1, 1).setNumberFormat('dd.mm.yyyy');
  ss.toast('Лист «Праздники» готов. Это пример — сверьте даты с официальным календарём.');
}
