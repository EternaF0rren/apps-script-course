/**
 * Модуль 1, урок 5. Пример библиотеки.
 *
 * Этот файл — не для копии песочницы, а для ОТДЕЛЬНОГО проекта на script.google.com
 * (например, «Team Shared Functions»). Другие таблицы подключают его как библиотеку
 * с идентификатором CourseLib и вызывают функции так: CourseLib.cleanName('  OLGA l. ').
 *
 * Правила хорошей библиотеки:
 *  - только общие помощники, которые нужны многим таблицам;
 *  - никаких имён листов и колонок конкретной таблицы — всё нужное передаётся аргументами;
 *  - служебные функции с «_» в конце: подключившим их не видно.
 */

/** Версия библиотеки. Меняйте при каждом новом развертывании, чтобы легко проверить, что подключено. */
function version() {
  return '1.0';
}

/**
 * «  OLGA l. » → «Olga L.»: без лишних пробелов, каждое слово с заглавной буквы.
 * @param {*} value Имя в любом виде.
 * @return {string} Аккуратное имя.
 */
function cleanName(value) {
  return squeezeSpaces_(value)
    .toLowerCase()
    .split(' ')
    .map(function (word) { return word.charAt(0).toUpperCase() + word.slice(1); })
    .join(' ');
}

/**
 * Текст → дата. Понимает 24.08.2026, 24/08/2026 и 2026-08-24.
 * Если в ячейке уже дата — возвращает её; если формат не узнан — возвращает исходное значение.
 * @param {*} value Дата или текст с датой.
 * @return {Date|*} Дата или исходное значение.
 */
function parseDate(value) {
  if (value instanceof Date) return value;
  const text = squeezeSpaces_(value);
  let m = text.match(/^(\d{1,2})[.\/](\d{1,2})[.\/](\d{4})$/);
  if (m) return new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
  m = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  return value;
}

/**
 * «3,5» → 3.5. Если число не получилось — возвращает исходное значение.
 * @param {*} value Число или текст с числом.
 * @return {number|*} Число или исходное значение.
 */
function parseNumber(value) {
  if (typeof value === 'number') return value;
  const n = Number(squeezeSpaces_(value).replace(',', '.'));
  return isNaN(n) || squeezeSpaces_(value) === '' ? value : n;
}

/** Служебная: убрать пробелы по краям и двойные внутри. С «_» — подключившим не видна. */
function squeezeSpaces_(value) {
  return String(value).trim().replace(/\s+/g, ' ');
}
