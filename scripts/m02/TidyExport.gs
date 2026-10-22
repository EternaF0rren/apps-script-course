/**
 * Модуль 2, урок 4. «Привести в порядок» одной кнопкой.
 *
 * Лист «Выгрузка» — типичная выгрузка из другой системы:
 * лишние пробелы, РАЗНЫЙ регистр, дубли, пустые строки, даты текстом
 * в трёх форматах, часы с запятой. Функция tidyExport чинит всё разом.
 *
 * Важно: функция меняет данные. Сначала пробуйте на копии листа
 * (правый клик по вкладке → «Создать копию»).
 */
function tidyExport() {
  const sheet = SpreadsheetApp.getActive().getSheetByName('Выгрузка');

  // 1. Встроенная чистка пробелов: убирает пробелы по краям и двойные внутри
  sheet.getDataRange().trimWhitespace();

  // 2. Читаем все данные одним вызовом — это двумерный массив [строка][колонка]
  const values = sheet.getDataRange().getValues();
  const header = values[0];

  // 3. Чиним каждую строку и выбрасываем пустые
  const rows = values.slice(1)
    .filter(function (row) { return row.join('') !== ''; })
    .map(function (row) {
      return [
        row[0],                  // id — без изменений
        toNameCase(row[1]),      // исполнитель: «ОЛЬГА Л.» → «Ольга Л.»
        capitalizeFirst(row[2]), // статус: «НА ПРОВЕРКЕ» → «На проверке»
        parseDate(row[3]),       // дата: текст → настоящая дата
        parseNumber(row[4]),     // часы: «3,5» → 3.5
      ];
    });

  // 4. Записываем обратно одним вызовом и стираем «хвост» старых строк
  sheet.getRange(2, 1, sheet.getMaxRows() - 1, header.length).clearContent();
  sheet.getRange(2, 1, rows.length, header.length).setValues(rows);

  // 5. Удаляем дубли (одинаковые строки целиком)
  sheet.getRange(1, 1, rows.length + 1, header.length).removeDuplicates();

  // 6. Единый вид
  sheet.getRange('D2:D').setNumberFormat('dd.MM.yyyy');
  sheet.getRange('E2:E').setNumberFormat('0.0');

  const left = sheet.getLastRow() - 1;
  SpreadsheetApp.getActive().toast(
    'Было строк: ' + (values.length - 1) + ', стало: ' + left, 'Выгрузка в порядке', 5);
}

/** «  ОЛЬГА л. » → «Ольга Л.»: первая буква каждого слова заглавная. */
function toNameCase(value) {
  return String(value)
    .trim()
    .toLowerCase()
    .split(' ')
    .map(function (word) { return word.charAt(0).toUpperCase() + word.slice(1); })
    .join(' ');
}

/** «  НА ПРОВЕРКЕ» → «На проверке»: заглавная только первая буква. */
function capitalizeFirst(value) {
  const text = String(value).trim().toLowerCase();
  return text.charAt(0).toUpperCase() + text.slice(1);
}

/**
 * Превращает текст в дату. Понимает 24.08.2026, 2026-08-24 и 24/08/2026.
 * Если в ячейке уже дата — возвращает её как есть.
 */
function parseDate(value) {
  if (value instanceof Date) return value;
  const text = String(value).trim();
  let m = text.match(/^(\d{1,2})[.\/](\d{1,2})[.\/](\d{4})$/);
  if (m) return new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
  m = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  return value; // не узнали формат — оставляем как было, чтобы ничего не потерять
}

/** «3,5» → 3.5. Если число не получилось — оставляем исходное значение. */
function parseNumber(value) {
  if (typeof value === 'number') return value;
  const n = Number(String(value).replace(',', '.').trim());
  return isNaN(n) ? value : n;
}
