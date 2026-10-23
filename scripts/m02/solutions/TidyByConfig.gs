/**
 * Модуль 2, задание со * 2. Чистка любой выгрузки по настройке — эталонное решение.
 *
 * Сначала попробуйте написать сами, потом сравните.
 *
 * Нужен файл TidyExport из урока 2.4: помощники toNameCase, parseDate и parseNumber
 * берём оттуда, а не копируем.
 *
 * Под новую выгрузку меняется только объект TIDY — код функции не трогаем.
 */
const TIDY = {
  SHEET: 'Export',
  // Что делать с каждой колонкой. Колонки, которых здесь нет, остаются как есть.
  //   text   — только убрать лишние пробелы
  //   name   — «  OLGA l. » → «Olga L.», «in review» → «In Review»
  //   date   — текст «24.08.2026», «2026-08-24», «24/08/2026» → настоящая дата
  //   number — «3,5» → 3.5
  COLUMNS: { A: 'text', B: 'name', C: 'name', D: 'date', E: 'number' },
  WARN_BG: '#fff2cc', // жёлтым подсвечиваем то, что распознать не удалось
};

// Формат ячеек для каждого типа колонки
const TIDY_FORMATS = { date: 'dd.MM.yyyy', number: '0.0' };

function tidyByConfig() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(TIDY.SHEET);
  if (!sheet) throw new Error('Sheet "' + TIDY.SHEET + '" not found. Check TIDY.SHEET');

  // 1. Пробелы — встроенной чисткой, затем читаем лист целиком
  sheet.getDataRange().trimWhitespace();
  const values = sheet.getDataRange().getValues();
  const width = values[0].length;
  const columns = readColumns_(width);

  // 2. Выбрасываем пустые строки, каждую оставшуюся чиним по настройке
  const rows = values.slice(1)
    .filter(function (row) { return row.join('') !== ''; })
    .map(function (row) {
      const fixed = row.slice(); // копия строки: исходный массив не меняем
      columns.forEach(function (col) {
        fixed[col.index] = cleanValue_(fixed[col.index], col.type);
      });
      return fixed;
    });

  // 3. Стираем старые данные вместе со старой подсветкой и записываем новые
  sheet.getRange(2, 1, sheet.getMaxRows() - 1, width).clearContent().setBackground(null);
  if (rows.length > 0) sheet.getRange(2, 1, rows.length, width).setValues(rows);
  sheet.getRange(1, 1, rows.length + 1, width).removeDuplicates();

  // 4. Единый вид дат и чисел
  columns.forEach(function (col) {
    const format = TIDY_FORMATS[col.type];
    if (format) sheet.getRange(2, col.index + 1, sheet.getMaxRows() - 1, 1).setNumberFormat(format);
  });

  // 5. Подсвечиваем то, что не распознали. Делаем это после удаления дублей:
  //    иначе подсветка «съехала» бы вместе со строками
  const unrecognized = findUnrecognized_(sheet, columns);
  if (unrecognized.length > 0) {
    sheet.getRangeList(unrecognized).setBackground(TIDY.WARN_BG);
  }

  ss.toast('Rows: ' + (sheet.getLastRow() - 1) + '. Not recognized: ' + unrecognized.length +
    (unrecognized.length ? ' (highlighted in yellow)' : ''), 'Tidy: ' + TIDY.SHEET, 8);
}

/** Превращает настройку { D: 'date' } в список { index: 3, letter: 'D', type: 'date' }. */
function readColumns_(width) {
  return Object.keys(TIDY.COLUMNS).map(function (letter) {
    const index = letter.toUpperCase().charCodeAt(0) - 65; // A → 0, B → 1, …
    const type = TIDY.COLUMNS[letter];
    if (index < 0 || index >= width) {
      throw new Error('Column ' + letter + ' from TIDY.COLUMNS is outside the data (A–' +
        String.fromCharCode(64 + width) + ')');
    }
    if (['text', 'name', 'date', 'number'].indexOf(type) === -1) {
      throw new Error('Unknown type "' + type + '" for column ' + letter);
    }
    return { index: index, letter: letter.toUpperCase(), type: type };
  });
}

/** Чинит одно значение по типу колонки. Пустая ячейка остаётся пустой, а не становится 0. */
function cleanValue_(value, type) {
  if (value === '') return '';
  if (type === 'name') return toNameCase(value);
  if (type === 'date') return parseDate(value);
  if (type === 'number') return parseNumber(value);
  return value; // text: пробелы уже убраны
}

/**
 * Читает лист заново и возвращает адреса ячеек, которые не удалось распознать:
 * в колонке дат — не дата, в колонке чисел — не число. Например ['D5', 'E12'].
 */
function findUnrecognized_(sheet, columns) {
  const values = sheet.getDataRange().getValues();
  const result = [];
  values.slice(1).forEach(function (row, i) {
    columns.forEach(function (col) {
      const value = row[col.index];
      if (value === '') return;
      const bad = (col.type === 'date' && !(value instanceof Date)) ||
                  (col.type === 'number' && typeof value !== 'number');
      if (bad) result.push(col.letter + (i + 2)); // +2: строка заголовков и счёт с 1
    });
  });
  return result;
}
