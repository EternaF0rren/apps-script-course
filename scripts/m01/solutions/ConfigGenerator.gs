/**
 * Модуль 1, задание со * 1. Генератор Config.gs — эталонное решение.
 *
 * Сначала попробуйте написать сами, потом сравните.
 *
 * generateConfig проходит по всем листам таблицы и печатает в журнал
 * готовый текст для Config.gs: объект SHEETS и объекты с номерами колонок.
 * Таблицу функция не меняет — только читает первую строку каждого листа.
 */
function generateConfig() {
  const sheets = SpreadsheetApp.getActive().getSheets();

  // 1. Объект SHEETS: ключ — имя листа заглавными буквами, значение — имя как есть
  let text = '// Имена листов\nconst SHEETS = {\n';
  sheets.forEach(function (sheet) {
    const name = sheet.getName();
    const key = toConstName_(name);
    if (key) {
      text += '  ' + key + ": '" + escapeQuotes_(name) + "',\n";
    } else {
      // Имя не на латинице — ставим номер листа, ключ переименуете вручную
      text += '  SHEET_' + sheet.getIndex() + ": '" + escapeQuotes_(name) + "', // переименуйте ключ\n";
    }
  });
  text += '};\n';

  // 2. Номера колонок для каждого листа со строкой заголовков
  sheets.forEach(function (sheet) {
    const name = sheet.getName();
    const lastColumn = sheet.getLastColumn();

    // getRange с нулём колонок падает, поэтому пустой лист проверяем заранее
    const headers = lastColumn > 0
      ? sheet.getRange(1, 1, 1, lastColumn).getValues()[0]
      : [];

    const filled = headers.filter(function (value) {
      return String(value).trim() !== '';
    });

    text += '\n';
    if (filled.length < 2) {
      text += '// Лист «' + name + '»: строки заголовков нет, колонки не выписаны\n';
      return;
    }

    const objectName = (toConstName_(name) || 'SHEET_' + sheet.getIndex()) + '_COLS';
    text += '// Колонки листа «' + name + '» (считаем с 1, как в таблице)\n';
    text += 'const ' + objectName + ' = {\n';
    headers.forEach(function (header, index) {
      const title = String(header).trim();
      if (!title) return; // пустую колонку пропускаем
      const column = index + 1;
      const key = toConstName_(title);
      if (key) {
        text += '  ' + key + ': ' + column + ',\n';
      } else {
        // Заголовок не на латинице — ключ придумать нельзя, оставляем подсказку
        text += '  COL_' + column + ': ' + column + ', // «' + title + '» — переименуйте ключ\n';
      }
    });
    text += '};\n';
  });

  console.log(text);
}

/**
 * «QC Score» → «QC_SCORE», «Report Template» → «REPORT_TEMPLATE».
 * Если в тексте есть не латинские буквы («ID заявки», «Статус») — возвращает пустую строку:
 * иначе от «ID заявки» остался бы ключ «ID», который легко перепутать.
 * Ключ не может начинаться с цифры, поэтому «90th» → «_90TH».
 */
function toConstName_(text) {
  if (/[^\u0000-\u007F]/.test(String(text))) return '';
  let key = String(text).trim().toUpperCase()
    .replace(/[^A-Z0-9]+/g, '_')   // всё, кроме латиницы и цифр, — в подчёркивание
    .replace(/^_+|_+$/g, '');      // убираем подчёркивания по краям
  if (/^[0-9]/.test(key)) key = '_' + key;
  return key;
}

/** Апостроф в имени листа сломал бы строку в одинарных кавычках. */
function escapeQuotes_(text) {
  return String(text).replace(/\\/g, '\\\\').replace(/'/g, "\\'");
}
