/**
 * Модуль 2, задание со * 1. Единый стиль для всей книги — эталонное решение.
 *
 * Сначала попробуйте написать сами, потом сравните.
 *
 * Нужен файл FormatOrders из урока 2.3: отсюда берём объект STYLE,
 * чтобы цвета и шрифт были заданы в одном месте на весь проект.
 */

// Служебные листы, которые не оформляем. Список — в одном месте, чтобы легко поменять.
const SKIP_SHEETS = ['Help', 'Settings'];

function applyStyleToAllSheets() {
  let styled = 0;
  let skipped = 0;

  SpreadsheetApp.getActive().getSheets().forEach(function (sheet) {
    const lastCol = sheet.getLastColumn();

    // Служебный или пустой лист пропускаем.
    // У пустого листа lastCol = 0, а getRange с нулём колонок упал бы с ошибкой.
    if (SKIP_SHEETS.indexOf(sheet.getName()) !== -1 || lastCol === 0) {
      skipped++;
      return;
    }

    sheet.getDataRange().setFontFamily(STYLE.FONT).setFontSize(10);
    sheet.getRange(1, 1, 1, lastCol)
      .setBackground(STYLE.HEADER_BG)
      .setFontColor(STYLE.HEADER_FONT)
      .setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, lastCol);
    styled++;
  });

  SpreadsheetApp.getActive().toast(
    'Sheets styled: ' + styled + ', skipped: ' + skipped, 'Workbook style', 5);
}
