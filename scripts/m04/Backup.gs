/**
 * Модуль 4, урок 4.4. Страховка перед опасными изменениями.
 *
 * backupSheet_ делает копию листа рядом с оригиналом, например
 * «Заказы · копия 19.10 21:05». Вызывайте её первой строкой в любой функции,
 * которая удаляет или переписывает данные. Что-то пошло не так — данные целы.
 */
function backupSheet_(sheetName) {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error('Не найден лист «' + sheetName + '»');

  const stamp = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'dd.MM HH:mm');
  const copy = sheet.copyTo(ss).setName(sheetName + ' · копия ' + stamp);
  copy.setTabColor('#9e9e9e'); // серая вкладка — сразу видно, что это копия
  return copy;
}

/** Пример: перед чисткой выгрузки сначала делаем копию листа. */
function tidyExportSafely() {
  backupSheet_('Выгрузка');
  tidyExport(); // функция из модуля 2
}

/**
 * Задание со *: удалить старые копии, оставив три последние.
 * Удаляет только листы с « · копия » в названии.
 */
function cleanupOldBackups() {
  const ss = SpreadsheetApp.getActive();
  const copies = ss.getSheets().filter(function (s) { return s.getName().indexOf(' · копия ') !== -1; });
  const toDelete = copies.slice(0, Math.max(0, copies.length - 3)); // копии добавляются в конец, старые — первыми

  if (toDelete.length === 0) return;
  const ui = SpreadsheetApp.getUi();
  const names = toDelete.map(function (s) { return s.getName(); }).join('\n');
  if (ui.alert('Удалить старые копии?', names, ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  toDelete.forEach(function (s) { ss.deleteSheet(s); });
}
