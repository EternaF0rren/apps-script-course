/**
 * Модуль 1, урок 3. Первый запуск.
 *
 * Выберите функцию helloSheet в списке над кодом и нажмите «Выполнить».
 * Результат появится в журнале выполнения внизу экрана.
 * При первом запуске Google попросит разрешение на доступ к таблице.
 */
function helloSheet() {
  const ss = SpreadsheetApp.getActive();
  console.log('Spreadsheet: ' + ss.getName());

  ss.getSheets().forEach(function (sheet) {
    // getLastRow() — номер последней заполненной строки
    console.log('Sheet "' + sheet.getName() + '": rows with data — ' + sheet.getLastRow());
  });
}

/**
 * То же самое, но ответ видно прямо в таблице — во всплывающем окне.
 * Работает, только если функция запущена из таблицы (меню или кнопка, модуль 3).
 */
function helloToast() {
  const orders = SpreadsheetApp.getActive().getSheetByName('Orders');
  const count = orders.getLastRow() - 1; // минус строка заголовка
  SpreadsheetApp.getActive().toast('The spreadsheet has ' + count + ' orders', 'Hello from the script', 5);
}
