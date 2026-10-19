/**
 * Модуль 1, урок 1.3. Первый запуск.
 *
 * Выберите функцию helloSheet в списке над кодом и нажмите «Выполнить».
 * Результат появится в журнале выполнения внизу экрана.
 * При первом запуске Google попросит разрешение на доступ к таблице.
 */
function helloSheet() {
  const ss = SpreadsheetApp.getActive();
  console.log('Таблица: ' + ss.getName());

  ss.getSheets().forEach(function (sheet) {
    // getLastRow() — номер последней заполненной строки
    console.log('Лист «' + sheet.getName() + '»: строк с данными — ' + sheet.getLastRow());
  });
}

/**
 * То же самое, но ответ видно прямо в таблице — во всплывающем окне.
 * Работает, только если функция запущена из таблицы (меню или кнопка, модуль 3).
 */
function helloToast() {
  const orders = SpreadsheetApp.getActive().getSheetByName('Заказы');
  const count = orders.getLastRow() - 1; // минус строка заголовка
  SpreadsheetApp.getActive().toast('В таблице ' + count + ' заказов', 'Привет от скрипта', 5);
}
