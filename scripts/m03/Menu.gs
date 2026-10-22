/**
 * Модуль 3, урок 1. Своё меню.
 *
 * onOpen — особое имя. Таблица сама запускает эту функцию каждый раз,
 * когда кто-то её открывает. Мы используем это, чтобы добавить меню «Курс».
 *
 * Пункты меню вызывают функции из других файлов проекта:
 * formatOrdersSheet и tidyExport (модуль 2), refreshSandboxDates (модуль 1).
 * Если какой-то функции в проекте нет, пункт выдаст ошибку при нажатии.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🛠 Курс')
    .addItem('Оформить лист «Заказы»', 'formatOrdersSheet')
    .addItem('Привести выгрузку в порядок', 'tidyExport')
    .addSeparator()
    .addItem('Сколько открытых заказов у исполнителя?', 'askOpenOrdersByWorker')
    .addItem('Форма «Добавить заказ»', 'showAddOrderSidebar')
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu('Песочница')
      .addItem('Обновить даты', 'refreshSandboxDates')
      .addItem('Добавить колонку с флажками «Закрыть»', 'addCloseCheckboxes'))
    .addItem('Что умеет это меню', 'showMenuHelp')
    .addToUi();
}

function showMenuHelp() {
  SpreadsheetApp.getUi().alert(
    'Меню «Курс»',
    'Оформить — шапка, форматы, подсветка просрочек.\n' +
    'Привести выгрузку в порядок — пробелы, регистр, дубли, даты.\n' +
    'Сколько открытых заказов — спросит ID исполнителя и посчитает.\n' +
    'Форма — добавить заказ без ручного ввода в таблицу.',
    SpreadsheetApp.getUi().ButtonSet.OK);
}
