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
    .createMenu('🛠 Course')
    .addItem('Format the Orders sheet', 'formatOrdersSheet')
    .addItem('Tidy up the export', 'tidyExport')
    .addSeparator()
    .addItem('Open orders per assignee', 'askOpenOrdersByWorker')
    .addItem('Add order form', 'showAddOrderSidebar')
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu('Sandbox')
      .addItem('Refresh dates', 'refreshSandboxDates')
      .addItem('Add Close checkboxes', 'addCloseCheckboxes'))
    .addItem('What this menu does', 'showMenuHelp')
    .addToUi();
}

function showMenuHelp() {
  SpreadsheetApp.getUi().alert(
    'Course menu',
    'Format — header, number formats, overdue highlighting.\n' +
    'Tidy up the export — spaces, case, duplicates, dates.\n' +
    'Open orders — asks for an assignee ID and counts.\n' +
    'Form — add an order without typing into the sheet.',
    SpreadsheetApp.getUi().ButtonSet.OK);
}
