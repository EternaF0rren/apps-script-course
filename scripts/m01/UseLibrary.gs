/**
 * Модуль 1, урок 1.5. Пользуемся библиотекой из своей таблицы.
 *
 * Перед запуском подключите библиотеку: Редактор → «Библиотеки» → «+» →
 * ID скрипта → «Найти» → версия → идентификатор CourseLib → «Добавить».
 * Код библиотеки — в папке scripts/m01/library (файл CourseLib.gs).
 *
 * Функция ничего не меняет в таблице: только печатает в журнал,
 * как библиотека приведёт в порядок первые строки листа «Выгрузка».
 */
function tryCourseLib() {
  console.log('Подключена библиотека CourseLib, версия ' + CourseLib.version());

  const rows = SpreadsheetApp.getActive().getSheetByName('Выгрузка')
    .getRange('A2:E6').getValues();

  rows.forEach(function (row) {
    console.log(
      CourseLib.cleanName(row[1]) + ' | ' +
      CourseLib.parseDate(row[3]) + ' | ' +
      CourseLib.parseNumber(row[4]) + ' ч'
    );
  });
}
