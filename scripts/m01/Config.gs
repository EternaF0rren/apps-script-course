/**
 * Модуль 1, урок 1.4. Порядок в проекте.
 *
 * Все имена листов и настройки — в одном месте.
 * Если лист переименуют, правим только здесь, а не в десяти функциях.
 */

// Имена листов песочницы
const SHEETS = {
  ORDERS: 'Заказы',
  EMPLOYEES: 'Сотрудники',
  HOURS: 'Часы',
  EXPORT: 'Выгрузка',
  ARCHIVE: 'Архив',
  SUMMARY: 'Сводка',
  SETTINGS: 'Настройки',
};

// Номера колонок листа «Заказы» (считаем с 1, как в таблице)
const ORDER_COLS = {
  ID: 1,
  CREATED: 2,
  DEADLINE: 3,
  TYPE: 4,
  STATUS: 5,
  EMPLOYEE_ID: 6,
  EMPLOYEE: 7,
  SUPERVISOR: 8,
  HOURS: 9,
  QC: 10,
  DONE: 11,
  COMMENT: 12,
};

/**
 * Возвращает лист по имени и сразу говорит понятную ошибку,
 * если такого листа нет (вместо загадочного «Cannot read properties of null»).
 */
function getSheet(name) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sheet) {
    throw new Error('Не найден лист «' + name + '». Проверьте название в Config.gs');
  }
  return sheet;
}

/**
 * Задание *: читаем настройки с листа «Настройки».
 * Возвращает объект вида { ЛИСТ_ЗАКАЗОВ: 'Заказы', ПОРОГ_ПРОСРОЧКИ_Ч: 4, ... }.
 * Пример: const settings = getSettings(); settings.ПОРОГ_ПРОСРОЧКИ_Ч
 */
function getSettings() {
  const rows = getSheet(SHEETS.SETTINGS).getDataRange().getValues();
  const settings = {};
  rows.slice(1).forEach(function (row) {
    const key = String(row[0]).trim();
    if (key) settings[key] = row[1];
  });
  return settings;
}

function testSettings() {
  console.log(getSettings());
}
