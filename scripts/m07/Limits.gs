/**
 * Модуль 7, урок 7.3. Ограничения своих формул — на живых примерах.
 * Попробуйте каждую функцию в ячейке и посмотрите, что произойдёт.
 */

/**
 * ✗ Так не работает: своя формула не может менять другие ячейки.
 * В ячейке будет ошибка «You do not have permission to call…».
 * Если нужно что-то записать — это работа для пункта меню или кнопки (модуль 3).
 *
 * @customfunction
 */
function BROKEN_WRITE() {
  SpreadsheetApp.getActive().getSheetByName('Сводка').getRange('Z1').setValue('привет');
  return 'записал';
}

/**
 * ✗ Работает, но не обновляется: функция сама читает лист «Сотрудники».
 * Если поменять супервайзера на листе «Сотрудники», ячейка с формулой
 * останется со старым значением — Таблицы не знают, что формула зависит от этого листа.
 *
 * Пример: =SUPERVISOR_OF_STALE(F2)
 * @param {string} employeeId ID исполнителя, например D-004.
 * @customfunction
 */
function SUPERVISOR_OF_STALE(employeeId) {
  const rows = SpreadsheetApp.getActive().getSheetByName('Сотрудники').getDataRange().getValues();
  const found = rows.filter(function (row) { return row[0] === employeeId; })[0];
  return found ? found[5] : 'нет такого ID';
}

/**
 * ✓ Правильно: справочник передаётся аргументом.
 * Теперь любая правка на листе «Сотрудники» пересчитает формулу.
 *
 * Пример: =SUPERVISOR_OF(F2; Сотрудники!A2:F)
 * @param {string} employeeId ID исполнителя или диапазон ID, например F2:F201.
 * @param {Array} employees Диапазон листа «Сотрудники»: ID в первой колонке, супервайзер в шестой.
 * @return Имя супервайзера.
 * @customfunction
 */
function SUPERVISOR_OF(employeeId, employees) {
  // Справочник строим один раз — даже если на входе 200 ID
  const supervisorById = {};
  employees.forEach(function (row) { supervisorById[row[0]] = row[5]; });

  if (Array.isArray(employeeId)) {
    return employeeId.map(function (row) { return [supervisorById[row[0]] || '']; });
  }
  return supervisorById[employeeId] || 'нет такого ID';
}

/**
 * Обходной путь для формул, которые зависят от текущего времени (как ORDER_STATUS).
 * Сама формула пересчитывается, только когда меняются её аргументы, а NOW() и RAND()
 * передавать аргументом нельзя. Поэтому раз в час скрипт записывает время в ячейку,
 * а формула получает эту ячейку пятым аргументом.
 *
 * 1. Запустите stampNow один раз — на листе «Настройки» появится строка СЕЙЧАС.
 * 2. Поставьте триггер «По времени → Таймер по часам» на stampNow (модуль 5).
 * 3. Пишите формулу так: =ORDER_STATUS(C2; K2; E2; 4; Настройки!B8)
 *    (в песочнице строка СЕЙЧАС — восьмая; в своей таблице проверьте номер строки).
 *    Лишний аргумент функция не использует — он нужен только как сигнал «пересчитай меня».
 */
function stampNow() {
  const sheet = SpreadsheetApp.getActive().getSheetByName('Настройки');
  const rows = sheet.getDataRange().getValues();
  let rowNumber = rows.findIndex(function (row) { return row[0] === 'СЕЙЧАС'; }) + 1;
  if (rowNumber === 0) {
    rowNumber = rows.length + 1;
    sheet.getRange(rowNumber, 1, 1, 3).setValues([['СЕЙЧАС', '', 'Время последнего обновления (модуль 7)']]);
  }
  sheet.getRange(rowNumber, 2).setValue(new Date()).setNumberFormat('dd.mm.yyyy hh:mm');
}
