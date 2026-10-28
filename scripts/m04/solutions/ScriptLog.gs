/**
 * Модуль 4, задание со *. Журнал запусков скриптов — эталонное решение.
 *
 * Сначала попробуйте написать сами, потом сравните.
 *
 * runLogged_ запускает любую функцию и записывает на лист «Script log»:
 * когда, кто, какую функцию, чем закончилось и сколько длилось.
 * Ошибка тоже попадает в журнал, но не «проглатывается»: пользователь её видит.
 *
 * Нужны файлы из курса: TidyExport (модуль 2), Backup и DebugPractice (модуль 4).
 */
const LOG_SHEET = 'Script log';
const LOG_HEADER = ['Time', 'User', 'Function', 'Result', 'Details'];

/** Пример: безопасная чистка выгрузки с записью в журнал. Запускать из меню или редактора. */
function loggedTidyExport() {
  runLogged_('tidyExportSafely', tidyExportSafely);
}

/** Пример с ошибкой: функция из тренажёра отладки ищет лист «Orders » с лишним пробелом. */
function loggedBugDemo() {
  runLogged_('bug1_countProblems', bug1_countProblems);
}

/**
 * Запускает fn и пишет результат в журнал.
 * @param {string} name   Имя для журнала — обычно имя самой функции.
 * @param {Function} fn   Функция, которую запускаем. Передаём без скобок: tidyExport, а не tidyExport().
 */
function runLogged_(name, fn) {
  const start = Date.now();
  try {
    fn();
    writeLog_(name, 'OK', 'Done in ' + ((Date.now() - start) / 1000).toFixed(1) + ' s');
  } catch (e) {
    writeLog_(name, 'ERROR', e.message);
    throw e; // пробрасываем дальше: ошибку увидят в журнале выполнения и в «Выполнениях»
  }
}

/** Дописывает строку в журнал. Если листа нет — создаёт его с шапкой. */
function writeLog_(name, result, details) {
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName(LOG_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(LOG_SHEET);
    sheet.appendRow(LOG_HEADER);
    sheet.getRange(1, 1, 1, LOG_HEADER.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  // Внутри домена компании getEmail() возвращает адрес того, кто запустил.
  // Если адрес недоступен (например, личный аккаунт), пишем «unknown».
  const user = Session.getActiveUser().getEmail() || 'unknown';
  sheet.appendRow([new Date(), user, name, result, details]);
}
