/**
 * Итоговый тест, практическое задание (вопрос И.26, 5 баллов).
 *
 * 1. Сделайте НОВУЮ копию песочницы (сайт курса → «Создать копию песочницы»)
 *    и ничего в ней не меняйте: код считается по исходным данным.
 * 2. Вставьте этот файл в редактор скриптов новой копии.
 * 3. Допишите функцию buildFinalTable_ по условию из теста.
 * 4. Запустите finalTask и выберите в тесте код из журнала выполнения.
 *
 * Функцию checksum_ не меняйте: она превращает вашу таблицу в короткий код.
 * Если в таблице ошибка хотя бы в одной цифре, код получится другим.
 */
function finalTask() {
  const rows = buildFinalTable_();
  console.log('Your table: ' + JSON.stringify(rows));
  console.log('Code for the test: ' + checksum_(rows));
}

/**
 * Условие. По листу «Orders» для каждого типа заказа возьмите только заказы
 * со статусом «Done» и посчитайте:
 *   - сколько таких заказов;
 *   - сумму часов (колонка «Hours»), округлённую до 0,1: Math.round(x * 10) / 10;
 *   - среднюю оценку QC, округлённую до 0,01: Math.round(x * 100) / 100.
 * Верните массив строк [тип, заказов, часов, средняя оценка] без шапки,
 * отсортированный по типу по алфавиту.
 */
function buildFinalTable_() {
  // Ваш код здесь
  return [];
}

/** Короткий код таблицы: первые 6 символов MD5 от её записи в виде текста. Не меняйте. */
function checksum_(rows) {
  const text = JSON.stringify(rows);
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, text, Utilities.Charset.UTF_8);
  return bytes.slice(0, 3)
    .map(function (b) { return ((b + 256) % 256).toString(16).padStart(2, '0'); })
    .join('')
    .toUpperCase();
}
