/**
 * Приймає результати онбордингу YBC і дописує їх у таблицю "YBC Onboarding".
 *
 * Куди що пишеться:
 *  - вкладка "Загальне": старт, проходження модулів 1-6, фінальний тест;
 *  - вкладка "Sales Warsaw": відповіді на тест сьомого модуля для посади
 *    "Сейлз-менеджер Варшава" (по одному рядку на відповідь + рядок "ПІДСУМОК").
 * Вкладки шукаються за номером (gid), потім за назвою. Якщо вкладки немає,
 * вона створюється. Заголовки додаються, якщо вкладка порожня.
 *
 * ГРУПУВАННЯ ПО СПІВРОБІТНИКУ:
 * У кожній вкладці все, що належить одній пошті, збирається в групу:
 * рядок-заголовок із назвою "УЧАСНИК" (ім'я, пошта, короткий підсумок)
 * і під ним згруповані рядки подій чи відповідей. Зліва від таблиці
 * з'являється "+" / "-", щоб розгортати й згортати групу. Нові групи
 * створюються згорнутими.
 *
 * ПІДСВІТКА: рядки з результатом "Помилка" фарбуються в ніжно-червоний
 * (колір задається змінною MISTAKE_COLOR).
 *
 * ЯК ПІДКЛЮЧИТИ (2 хвилини):
 * 1. Увійдіть в Google під акаунтом, який має РЕДАГУВАТИ таблицю
 *    (власник таблиці: ybc.team.hr@gmail.com).
 * 2. Відкрийте https://script.google.com і натисніть "Новий проєкт".
 * 3. Видаліть заглушку, вставте весь цей файл, збережіть (Ctrl+S).
 * 4. Розгорнути -> Новий деплой -> тип "Веб-додаток":
 *    - Виконати як: Я
 *    - Хто має доступ: Будь-хто
 * 5. Натисніть "Розгорнути", підтвердіть дозволи Google.
 * 6. Скопіюйте URL веб-додатка (закінчується на /exec) і надішліть його
 *    в чат. Він вставляється в index.html у змінну SHEET_WEBHOOK_URL.
 *
 * Якщо код скрипта змінюється, треба: Розгорнути -> Керувати деплоями ->
 * олівець -> Версія: "Нова версія" -> Розгорнути. Інакше зміни не діють.
 * URL при цьому не змінюється.
 */

var SPREADSHEET_ID = "1sQ4SH4PRsh0HmFU2-atE6ewLFvLDDLzvXn9YJ7TT4Ec";

var MARKER = "УЧАСНИК";          // позначка рядка-заголовка групи
var MISTAKE_COLOR = "#FDECEC";   // ніжно-червоний для рядків з помилками
var HEADER_COLOR = "#EFEFEF";    // фон рядка-заголовка групи
var WHITE = "#FFFFFF";

var GENERAL_TAB = { gid: 0, name: "Загальне" };
var GENERAL_HEADERS = ["Дата/час", "Ім'я", "Посада", "Email", "Подія", "Деталі", "Бал/Рахунок"];

// Дозволені вкладки для відповідей на тести (ключ приходить із сайту).
var ANSWER_TABS = {
  "Sales Warsaw": { gid: 1685318998, name: "Sales Warsaw" }
};
var ANSWER_HEADERS = [
  "Дата/час", "Ім'я", "Email", "Спроба", "Тип", "№ у спробі", "ID питання", "Категорія",
  "Питання", "Обрана відповідь", "Правильна відповідь", "Результат", "Підсумок спроби"
];

function getTab_(ss, tab) {
  var sheets = ss.getSheets();
  var i;
  for (i = 0; i < sheets.length; i++) {
    if (sheets[i].getSheetId() === tab.gid) return sheets[i];
  }
  for (i = 0; i < sheets.length; i++) {
    if (sheets[i].getName() === tab.name) return sheets[i];
  }
  return ss.insertSheet(tab.name);
}

function ensureHeaders_(sheet, headers) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
    sheet.getRange("A:A").setNumberFormat("dd.mm.yyyy hh:mm:ss");
  }
  // "+" / "-" групи розташовані біля рядка-заголовка зверху, а не під групою.
  sheet.setRowGroupControlAfter(false);
}

// Текст, що починається з = + - @, Google Таблиці вважають формулою.
// Сайт публічний, тому такі значення екрануємо.
function clean_(value) {
  var s = value === undefined || value === null ? "" : String(value);
  if (s.length > 1000) s = s.substring(0, 1000);
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}

// ---------------------------------------------------------------------
// Групи по пошті
// ---------------------------------------------------------------------

// Знімає наявне групування з рядків і створює одну групу, потім згортає її.
function regroup_(sheet, from, to) {
  if (to < from) return;
  var rng = sheet.getRange(from, 1, to - from + 1, 1);
  try {
    var guard = 0;
    while (guard++ < 5 && (sheet.getRowGroupDepth(from) > 0 || sheet.getRowGroupDepth(to) > 0)) {
      rng.shiftRowGroupDepth(-1);
    }
  } catch (err) {
    // запасний шлях: знімаємо групування по одному рядку
    for (var r = from; r <= to; r++) {
      var g = 0;
      while (g++ < 5 && sheet.getRowGroupDepth(r) > 0) {
        sheet.getRange(r, 1, 1, 1).shiftRowGroupDepth(-1);
      }
    }
  }
  rng.shiftRowGroupDepth(1);
  try {
    sheet.getRowGroup(from, 1).collapse();
  } catch (err) {}
}

/**
 * Дописує рядки newRows у групу співробітника з поштою email.
 * Якщо групи ще немає, створює її в кінці вкладки.
 * cfg: { ncols, emailCol, markerCol, buildHeader(email, rows), isMistake(row), isSummary(row) }
 */
function appendGrouped_(sheet, cfg, email, newRows) {
  var ncols = cfg.ncols;
  var last = sheet.getLastRow();
  var values = last > 0 ? sheet.getRange(1, 1, last, ncols).getValues() : [];
  var key = String(email).toLowerCase();

  // 0-based індекси рядків-заголовків груп (рядок 0 це заголовки колонок)
  var headerIdx = [];
  for (var i = 1; i < values.length; i++) {
    if (values[i][cfg.markerCol - 1] === MARKER) headerIdx.push(i);
  }
  var found = -1;
  for (var k = 0; k < headerIdx.length; k++) {
    if (String(values[headerIdx[k]][cfg.emailCol - 1]).toLowerCase() === key) { found = k; break; }
  }

  var n = newRows.length;
  var headerRow, childStart, childEnd, allChild;
  if (found >= 0) {
    var h = headerIdx[found];
    var blockEnd = found + 1 < headerIdx.length ? headerIdx[found + 1] : last; // 1-based останній рядок групи
    headerRow = h + 1;
    var existing = values.slice(h + 1, blockEnd);
    sheet.insertRowsAfter(blockEnd, n);
    sheet.getRange(blockEnd + 1, 1, n, ncols).setValues(newRows);
    childStart = headerRow + 1;
    childEnd = blockEnd + n;
    allChild = existing.concat(newRows);
    var newStart = blockEnd + 1;
  } else {
    headerRow = last + 1;
    var blank = [];
    for (var c = 0; c < ncols; c++) blank.push("");
    sheet.getRange(headerRow, 1, n + 1, ncols).setValues([blank].concat(newRows));
    childStart = headerRow + 1;
    childEnd = headerRow + n;
    allChild = newRows;
    var newStart = headerRow + 1;
  }

  // Заголовок групи з коротким підсумком
  var headerRange = sheet.getRange(headerRow, 1, 1, ncols);
  headerRange.setValues([cfg.buildHeader(email, allChild)]);
  headerRange.setBackground(HEADER_COLOR).setFontWeight("bold");

  // Форматування нових рядків (вставлені рядки успадковують формат сусідніх)
  var bg = [], fw = [];
  for (var r = 0; r < n; r++) {
    var row = newRows[r];
    var color = cfg.isMistake(row) ? MISTAKE_COLOR : WHITE;
    var bold = cfg.isSummary(row) ? "bold" : "normal";
    var bgRow = [], fwRow = [];
    for (var cc = 0; cc < ncols; cc++) { bgRow.push(color); fwRow.push(bold); }
    bg.push(bgRow);
    fw.push(fwRow);
  }
  var newRange = sheet.getRange(newStart, 1, n, ncols);
  newRange.setBackgrounds(bg);
  newRange.setFontWeights(fw);

  regroup_(sheet, childStart, childEnd);
}

// ---------------------------------------------------------------------
// Вкладка "Загальне"
// ---------------------------------------------------------------------
var GENERAL_CFG = {
  ncols: 7, emailCol: 4, markerCol: 5,
  isMistake: function (row) { return false; },
  isSummary: function (row) { return false; },
  buildHeader: function (email, rows) {
    var name = "", position = "", lastDate = "", modules = {}, finalDone = false;
    for (var i = 0; i < rows.length; i++) {
      var r = rows[i];
      if (r[1]) name = r[1];
      if (r[2]) position = r[2];
      if (r[0]) lastDate = r[0];
      if (r[4] === "module_pass" || r[4] === "module_read") modules[String(r[5])] = true;
      if (r[4] === "final_pass") finalDone = true;
    }
    var count = Object.keys(modules).length;
    return [lastDate, name, position, email, MARKER,
      "Пройдено модулів: " + count + " з 7" + (finalDone ? " · фінальний тест пройдено" : ""), ""];
  }
};

function writeGeneral_(ss, data) {
  var sheet = getTab_(ss, GENERAL_TAB);
  ensureHeaders_(sheet, GENERAL_HEADERS);
  // "5/7" без текстового формату перетворюється на дату.
  sheet.getRange("G:G").setNumberFormat("@");
  appendGrouped_(sheet, GENERAL_CFG, clean_(data.email), [[
    new Date(),
    clean_(data.name),
    clean_(data.position),
    clean_(data.email),
    clean_(data.event),
    clean_(data.details),
    data.score !== undefined ? clean_(data.score) : ""
  ]]);
}

// ---------------------------------------------------------------------
// Вкладки з відповідями на тести (Sales Warsaw)
// ---------------------------------------------------------------------
var ANSWERS_CFG = {
  ncols: 13, emailCol: 3, markerCol: 5,
  isMistake: function (row) { return row[11] === "Помилка"; },
  isSummary: function (row) { return row[8] === "ПІДСУМОК"; },
  buildHeader: function (email, rows) {
    var name = "", lastDate = "", attempts = 0, lastRes = "", best = "", bestNum = -1;
    for (var i = 0; i < rows.length; i++) {
      var r = rows[i];
      if (r[1]) name = r[1];
      if (r[0]) lastDate = r[0];
      if (r[8] === "ПІДСУМОК" && r[4] === "Тест") {
        attempts++;
        lastRes = String(r[12]);
        var num = parseInt(lastRes, 10);
        if (!isNaN(num) && num > bestNum) { bestNum = num; best = lastRes; }
      }
    }
    var text = "Спроб тесту: " + attempts;
    if (attempts) text += " · остання: " + lastRes + " · найкраща: " + best;
    return [lastDate, name, email, "", MARKER, "", "", "", text, "", "", "", ""];
  }
};

function writeAnswers_(ss, data) {
  var tab = ANSWER_TABS[data.tab];
  if (!tab) return;
  var sheet = getTab_(ss, tab);
  ensureHeaders_(sheet, ANSWER_HEADERS);
  sheet.getRange("M:M").setNumberFormat("@");
  sheet.getRange("F:F").setNumberFormat("@");

  var now = new Date();
  var answers = data.answers || [];
  var rows = [];
  for (var i = 0; i < answers.length && i < 60; i++) {
    var a = answers[i];
    rows.push([
      now, clean_(data.name), clean_(data.email), clean_(data.attempt), clean_(data.kind),
      clean_(a.n), clean_(a.id), clean_(a.cat), clean_(a.q), clean_(a.chosen), clean_(a.correct),
      a.ok ? "Правильно" : "Помилка", ""
    ]);
  }
  rows.push([
    now, clean_(data.name), clean_(data.email), clean_(data.attempt), clean_(data.kind),
    "", "", "", "ПІДСУМОК", "", "", "", clean_(data.summary)
  ]);
  appendGrouped_(sheet, ANSWERS_CFG, clean_(data.email), rows);
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var data = {};
    try {
      data = JSON.parse(e.postData.contents);
    } catch (err) {
      data = {};
    }
    // Порожній або некоректний запит (наприклад, ручний запуск) нічого не пише.
    if (!data.email) {
      return ContentService.createTextOutput(JSON.stringify({ status: "ignored" }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    if (data.type === "answers") {
      writeAnswers_(ss, data);
    } else {
      writeGeneral_(ss, data);
    }

    return ContentService
      .createTextOutput(JSON.stringify({ status: "ok" }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  return ContentService.createTextOutput("YBC onboarding webhook працює.");
}

// Запустіть цю функцію один раз вручну (вибрати "testWrite" і натиснути Run),
// щоб перевірити запис, групи й підсвітку в обох вкладках.
// Тестові рядки можна видалити.
function testWrite() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  writeGeneral_(ss, { name: "[тест]", position: "Тест", email: "test@example.com", event: "start", details: "Перевірка запису", score: "0" });
  writeGeneral_(ss, { name: "[тест]", position: "Тест", email: "test@example.com", event: "module_pass", details: "Про компанію", score: "7/8" });
  writeAnswers_(ss, {
    tab: "Sales Warsaw", name: "[тест]", email: "test@example.com", attempt: 1, kind: "Тест",
    summary: "1/2 · Не пройдено",
    answers: [
      { n: 1, id: "sw-01", cat: "Обов'язкове", q: "Тестове питання 1", chosen: "Відповідь А", correct: "Відповідь А", ok: true },
      { n: 2, id: "sw-02", cat: "Обов'язкове", q: "Тестове питання 2", chosen: "Відповідь Б", correct: "Відповідь В", ok: false }
    ]
  });
  writeAnswers_(ss, {
    tab: "Sales Warsaw", name: "[тест]", email: "test@example.com", attempt: 2, kind: "Тест",
    summary: "2/2 · Пройдено",
    answers: [
      { n: 1, id: "sw-02", cat: "Обов'язкове", q: "Тестове питання 2", chosen: "Відповідь В", correct: "Відповідь В", ok: true },
      { n: 2, id: "sw-01", cat: "Обов'язкове", q: "Тестове питання 1", chosen: "Відповідь А", correct: "Відповідь А", ok: true }
    ]
  });
}
