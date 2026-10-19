/**
 * Приймає результати онбордингу YBC і дописує їх у таблицю "YBC Onboarding".
 *
 * Куди що пишеться:
 *  - вкладка "Загальне": старт, проходження модулів 1-6, фінальний тест;
 *  - вкладка "Sales Warsaw": відповіді на тест сьомого модуля для посади
 *    "Сейлз-менеджер Варшава" (по одному рядку на відповідь + рядок "ПІДСУМОК").
 * Вкладки шукаються за номером (gid), потім за назвою. Якщо вкладки немає,
 * вона створюється.
 *
 * КОЛОНКИ ШУКАЮТЬСЯ ЗА НАЗВОЮ в першому рядку. Тому колонки в таблиці можна
 * переставляти, перейменовувати не можна: скрипт знайде кожну за заголовком
 * (Дата/час, Ім'я, Email, Спроба, Тип, Питання, Обрана відповідь,
 * Правильна відповідь, Результат, Підсумок спроби, № у спробі, Категорія,
 * ID питання). Якщо заголовка немає, ця колонка просто пропускається.
 *
 * ПІДСВІТКА: у вкладці "Sales Warsaw" увесь рядок, де в колонці "Результат"
 * стоїть "Помилка", фарбується в ніжно-червоний (колір MISTAKE_COLOR) через
 * умовне форматування. Воно діє на старі й нові рядки. Решту форматування
 * робите самі.
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

var MISTAKE_COLOR = "#F4CCCC";   // ніжно-червоний для клітинки "Помилка"

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
  }
}

// Текст, що починається з = + - @, Google Таблиці вважають формулою.
// Сайт публічний, тому такі значення екрануємо.
function clean_(value) {
  var s = value === undefined || value === null ? "" : String(value);
  if (s.length > 1000) s = s.substring(0, 1000);
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}

// Дописує один або кілька рядків (кожен це об'єкт {назва колонки: значення})
// у кінець вкладки, розкладаючи значення за заголовками першого рядка.
// Повертає номер першого записаного рядка.
function appendByHeaders_(sheet, rowObjects, textColumns) {
  var width = Math.max(sheet.getLastColumn(), 1);
  var headers = sheet.getRange(1, 1, 1, width).getValues()[0].map(function (h) {
    return String(h).trim();
  });

  // Колонки зі значеннями на кшталт "5/7" мають бути текстовими, інакше
  // Google перетворить їх на дату.
  (textColumns || []).forEach(function (name) {
    var idx = headers.indexOf(name);
    if (idx >= 0) sheet.getRange(1, idx + 1, sheet.getMaxRows(), 1).setNumberFormat("@");
  });

  var rows = rowObjects.map(function (obj) {
    return headers.map(function (h) {
      return Object.prototype.hasOwnProperty.call(obj, h) ? obj[h] : "";
    });
  });
  var start = sheet.getLastRow() + 1;
  sheet.getRange(start, 1, rows.length, headers.length).setValues(rows);
  return { start: start, headers: headers };
}

function writeGeneral_(ss, data) {
  var sheet = getTab_(ss, GENERAL_TAB);
  ensureHeaders_(sheet, GENERAL_HEADERS);
  appendByHeaders_(sheet, [{
    "Дата/час": new Date(),
    "Ім'я": clean_(data.name),
    "Посада": clean_(data.position),
    "Email": clean_(data.email),
    "Подія": clean_(data.event),
    "Деталі": clean_(data.details),
    "Бал/Рахунок": data.score !== undefined ? clean_(data.score) : ""
  }], ["Бал/Рахунок"]);
}

function writeAnswers_(ss, data) {
  var tab = ANSWER_TABS[data.tab];
  if (!tab) return;
  var sheet = getTab_(ss, tab);
  ensureHeaders_(sheet, ANSWER_HEADERS);

  var now = new Date();
  var answers = data.answers || [];
  var objects = [];
  for (var i = 0; i < answers.length && i < 60; i++) {
    var a = answers[i];
    objects.push({
      "Дата/час": now,
      "Ім'я": clean_(data.name),
      "Email": clean_(data.email),
      "Спроба": clean_(data.attempt),
      "Тип": clean_(data.kind),
      "№ у спробі": clean_(a.n),
      "ID питання": clean_(a.id),
      "Категорія": clean_(a.cat),
      "Питання": clean_(a.q),
      "Обрана відповідь": clean_(a.chosen),
      "Правильна відповідь": clean_(a.correct),
      "Результат": a.ok ? "Правильно" : "Помилка"
    });
  }
  objects.push({
    "Дата/час": now,
    "Ім'я": clean_(data.name),
    "Email": clean_(data.email),
    "Спроба": clean_(data.attempt),
    "Тип": clean_(data.kind),
    "Питання": "ПІДСУМОК",
    "Підсумок спроби": clean_(data.summary)
  });

  var res = appendByHeaders_(sheet, objects, ["Підсумок спроби", "№ у спробі"]);
  ensureHighlightRule_(sheet, res.headers);
}

// Номер колонки (1, 2, ...) у літеру (A, B, ... AA).
function colLetter_(n) {
  var s = "";
  while (n > 0) {
    var m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

/**
 * Умовне форматування: увесь рядок, де в колонці "Результат" стоїть
 * "Помилка", фарбується в MISTAKE_COLOR. Діє на старі й нові рядки
 * автоматично. Правило можна побачити й змінити в таблиці:
 * Формат -> Умовне форматування.
 */
function ensureHighlightRule_(sheet, headers) {
  var col = headers.indexOf("Результат");
  if (col < 0) return;
  var formula = '=$' + colLetter_(col + 1) + '2="Помилка"';
  var width = headers.length;
  var rows = Math.max(sheet.getMaxRows() - 1, 1);

  var rules = sheet.getConditionalFormatRules();
  var kept = [];
  var alreadyOk = false;
  for (var i = 0; i < rules.length; i++) {
    var cond = rules[i].getBooleanCondition();
    var isOurs = cond &&
      cond.getCriteriaType() === SpreadsheetApp.BooleanCriteria.CUSTOM_FORMULA &&
      /="Помилка"$/.test(String(cond.getCriteriaValues()[0]));
    if (!isOurs) { kept.push(rules[i]); continue; }
    var ranges = rules[i].getRanges();
    if (!alreadyOk && cond.getCriteriaValues()[0] === formula &&
        ranges.length === 1 && ranges[0].getNumColumns() === width && ranges[0].getRow() === 2) {
      alreadyOk = true;
      kept.push(rules[i]);
    }
    // інші наші застарілі правила (інша колонка чи ширина) відкидаємо
  }
  if (alreadyOk && kept.length === rules.length) return;
  if (!alreadyOk) {
    kept.push(SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(formula)
      .setBackground(MISTAKE_COLOR)
      .setRanges([sheet.getRange(2, 1, rows, width)])
      .build());
  }
  sheet.setConditionalFormatRules(kept);
}

// Запустіть один раз вручну (вибрати "setupHighlight" і натиснути Run),
// щоб підсвітка з'явилась на вже наявних рядках без нового тесту.
function setupHighlight() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  Object.keys(ANSWER_TABS).forEach(function (key) {
    var sheet = getTab_(ss, ANSWER_TABS[key]);
    var width = Math.max(sheet.getLastColumn(), 1);
    var headers = sheet.getRange(1, 1, 1, width).getValues()[0].map(function (h) {
      return String(h).trim();
    });
    ensureHighlightRule_(sheet, headers);
  });
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
// щоб перевірити запис і підсвітку в обох вкладках. Тестові рядки можна видалити.
function testWrite() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  writeGeneral_(ss, { name: "[тест]", position: "Тест", email: "test@example.com", event: "test", details: "Перевірка запису", score: "0" });
  writeAnswers_(ss, {
    tab: "Sales Warsaw", name: "[тест]", email: "test@example.com", attempt: 1, kind: "Тест",
    summary: "1/2 · Не пройдено",
    answers: [
      { n: 1, id: "sw-01", cat: "Обов'язкове", q: "Тестове питання 1", chosen: "Відповідь А", correct: "Відповідь А", ok: true },
      { n: 2, id: "sw-02", cat: "Обов'язкове", q: "Тестове питання 2", chosen: "Відповідь Б", correct: "Відповідь В", ok: false }
    ]
  });
}
