/**
 * Приймає результати онбордингу YBC і дописує їх у таблицю "YBC Onboarding".
 *
 * Куди що пишеться:
 *  - вкладка "Загальне": старт, проходження модулів 1-6, фінальний тест;
 *  - вкладка "Sales Warsaw": відповіді на тест сьомого модуля для посади
 *    "Сейлз-менеджер Варшава" (по одному рядку на відповідь + рядок "ПІДСУМОК").
 *  - вкладка "Прогрес" (службова, прихована): збережений прогрес кожного
 *    співробітника за поштою, щоб він не залежав від браузера й пристрою.
 *    Її не редагуйте вручну.
 *  - Хід проходження співробітників "Сейлз-менеджер Варшава" (початок,
 *    пройдені модулі, фінальний тест) дублюється рядками "Прогрес" у вкладку
 *    "Sales Warsaw", поруч із відповідями на тест.
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

// Посада -> ключ вкладки (ANSWER_TABS), куди дублюється хід проходження
// (старт, пройдені модулі, фінальний тест), щоб він був видний поруч із відповідями.
var POSITION_TAB = {
  "Сейлз-менеджер Варшава": "Sales Warsaw"
};

function eventLabel_(event, details) {
  if (event === "start") return "Початок онбордингу";
  if (event === "module_pass") return "Модуль пройдено: " + details;
  if (event === "module_read") return "Модуль завершено: " + details;
  if (event === "final_pass") return "Фінальний тест пройдено";
  return String(event);
}

// Рядок "Прогрес" у вкладці посади: що й коли пройшла людина.
function progressObject_(data, when) {
  var passed = data.event === "module_pass" || data.event === "module_read" || data.event === "final_pass";
  return {
    "Дата/час": when,
    "Ім'я": clean_(data.name),
    "Email": clean_(data.email),
    "Тип": "Прогрес",
    "Питання": clean_(eventLabel_(data.event, data.details)),
    "Результат": passed ? "Пройдено" : "",
    "Підсумок спроби": data.event === "start" ? "" : clean_(data.score)
  };
}

function writeProgressRow_(ss, tabKey, data) {
  var tab = ANSWER_TABS[tabKey];
  if (!tab) return;
  var sheet = getTab_(ss, tab);
  ensureHeaders_(sheet, ANSWER_HEADERS);
  appendByHeaders_(sheet, [progressObject_(data, new Date())], ["Підсумок спроби"]);
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

  var tabKey = POSITION_TAB[String(data.position || "")];
  if (tabKey) writeProgressRow_(ss, tabKey, data);
}

/**
 * Одноразово: переносить у вкладку "Sales Warsaw" уже наявні події проходження
 * (старт, модулі, фінальний тест) співробітників "Сейлз-менеджер Варшава" з
 * вкладки "Загальне". Повторний запуск дублікатів не створює.
 * Запуск: вибрати "backfillSalesWarsaw" і натиснути Run.
 */
function backfillSalesWarsaw() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  progressSheet_(ss); // заодно ховає службову вкладку "Прогрес"
  var general = getTab_(ss, GENERAL_TAB);
  var last = general.getLastRow();
  if (last < 2) return;
  var gHeaders = general.getRange(1, 1, 1, general.getLastColumn()).getValues()[0].map(function (h) { return String(h).trim(); });
  var rows = general.getRange(2, 1, last - 1, general.getLastColumn()).getValues();
  var idx = function (name) { return gHeaders.indexOf(name); };

  var tab = ANSWER_TABS["Sales Warsaw"];
  var sheet = getTab_(ss, tab);
  ensureHeaders_(sheet, ANSWER_HEADERS);
  var sHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(function (h) { return String(h).trim(); });
  var sLast = sheet.getLastRow();
  var existing = {};
  if (sLast > 1) {
    var sRows = sheet.getRange(2, 1, sLast - 1, sHeaders.length).getValues();
    var iEmail = sHeaders.indexOf("Email"), iQ = sHeaders.indexOf("Питання"), iT = sHeaders.indexOf("Тип");
    sRows.forEach(function (r) {
      if (iT >= 0 && r[iT] === "Прогрес") existing[String(r[iEmail]).toLowerCase() + "|" + r[iQ]] = true;
    });
  }

  var toAdd = [];
  rows.forEach(function (r) {
    if (String(r[idx("Посада")]) !== "Сейлз-менеджер Варшава") return;
    var event = String(r[idx("Подія")]);
    if (["start", "module_pass", "module_read", "final_pass"].indexOf(event) < 0) return;
    var data = {
      name: r[idx("Ім'я")], email: r[idx("Email")], event: event,
      details: r[idx("Деталі")], score: r[idx("Бал/Рахунок")]
    };
    var key = String(data.email).toLowerCase() + "|" + eventLabel_(event, data.details);
    if (existing[key]) return;
    existing[key] = true;
    var when = r[idx("Дата/час")];
    if (!(when instanceof Date)) {
      var parsed = new Date(when);
      when = isNaN(parsed.getTime()) ? new Date() : parsed;
    }
    toAdd.push(progressObject_(data, when));
  });
  if (toAdd.length) appendByHeaders_(sheet, toAdd, ["Підсумок спроби"]);
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

// ---------------------------------------------------------------------
// Прогрес співробітників (щоб він не залежав від браузера й пристрою)
// Вкладка "Прогрес": один рядок на пошту, прогрес зберігається як JSON.
// Цю вкладку не редагуйте вручну: сайт читає її при вході.
// ---------------------------------------------------------------------
var PROGRESS_TAB = { gid: -1, name: "Прогрес" };
var PROGRESS_HEADERS = ["Email", "Ім'я", "Посада", "Оновлено (мс)", "Прогрес (JSON, не редагувати)"];

function progressSheet_(ss) {
  var sheet = getTab_(ss, PROGRESS_TAB);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(PROGRESS_HEADERS);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, PROGRESS_HEADERS.length).setFontWeight("bold");
    sheet.getRange("A:E").setNumberFormat("@");
  }
  // Це службова вкладка з кодами: ховаємо її. Скрипт читає її й прихованою.
  try {
    if (!sheet.isSheetHidden()) sheet.hideSheet();
  } catch (err) {}
  return sheet;
}

function findProgressRow_(sheet, email) {
  var last = sheet.getLastRow();
  if (last < 2) return 0;
  var values = sheet.getRange(2, 1, last - 1, 1).getValues();
  for (var i = 0; i < values.length; i++) {
    if (String(values[i][0]).toLowerCase() === email) return i + 2;
  }
  return 0;
}

function saveProgress_(ss, data) {
  var email = String(data.email || "").trim().toLowerCase();
  if (!email) return;
  var json = JSON.stringify(data.progress || {});
  if (json.length > 45000) return;
  var incoming = Number(data.updatedAt) || 0;
  var sheet = progressSheet_(ss);
  var row = findProgressRow_(sheet, email);
  var record = [email, clean_(data.name), clean_(data.position), String(incoming), json];
  if (row) {
    // Старіша копія (наприклад, зі застарілої вкладки) не затирає новішу.
    var stored = Number(sheet.getRange(row, 4).getValue()) || 0;
    if (incoming < stored) return;
    sheet.getRange(row, 1, 1, record.length).setValues([record]);
  } else {
    sheet.getRange(sheet.getLastRow() + 1, 1, 1, record.length).setValues([record]);
  }
}

function loadProgress_(ss, data) {
  var email = String(data.email || "").trim().toLowerCase();
  if (!email) return { found: false };
  var sheet = progressSheet_(ss);
  var row = findProgressRow_(sheet, email);
  if (!row) return { found: false };
  var cells = sheet.getRange(row, 4, 1, 2).getValues()[0];
  try {
    return { found: true, updatedAt: Number(cells[0]) || 0, progress: JSON.parse(cells[1]) };
  } catch (err) {
    return { found: false };
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
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
    if (!data.email) return json_({ status: "ignored" });

    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    if (data.type === "load") {
      return json_(loadProgress_(ss, data));
    }
    if (data.type === "progress") {
      saveProgress_(ss, data);
    } else if (data.type === "answers") {
      writeAnswers_(ss, data);
    } else if (data.event) {
      writeGeneral_(ss, data);
    } else {
      return json_({ status: "ignored" });
    }
    return json_({ status: "ok" });
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
