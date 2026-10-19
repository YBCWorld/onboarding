/**
 * Приймає результати онбордингу YBC і дописує їх у таблицю "YBC Onboarding".
 *
 * Куди що пишеться:
 *  - вкладка "Загальне": старт, проходження модулів 1-6, фінальний тест
 *    (по одному рядку на подію, як і раніше);
 *  - вкладка "Sales Warsaw": відповіді на тест сьомого модуля для посади
 *    "Сейлз-менеджер Варшава" (по одному рядку на відповідь + рядок "ПІДСУМОК").
 * Вкладки шукаються за номером (gid), потім за назвою. Якщо вкладки немає,
 * вона створюється. Заголовки додаються, якщо вкладка порожня.
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
 * Якщо код скрипта змінюється, треба: Керувати деплоями -> редагувати ->
 * Нова версія. Інакше зміни не діють. URL при цьому не змінюється.
 */

var SPREADSHEET_ID = "1sQ4SH4PRsh0HmFU2-atE6ewLFvLDDLzvXn9YJ7TT4Ec";

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
}

// Текст, що починається з = + - @, Google Таблиці вважають формулою.
// Сайт публічний, тому такі значення екрануємо.
function clean_(value) {
  var s = value === undefined || value === null ? "" : String(value);
  if (s.length > 1000) s = s.substring(0, 1000);
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}

function writeGeneral_(ss, data) {
  var sheet = getTab_(ss, GENERAL_TAB);
  ensureHeaders_(sheet, GENERAL_HEADERS);
  // "5/7" без текстового формату перетворюється на дату.
  sheet.getRange("G:G").setNumberFormat("@");
  sheet.appendRow([
    new Date(),
    clean_(data.name),
    clean_(data.position),
    clean_(data.email),
    clean_(data.event),
    clean_(data.details),
    data.score !== undefined ? clean_(data.score) : ""
  ]);
}

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
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, ANSWER_HEADERS.length).setValues(rows);
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

// Запустіть цю функцію один раз вручну (кнопка "Виконати"), щоб перевірити,
// що рядки з'являються в обох вкладках. Тестові рядки можна видалити.
function testWrite() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  writeGeneral_(ss, { name: "[тест]", position: "Тест", email: "test@example.com", event: "test", details: "Перевірка запису", score: "0" });
  writeAnswers_(ss, {
    tab: "Sales Warsaw", name: "[тест]", email: "test@example.com", attempt: 1, kind: "Тест",
    summary: "1/1 · Пройдено",
    answers: [{ n: 1, id: "sw-01", cat: "Обов'язкове", q: "Тестове питання", chosen: "Відповідь А", correct: "Відповідь А", ok: true }]
  });
}
