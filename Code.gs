// ─── Триггер: создаёт пункт меню при открытии таблицы ───────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Курсы ЦБ")
    .addItem("Обновить курсы", "fetchCbrRates")
    .addToUi();
}

// ─── Валюты, которые запишем в таблицу ───────────────────────────────────────
const CURRENCIES = ["USD", "EUR", "CNY", "GBP", "JPY"];

/**
 * Основная функция: запрашивает XML с сайта ЦБ РФ,
 * парсит курсы валют и записывает их в активный лист.
 *
 * Триггер  — пункт меню «Курсы ЦБ → Обновить курсы» (добавляется через onOpen).
 * API      — https://www.cbr.ru/scripts/XML_daily.asp (бесплатно, без ключа).
 * Статус   — записывается в ячейку A1 (ошибка) или F1 (время обновления).
 */
function fetchCbrRates() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const statusCell = sheet.getRange("A1"); // ячейка для сообщений об ошибках
  const API_URL = "https://www.cbr.ru/scripts/XML_daily.asp";

  // ── 1. HTTP-запрос ───────────────────────────────────────────────────────────
  let response;
  try {
    response = UrlFetchApp.fetch(API_URL, { muteHttpExceptions: true });
  } catch (e) {
    statusCell.setValue("❌ Ошибка сети: " + e.message);
    return;
  }

  // ── 2. Проверка HTTP-статуса ─────────────────────────────────────────────────
  if (response.getResponseCode() !== 200) {
    statusCell.setValue(`❌ HTTP ${response.getResponseCode()}`);
    return;
  }

  // ── 3. Парсинг XML ───────────────────────────────────────────────────────────
  let root;
  try {
    // ЦБ РФ отдаёт тело в Windows-1251; декодируем байты явно,
    // затем убираем XML-декларацию с encoding — иначе XmlService
    // попытается перекодировать строку ещё раз и получится мусор
    const bytes = response.getContent();
    const xmlText = Utilities.newBlob(bytes).getDataAsString("windows-1251");
    const cleaned = xmlText.replace(/<\?xml[^?]*\?>\s*/, "");
    root = XmlService.parse(cleaned).getRootElement();
  } catch (e) {
    statusCell.setValue("❌ Некорректный XML: " + e.message);
    return;
  }

  // ── 4. Запись шапки ─────────────────────────────────────────────────────────
  sheet
    .getRange("A1:D1")
    .setValues([["Валюта", "Название", "Курс (руб.)", "Единиц"]])
    .setFontWeight("bold");

  // ── 5. Извлечение данных и запись строк ─────────────────────────────────────
  const valutes = root.getChildren("Valute");
  const rows = [];

  valutes.forEach((v) => {
    const code = v.getChildText("CharCode");
    if (!CURRENCIES.includes(code)) return;

    const nominal = Number(v.getChildText("Nominal"));
    // ЦБ РФ использует запятую как десятичный разделитель — заменяем на точку
    const value = parseFloat(v.getChildText("Value").replace(",", "."));
    const rate = (value / nominal).toFixed(4); // курс за 1 единицу валюты

    rows.push([code, v.getChildText("Name"), rate, nominal]);
  });

  if (rows.length === 0) {
    statusCell.setValue("⚠️ Нет данных по выбранным валютам");
    return;
  }

  // Очищаем старые данные (кроме шапки) и записываем свежие
  sheet.getRange(2, 1, sheet.getLastRow(), 4).clearContent();
  sheet.getRange(2, 1, rows.length, 4).setValues(rows);

  // ── 6. Статус успеха + время обновления ─────────────────────────────────────
  const now = new Date();
  sheet.getRange("F1").setValue("Обновлено: " + now.toLocaleString("ru-RU"));
}
