/**
 * GPS-трекер → Google Таблица
 * Скрипт Apps Script (привязан к таблице: Расширения → Apps Script).
 *
 * 1. Вставьте этот код в Code.gs
 * 2. Задайте TOKEN (любая строка-пароль) — он должен совпадать с настройками приложения
 * 3. Разверните: Развернуть → Новое развёртывание → Тип: Веб-приложение
 *    → Выполнять от: Я → Доступ: все
 * 4. Скопируйте URL вида https://script.google.com/macros/s/.../exec
 *    и вставьте его в настройках приложения (index.html)
 */

const TOKEN = 'school-2026';          // общий пароль (пустая строка = без проверки)
const SHEET_NAME = 'GPS';             // имя листа (создастся автоматически)

const HEADERS = [
  'Время', 'Широта', 'Долгота', 'Точность, м',
  'Скорость, м/с', 'Скорость, км/ч', 'Высота, м',
  'Курс, °', 'Метка', 'Устройство'
];

/** Проверка, что веб-приложение живо (откройте URL в браузере). */
function doGet() {
  return json_({ ok: true, app: 'gps-tracker', sheet: SHEET_NAME });
}

/** Приём точек от приложения. */
function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);                       // защита от одновременных записей

    const data = JSON.parse(e.postData.contents || '{}');

    if (TOKEN && data.token !== TOKEN) {
      return json_({ ok: false, error: 'Неверный токен' });
    }

    const points = Array.isArray(data.points) ? data.points : [];
    if (!points.length) return json_({ ok: true, added: 0 });

    const sh = getSheet_();
    const rows = points.map(p => [
      p.time || '',
      num_(p.lat), num_(p.lon), num_(p.acc),
      num_(p.speed),
      p.speed != null && isFinite(p.speed) ? Math.round(p.speed * 3.6 * 10) / 10 : '',
      num_(p.alt), num_(p.head),
      p.label || '',
      p.device || ''
    ]);

    sh.getRange(sh.getLastRow() + 1, 1, rows.length, HEADERS.length).setValues(rows);

    return json_({ ok: true, added: rows.length });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (ignored) {}
  }
}

/** Лист с заголовками и форматом чисел. */
function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS])
      .setFontWeight('bold').setBackground('#dbe9ff');
    sh.setFrozenRows(1);
    // формат: широта/долгота, высота, скорость
    sh.getRange(2, 2, sh.getMaxRows() - 1, 2).setNumberFormat('0.0000000');
    sh.getRange(2, 6, sh.getMaxRows() - 1, 3).setNumberFormat('0.0#');
    sh.autoResizeColumns(1, HEADERS.length);
  }
  return sh;
}

function num_(v) {
  return (v !== undefined && v !== null && isFinite(v)) ? Number(v) : '';
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
