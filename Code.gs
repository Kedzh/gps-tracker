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
const KEEP_ROWS = 60;                 // сколько ПОСЛЕДНИХ точек хранить в таблице

const HEADERS = [
  'Время', 'Широта', 'Долгота', 'Точность, м',
  'Скорость, м/с', 'Скорость, км/ч', 'Высота, м',
  'Курс, °', 'Метка', 'Устройство'
];

/** Проверка, что веб-приложение живо (откройте URL в браузере).
 *  ?action=data&token=XXX — отдать точки картой (track.html). */
function doGet(e) {
  const p = (e && e.parameter) || {};

  if (p.action === 'data') {
    if (TOKEN && p.token !== TOKEN) return json_({ ok: false, error: 'Неверный токен' });
    try {
      return json_({ ok: true, points: readPoints_(), updated: now_() });
    } catch (err) {
      return json_({ ok: false, error: String(err) });
    }
  }
  return json_({ ok: true, app: 'gps-tracker', sheet: SHEET_NAME });
}

/** Все точки листа в JSON для карты (только последние KEEP_ROWS). */
function readPoints_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sh) return [];

  const last = sh.getLastRow();
  if (last < 2) return [];

  const start = Math.max(2, last - KEEP_ROWS + 1);      // хвост таблицы
  const rows = sh.getRange(start, 1, last - start + 1, HEADERS.length).getValues();
  return rows
    .map(r => ({
      time:    fmtDate_(r[0]),
      lat:     r[1], lon:    r[2], acc:   r[3],
      speed:   r[4],                      // м/с
      alt:     r[6], head:   r[7],
      label:   r[8], device: r[9]
    }))
    .filter(p => isFinite(p.lat) && isFinite(p.lon));
}

/** Удаляет сверху всё, кроме KEEP_ROWS последних точек. */
function trim_(sh) {
  const dataRows = sh.getLastRow() - 1;
  if (dataRows > KEEP_ROWS) {
    sh.deleteRows(2, dataRows - KEEP_ROWS);            // строка 1 — заголовок
  }
}

function fmtDate_(d) {
  if (d instanceof Date) return Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd.MM.yyyy HH:mm:ss');
  return d === undefined || d === null ? '' : String(d);
}

function now_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd.MM.yyyy HH:mm:ss');
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
    trim_(sh);                                  // оставляем только последние KEEP_ROWS точек

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
