const SHEET_NAME = 'Заявки';
const HEADERS = [
  'Дата и время',
  'Имя',
  'Телефон',
  'Возраст ребёнка',
  'Дата встречи',
  'Комментарий',
  'Страница',
  'Браузер'
];

function doGet() {
  return jsonResponse({ ok: true, service: 'psychologist-signup' });
}

function doPost(event) {
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(10000);

    const params = event && event.parameter ? event.parameter : {};

    // Скрытое поле заполняют боты. Для них возвращаем успех, но строку не добавляем.
    if (params.website) {
      return jsonResponse({ ok: true });
    }

    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    if (!spreadsheet) {
      throw new Error('Скрипт должен быть создан из Google Таблицы.');
    }

    let sheet = spreadsheet.getSheetByName(SHEET_NAME);
    if (!sheet) {
      sheet = spreadsheet.insertSheet(SHEET_NAME);
    }

    if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
      sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
      sheet.setFrozenRows(1);
    }

    sheet.appendRow([
      new Date(),
      safeCell(params.name, 200),
      safeCell(params.phone, 100),
      safeCell(params.childAge, 100),
      safeCell(params.date, 100),
      safeCell(params.comment, 5000),
      safeCell(params.page, 1000),
      safeCell(params.userAgent, 1000)
    ]);

    return jsonResponse({ ok: true });
  } catch (error) {
    console.error(error);
    return jsonResponse({ ok: false, error: String(error.message || error) });
  } finally {
    if (lock.hasLock()) {
      lock.releaseLock();
    }
  }
}

function safeCell(value, maxLength) {
  let text = String(value || '').trim().slice(0, maxLength);

  // Не даём пользовательскому вводу стать формулой в Google Таблице.
  if (/^[=+\-@]/.test(text)) {
    text = "'" + text;
  }

  return text;
}

function jsonResponse(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}
