/************************************************************
 * STROY PODRYAD — CRM кандидатов
 * Google Apps Script
 *
 * Рабочая таблица: Кандидаты — Светский лес
 * Лист: Кандидаты
 * Часовой пояс: Europe/Moscow
 ************************************************************/

const CALENDAR_ID = "fa4f2d45502f74f4e23d8cadcc234a0c2e1daed5293e3c7858433b2d1e08a97f@group.calendar.google.com";

const MOSCOW_TZ = "Europe/Moscow";
const SHEET_ID_PROPERTY = "CANDIDATE_SHEET_ID";
const SHEET_NAME = "Кандидаты";

const CANDIDATE_STATUSES = [
  "Новый",
  "Записан",
  "Приехал на объект",
  "На испытательном сроке",
  "Работает",
  "Отказ",
  "Не приехал",
  "Не прошёл испытательный срок"
];


/************************************************************
 * 1. ОТКРЫТИЕ ВЕБ-ПРИЛОЖЕНИЯ
 ************************************************************/


function doGet(e) {
  const page = e && e.parameter ? e.parameter.page : "";

  if (page === "mobile") {
    return HtmlService
      .createHtmlOutputFromFile("Mobile")
      .setTitle("STROY PODRYAD — Mobile")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  if (page === "foreman") {
    return HtmlService
      .createHtmlOutputFromFile("Foreman")
      .setTitle("STROY PODRYAD — Прораб")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  return HtmlService
    .createHtmlOutputFromFile("Index")
    .setTitle("STROY PODRYAD — CRM")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/************************************************************
 * 2. ВНЕШНИЙ API
 ************************************************************/

function doPost(e) {
  let event = null;

  try {
    if (!e || !e.postData || !e.postData.contents) {
      throw new Error("Не получены данные запроса.");
    }

    const data = JSON.parse(e.postData.contents);

    const name = String(data.name || "").trim();
    const phone = String(data.phone || "").trim();
    const specialty = String(data.specialty || "").trim();
    const date = String(data.date || "").trim();
    const time = String(data.time || "").trim();

    const objectName = String(
      data.object || "ЖК «Светский лес»"
    ).trim();

    if (!name || !phone || !specialty) {
      throw new Error("Заполнены не все обязательные поля.");
    }

    const start = parseMoscowDateTime_(date, time);
    const end = new Date(start.getTime() + 30 * 60 * 1000);

    const calendar = CalendarApp.getCalendarById(CALENDAR_ID);

    if (!calendar) {
      throw new Error("Календарь не найден или нет доступа.");
    }

    event = calendar.createEvent(
      "Встреча с кандидатом: " + name,
      start,
      end,
      {
        description:
          "Кандидат: " + name +
          "\nТелефон: " + phone +
          "\nСпециальность: " + specialty +
          "\nОбъект: " + objectName
      }
    );

    event.removeAllReminders();
    event.addPopupReminder(60);

    return jsonResponse_({
      success: true,
      message: "Встреча добавлена в календарь!"
    });

  } catch (error) {
    if (event) {
      try {
        event.deleteEvent();
      } catch (deleteError) {
        Logger.log("Не удалось удалить событие: " + deleteError);
      }
    }

    return jsonResponse_({
      success: false,
      message: error.message || String(error)
    });
  }
}


/************************************************************
 * 3. ДОБАВЛЕНИЕ КАНДИДАТА
 ************************************************************/

function createCandidate(data) {
  const lock = LockService.getScriptLock();

  let event = null;
  let sheet = null;
  let historySheet = null;
  let meetingsSheet = null;

  let writtenRow = null;
  let writtenHistoryRow = null;
  let writtenMeetingRow = null;

  try {
    lock.waitLock(30000);

    if (!data || typeof data !== "object") {
      throw new Error("Не получены данные кандидата.");
    }

    const name = String(data.name || "").trim();
    const phone = String(data.phone || "").trim();
    const specialty = String(data.specialty || "").trim();
    const date = String(data.date || "").trim();
    const time = String(data.time || "").trim();
    const comment = String(data.comment || "").trim();
    const source = String(data.source || "Не указан").trim();
    const objectName = String(data.object || "").trim();

    if (!name) throw new Error("Укажи ФИО кандидата.");
    if (!phone) throw new Error("Укажи номер телефона.");
    if (!specialty) throw new Error("Выбери специальность.");
    if (!objectName) throw new Error("Выбери строительный объект.");

    const phoneDigits = phone.replace(/\D/g, "");

    if (phoneDigits.length < 7 || phoneDigits.length > 15) {
      throw new Error("Проверь номер телефона.");
    }

    const start = parseMoscowDateTime_(date, time);
    const end = new Date(start.getTime() + 30 * 60 * 1000);

    const properties = PropertiesService.getScriptProperties();
    const spreadsheetId = properties.getProperty(SHEET_ID_PROPERTY);

    if (!spreadsheetId) {
      throw new Error("Не указан ID рабочей таблицы в Script Properties.");
    }

    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);

    sheet = spreadsheet.getSheetByName(SHEET_NAME);
    historySheet = spreadsheet.getSheetByName("История");
    meetingsSheet = spreadsheet.getSheetByName("Встречи");

    if (!sheet) throw new Error('Лист "Кандидаты" не найден.');
    if (!historySheet) throw new Error('Лист "История" не найден.');
    if (!meetingsSheet) throw new Error('Лист "Встречи" не найден.');

    if (sheet.getRange("I1").getValue() !== "ID кандидата") {
      throw new Error("Не найден заголовок ID кандидата в столбце I.");
    }

    if (sheet.getRange("J1").getValue() !== "Объект") {
      throw new Error("В ячейке J1 должен быть заголовок «Объект».");
    }

    // Проверка дубликата по телефону (сравниваем последние 10 цифр).
    const phoneKey = phoneDigits.length >= 10
      ? phoneDigits.slice(-10)
      : phoneDigits;

    const lastCandidateRow = sheet.getLastRow();

    if (lastCandidateRow >= 2 && phoneKey) {
      const phoneColumn = sheet
        .getRange(2, 3, lastCandidateRow - 1, 1)
        .getValues();

      for (let i = 0; i < phoneColumn.length; i++) {
        const existingDigits = String(phoneColumn[i][0] || "").replace(/\D/g, "");
        const existingKey = existingDigits.length >= 10
          ? existingDigits.slice(-10)
          : existingDigits;

        if (existingKey && existingKey === phoneKey) {
          const rowIndex = i + 2;
          const existingName = String(sheet.getRange(rowIndex, 2).getValue() || "").trim();
          const existingStatus = String(sheet.getRange(rowIndex, 7).getValue() || "").trim();
          const existingId = String(sheet.getRange(rowIndex, 9).getValue() || "").trim();

          return {
            success: false,
            duplicate: true,
            message:
              "Кандидат с таким телефоном уже есть в базе.\n" +
              "ФИО: " + (existingName || "не указано") + "\n" +
              "Статус: " + (existingStatus || "не указан") + "\n" +
              "ID: " + (existingId || "—") + "\n\n" +
              "Новая запись не создана. Открой существующего кандидата или измени номер.",
            existing: {
              row: rowIndex,
              name: existingName,
              status: existingStatus,
              candidateId: existingId,
              phone: String(phoneColumn[i][0] || "")
            }
          };
        }
      }
    }

    const candidateId = "CAND-" +
      Utilities.getUuid().substring(0, 8).toUpperCase();

    const meetingId = "MEET-" +
      Utilities.getUuid().substring(0, 8).toUpperCase();

    const calendar = CalendarApp.getCalendarById(CALENDAR_ID);

    if (!calendar) {
      throw new Error("Календарь не найден или нет доступа.");
    }

    event = calendar.createEvent(
      "Встреча с кандидатом: " + name,
      start,
      end,
      {
        description:
          "ID кандидата: " + candidateId +
          "\nКандидат: " + name +
          "\nТелефон: " + phone +
          "\nСпециальность: " + specialty +
          "\nОбъект: " + objectName +
          "\nИсточник: " + source +
          "\nКомментарий: " + (comment || "Нет")
      }
    );

    event.removeAllReminders();
    event.addPopupReminder(60);

    const now = new Date();

    // Запись кандидата в A:J.
    writtenRow = Math.max(sheet.getLastRow() + 1, 2);

    sheet.getRange(writtenRow, 1, 1, 10).setValues([[
      now,
      name,
      phone,
      specialty,
      start,
      comment || "Нет",
      "Записан",
      source,
      candidateId,
      objectName
    ]]);

    // Форматирование записи.
    try {
      sheet.getRange(writtenRow, 1)
        .setNumberFormat("dd.MM.yyyy HH:mm");

      sheet.getRange(writtenRow, 3)
        .setNumberFormat("@");

      sheet.getRange(writtenRow, 5)
        .setNumberFormat("dd.MM.yyyy HH:mm");

      sheet.getRange(writtenRow, 9)
        .setHorizontalAlignment("center");

    } catch (formatError) {
      Logger.log("Ошибка форматирования: " + formatError);
    }

    // Запись встречи.
    writtenMeetingRow = Math.max(
      meetingsSheet.getLastRow() + 1, 2
    );

    meetingsSheet.getRange(writtenMeetingRow, 1, 1, 10)
      .setValues([[
        meetingId,
        candidateId,
        name,
        phone,
        specialty,
        start,
        "Запланирована",
        event.getId(),
        now,
        ""
      ]]);

    // Запись действия в историю.
    writtenHistoryRow = Math.max(
      historySheet.getLastRow() + 1, 2
    );

    historySheet.getRange(writtenHistoryRow, 1, 1, 8)
      .setValues([[
        "HIST-" + Utilities.getUuid().substring(0, 8).toUpperCase(),
        candidateId,
        now,
        "Создание кандидата",
        "",
        "Записан",
        "Кандидат добавлен и назначена встреча",
        "CRM"
      ]]);

    SpreadsheetApp.flush();

    // Сортировка встреч.
    try {
      sortMeetingsByDate();
    } catch (sortError) {
      Logger.log("Ошибка автоматической сортировки: " + sortError);
    }

    // Уведомление в MAX через GitHub Actions.
    let notificationSent = false;
    let notificationMessage = "";

    try {
      sendCandidateToGitHub_({
        candidateId: candidateId,
        name: name,
        phone: phone,
        specialty: specialty,
        meeting: Utilities.formatDate(
          start,
          MOSCOW_TZ,
          "dd.MM.yyyy HH:mm"
        ),
        source: source,
        comment: comment || "Нет",
        object: objectName
      });

      notificationSent = true;
      notificationMessage = " Уведомление отправлено в MAX.";

    } catch (notificationError) {
      Logger.log("Ошибка уведомления MAX: " + notificationError);

      notificationMessage =
        " Кандидат записан, но уведомление MAX не отправлено.";
    }

    return {
      success: true,
      notificationSent: notificationSent,
      candidateId: candidateId,
      meetingId: meetingId,
      message: "Кандидат успешно добавлен!" + notificationMessage
    };

  } catch (error) {

    // Откат записей при критической ошибке.
    if (historySheet && writtenHistoryRow) {
      try {
        historySheet.deleteRow(writtenHistoryRow);
      } catch (e) {
        Logger.log("Ошибка отката истории: " + e);
      }
    }

    if (meetingsSheet && writtenMeetingRow) {
      try {
        meetingsSheet.deleteRow(writtenMeetingRow);
      } catch (e) {
        Logger.log("Ошибка отката встречи: " + e);
      }
    }

    if (sheet && writtenRow) {
      try {
        sheet.deleteRow(writtenRow);
      } catch (e) {
        Logger.log("Ошибка отката кандидата: " + e);
      }
    }

    if (event) {
      try {
        event.deleteEvent();
      } catch (e) {
        Logger.log("Ошибка удаления события календаря: " + e);
      }
    }

    Logger.log("createCandidate error: " + error.stack);

    return {
      success: false,
      notificationSent: false,
      message: error.message || "Не удалось добавить кандидата."
    };

  } finally {
    if (lock.hasLock()) {
      lock.releaseLock();
    }
  }
}


/************************************************************
 * 4. ПОЛУЧЕНИЕ СПИСКА КАНДИДАТОВ И СТАТИСТИКИ
 ************************************************************/

function getCandidates() {
  try {
    const properties = PropertiesService.getScriptProperties();
    const spreadsheetId = properties.getProperty(SHEET_ID_PROPERTY);

    if (!spreadsheetId) {
      throw new Error("Не указан ID рабочей таблицы.");
    }

    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheet = spreadsheet.getSheetByName(SHEET_NAME);
    const meetingsSheet = spreadsheet.getSheetByName("Встречи");

    if (!sheet) {
      throw new Error('Лист "' + SHEET_NAME + '" не найден.');
    }

    if (!meetingsSheet) {
      throw new Error('Лист "Встречи" не найден.');
    }

    const lastRow = sheet.getLastRow();

    if (lastRow < 2) {
      return {
        success: true,
        candidates: [],
        stats: {
          total: 0,
          newCount: 0,
          today: 0,
          started: 0
        }
      };
    }

    const values = sheet
      .getRange(2, 1, lastRow - 1, 10)
      .getValues();

    const meetingsLastRow = meetingsSheet.getLastRow();

    const meetings = meetingsLastRow >= 2
      ? meetingsSheet.getRange(2, 1, meetingsLastRow - 1, 10).getValues()
      : [];

    // Последняя актуальная встреча каждого кандидата.
    const meetingsMap = {};

    meetings.forEach(function(row) {
      const meetingId = String(row[0] || "");
      const candidateId = String(row[1] || "");
      const meetingStatus = String(row[6] || "").trim();

      if (!candidateId || !meetingId || meetingStatus === "Отменена") {
        return;
      }

      const createdAt = row[8] instanceof Date
        ? row[8].getTime()
        : 0;

      const existing = meetingsMap[candidateId];

      if (!existing || createdAt >= existing.createdAt) {
        meetingsMap[candidateId] = {
          meetingId: meetingId,
          createdAt: createdAt
        };
      }
    });

    const today = Utilities.formatDate(
      new Date(),
      MOSCOW_TZ,
      "yyyy-MM-dd"
    );

    let newCount = 0;
    let todayCount = 0;
    let startedCount = 0;

    const candidates = values.map(function(row, index) {
      const addedDate = row[0];
      const meetingDate = row[4];
      const candidateId = String(row[8] || "");
      const status = String(row[6] || "").trim();
      const object = String(row[9] || "Не указан");

      if (status === "Новый") newCount++;
      if (status === "Работает") startedCount++;

      if (
        meetingDate instanceof Date &&
        !isNaN(meetingDate.getTime())
      ) {
        const meetingDay = Utilities.formatDate(
          meetingDate,
          MOSCOW_TZ,
          "yyyy-MM-dd"
        );

        if (meetingDay === today) todayCount++;
      }

      return {
        row: index + 2,
        candidateId: candidateId,

        meetingId: meetingsMap[candidateId]
          ? meetingsMap[candidateId].meetingId
          : "",

        name: String(row[1] || ""),
        phone: String(row[2] || ""),
        specialty: String(row[3] || ""),

        added: formatDate_(addedDate),
        meeting: formatDate_(meetingDate),

        comment: String(row[5] || ""),
        status: status,
        source: String(row[7] || "Не указан"),
        object: object
      };
    });

    // Новые кандидаты сверху.
    candidates.reverse();

    return {
      success: true,
      candidates: candidates,
      stats: {
        total: values.length,
        newCount: newCount,
        today: todayCount,
        started: startedCount
      }
    };

  } catch (error) {
    Logger.log("getCandidates error: " + error.stack);

    return {
      success: false,
      message: error.message || "Не удалось загрузить кандидатов.",
      candidates: [],
      stats: {
        total: 0,
        newCount: 0,
        today: 0,
        started: 0
      }
    };
  }
}


/************************************************************
 * 5. ИЗМЕНЕНИЕ СТАТУСА КАНДИДАТА
 ************************************************************/

function updateCandidateStatus(row, status) {
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(30000);

    const rowNumber = Number(row);
    const newStatus = String(status || "").trim();

    if (!Number.isInteger(rowNumber) || rowNumber < 2) {
      throw new Error("Некорректный номер строки.");
    }

    if (CANDIDATE_STATUSES.indexOf(newStatus) === -1) {
      throw new Error("Недопустимый статус кандидата.");
    }

    const spreadsheetId = PropertiesService
      .getScriptProperties()
      .getProperty(SHEET_ID_PROPERTY);

    if (!spreadsheetId) {
      throw new Error("Не указан ID рабочей таблицы.");
    }

    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheet = spreadsheet.getSheetByName(SHEET_NAME);
    const historySheet = spreadsheet.getSheetByName("История");

    if (!sheet) {
      throw new Error('Лист "Кандидаты" не найден.');
    }

    if (!historySheet) {
      throw new Error('Лист "История" не найден.');
    }

    if (rowNumber > sheet.getLastRow()) {
      throw new Error("Кандидат не найден.");
    }

    const candidate = sheet.getRange(rowNumber, 1, 1, 9)
      .getValues()[0];

    const candidateId = candidate[8];
    const name = candidate[1];
    const oldStatus = String(candidate[6] || "");

    if (!candidateId) {
      throw new Error("У кандидата отсутствует ID.");
    }

    if (oldStatus === newStatus) {
      return {
        success: true,
        message: "Статус не изменился."
      };
    }

    const now = new Date();

    sheet.getRange(rowNumber, 7).setValue(newStatus);

    const historyRow = Math.max(
      historySheet.getLastRow() + 1, 2
    );

    historySheet.getRange(historyRow, 1, 1, 8).setValues([[
      "HIST-" + Utilities.getUuid().substring(0, 8).toUpperCase(),
      candidateId,
      now,
      "Изменение статуса",
      oldStatus,
      newStatus,
      "Изменён статус кандидата: " + name,
      "CRM"
    ]]);

    SpreadsheetApp.flush();

    return {
      success: true,
      message: "Статус обновлён, изменение сохранено в истории."
    };

  } catch (error) {
    Logger.log("updateCandidateStatus error: " + error.stack);

    return {
      success: false,
      message: error.message || "Не удалось обновить статус."
    };

  } finally {
    if (lock.hasLock()) {
      lock.releaseLock();
    }
  }
}


/************************************************************
 * 6. ПРОВЕРКА ДОСТУПА К GOOGLE CALENDAR
 ************************************************************/

function testCalendarAccess() {
  const calendar = CalendarApp.getCalendarById(CALENDAR_ID);

  if (!calendar) {
    throw new Error("Календарь не найден или нет доступа.");
  }

  Logger.log("Календарь подключен: " + calendar.getName());

  return calendar.getName();
}


/************************************************************
 * 7. ТЕСТОВОЕ СОЗДАНИЕ СОБЫТИЯ
 ************************************************************/

function testCreateEvent() {
  const calendar = CalendarApp.getCalendarById(CALENDAR_ID);

  if (!calendar) {
    throw new Error("Календарь не найден.");
  }

  const start = new Date(Date.now() + 24 * 60 * 60 * 1000);
  const end = new Date(start.getTime() + 30 * 60 * 1000);

  const event = calendar.createEvent(
    "ТЕСТ — CRM STROY PODRYAD",
    start,
    end,
    {
      description: "Тестовое событие. Можно удалить."
    }
  );

  Logger.log("Тестовая встреча создана: " + event.getId());

  return event.getId();
}


/************************************************************
 * 8. ЗАЩИТА ОТ СОЗДАНИЯ НОВОЙ ТАБЛИЦЫ
 ************************************************************/

function setupCandidateSheet() {
  throw new Error(
    "Функция отключена для защиты рабочей таблицы. " +
    "Не создавай новую таблицу поверх действующей CRM."
  );
}


/************************************************************
 * 9. ЗАЩИТА ОТ ПОВТОРНОГО ИМПОРТА
 ************************************************************/

function importOldCandidates() {
  throw new Error(
    "Импорт отключён, чтобы не создавать дубликаты " +
    "в рабочей таблице. Для импорта нужна отдельная проверка."
  );
}


/************************************************************
 * 10. ПРОВЕРКА СОБЫТИЙ КАЛЕНДАРЯ
 ************************************************************/

function checkCalendarEvents() {
  const calendar = CalendarApp.getCalendarById(CALENDAR_ID);

  if (!calendar) {
    throw new Error("Календарь не найден.");
  }

  const start = parseMoscowDateTime_("2026-11-23", "00:00");
  const end = parseMoscowDateTime_("2026-11-24", "00:00");

  const events = calendar.getEvents(start, end);

  Logger.log("Найдено событий: " + events.length);

  events.forEach(function(event) {
    Logger.log(
      event.getTitle() +
      " | " +
      Utilities.formatDate(
        event.getStartTime(),
        MOSCOW_TZ,
        "dd.MM.yyyy HH:mm"
      )
    );
  });

  return events.length;
}


/************************************************************
 * 11. СЛУЖЕБНЫЕ ФУНКЦИИ
 ************************************************************/

function parseMoscowDateTime_(date, time) {
  const dateMatch = /^(\d{4})-(\d{2})-(\d{2})$/.exec(date);
  const timeMatch = /^(\d{2}):(\d{2})$/.exec(time);

  if (!dateMatch) {
    throw new Error("Некорректный формат даты.");
  }

  if (!timeMatch) {
    throw new Error("Некорректный формат времени.");
  }

  const year = Number(dateMatch[1]);
  const month = Number(dateMatch[2]);
  const day = Number(dateMatch[3]);

  const hours = Number(timeMatch[1]);
  const minutes = Number(timeMatch[2]);

  const checkDate = new Date(Date.UTC(year, month - 1, day));

  if (
    checkDate.getUTCFullYear() !== year ||
    checkDate.getUTCMonth() !== month - 1 ||
    checkDate.getUTCDate() !== day
  ) {
    throw new Error("Указана несуществующая дата.");
  }

  if (
    hours < 0 ||
    hours > 23 ||
    minutes < 0 ||
    minutes > 59
  ) {
    throw new Error("Указано некорректное время.");
  }

  if (
    hours < 8 ||
    hours > 19 ||
    (minutes !== 0 && minutes !== 30) ||
    (hours === 19 && minutes !== 0)
  ) {
    throw new Error("Выбери время встречи с 08:00 до 19:00.");
  }

  return new Date(
    date + "T" + time + ":00+03:00"
  );
}


function formatDate_(value) {
  if (
    !(value instanceof Date) ||
    isNaN(value.getTime())
  ) {
    return "";
  }

  return Utilities.formatDate(
    value,
    MOSCOW_TZ,
    "dd.MM.yyyy HH:mm"
  );
}


function jsonResponse_(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}


function testMaxGetUpdates() {
  const token = PropertiesService
    .getScriptProperties()
    .getProperty("MAX_BOT_TOKEN");

  if (!token) {
    throw new Error("Не найден MAX_BOT_TOKEN");
  }

  const response = UrlFetchApp.fetch(
    "https://platform-api2.max.ru/updates?limit=100&timeout=0",
    {
      method: "get",
      headers: {
        Authorization: token
      },
      muteHttpExceptions: true
    }
  );

  Logger.log("HTTP: " + response.getResponseCode());
  Logger.log(response.getContentText());
}


/************************************************************
 * 12. ОТПРАВКА НОВОГО КАНДИДАТА В GITHUB ACTIONS / MAX
 ************************************************************/

function sendCandidateToGitHub_(candidate) {
  const properties = PropertiesService.getScriptProperties();

  const token = properties.getProperty("GITHUB_ACTIONS_TOKEN");

  if (!token) {
    throw new Error("Не найден GITHUB_ACTIONS_TOKEN.");
  }

  const url =
    "https://api.github.com/repos/Alqoire/candidate-booking/actions/workflows/notify-max.yml/dispatches";

  const payload = {
    ref: "main",
    inputs: {
      event_type: "new_candidate",
      full_name: candidate.name,
      phone: candidate.phone,
      specialty: candidate.specialty,
      meeting: candidate.meeting,
      source: candidate.source,
      comment: candidate.comment,
      site_name: String(candidate.object || "Не указан")
    }
  };

  const response = UrlFetchApp.fetch(url, {
    method: "post",
    contentType: "application/json",
    headers: {
      Authorization: "Bearer " + token,
      Accept: "application/vnd.github+json",
      "X-GitHub-Api-Version": "2022-11-28"
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const status = response.getResponseCode();

  if (status !== 204) {
    throw new Error(
      "GitHub вернул HTTP " + status + ": " +
      response.getContentText()
    );
  }

  Logger.log("Уведомление о кандидате отправлено в GitHub Actions.");
}


/************************************************************
 * 13. НАСТРОЙКА ДОПОЛНИТЕЛЬНЫХ ЛИСТОВ CRM
 ************************************************************/

function setupCrmStructure() {
  const properties = PropertiesService.getScriptProperties();
  const spreadsheetId = properties.getProperty(SHEET_ID_PROPERTY);

  if (!spreadsheetId) {
    throw new Error("Не указан ID рабочей таблицы.");
  }

  const spreadsheet = SpreadsheetApp.openById(spreadsheetId);

  const structures = {
    "История": [
      "ID записи",
      "ID кандидата",
      "Дата и время",
      "Действие",
      "Старое значение",
      "Новое значение",
      "Комментарий",
      "Автор"
    ],

    "Встречи": [
      "ID встречи",
      "ID кандидата",
      "ФИО кандидата",
      "Телефон",
      "Специальность",
      "Дата и время встречи",
      "Статус встречи",
      "ID события Google Calendar",
      "Дата создания",
      "Причина изменения"
    ]
  };

  const result = [];

  Object.keys(structures).forEach(function(sheetName) {
    let sheet = spreadsheet.getSheetByName(sheetName);

    if (!sheet) {
      sheet = spreadsheet.insertSheet(sheetName);
    }

    const headers = structures[sheetName];

    const firstRow = sheet
      .getRange(1, 1, 1, headers.length)
      .getValues()[0];

    const isEmpty = firstRow.every(function(value) {
      return value === "" || value === null;
    });

    if (isEmpty) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

      sheet.getRange(1, 1, 1, headers.length)
        .setFontWeight("bold");

      sheet.setFrozenRows(1);

      result.push(sheetName + ": заголовки созданы");
    } else {
      result.push(sheetName + ": данные уже есть, лист не изменён");
    }
  });

  SpreadsheetApp.flush();

  Logger.log(result.join("\n"));

  return result.join("\n");
}


/************************************************************
 * 14. ФОРМАТИРОВАНИЕ ТАБЛИЦ CRM
 ************************************************************/

function formatCrmSheets() {
  const ssId = PropertiesService.getScriptProperties()
    .getProperty(SHEET_ID_PROPERTY);

  if (!ssId) {
    throw new Error("Не найден ID таблицы.");
  }

  const ss = SpreadsheetApp.openById(ssId);

  const navy = "#17243A";
  const gold = "#D6AE59";
  const darkText = "#263449";

  function styleSheet(name, columns, widths, tabColor) {
    const sheet = ss.getSheetByName(name);

    if (!sheet) {
      throw new Error('Не найден лист "' + name + '".');
    }

    sheet.setFrozenRows(1);
    sheet.getBandings().forEach(b => b.remove());

    const range = sheet.getRange(1, 1, 1000, columns);

    const banding = range.applyRowBanding(
      SpreadsheetApp.BandingTheme.LIGHT_GREY
    );

    banding.setHeaderRowColor(navy);
    banding.setFirstRowColor("#FFFFFF");
    banding.setSecondRowColor("#F1F4F8");

    sheet.getRange(1, 1, 1, columns)
      .setBackground(navy)
      .setFontColor(gold)
      .setFontFamily("Arial")
      .setFontSize(10)
      .setFontWeight("bold")
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle")
      .setWrap(true);

    sheet.setRowHeight(1, 44);

    sheet.getRange(2, 1, 999, columns)
      .setFontFamily("Arial")
      .setFontSize(10)
      .setFontColor(darkText)
      .setVerticalAlignment("middle");

    sheet.setTabColor(tabColor);

    widths.forEach((width, index) => {
      sheet.setColumnWidth(index + 1, width);
    });

    if (!sheet.getFilter()) {
      sheet.getRange(1, 1, 1000, columns).createFilter();
    }
  }

  // ИСТОРИЯ
  styleSheet(
    "История",
    8,
    [125, 150, 165, 160, 180, 180, 280, 130],
    "#8B9AAF"
  );

  const history = ss.getSheetByName("История");

  history.getRange("C2:C1000")
    .setNumberFormat("dd.MM.yyyy HH:mm");

  history.getRange("A2:B1000")
    .setHorizontalAlignment("center");

  // ВСТРЕЧИ
  styleSheet(
    "Встречи",
    10,
    [130, 150, 200, 145, 190, 180, 165, 230, 165, 260],
    gold
  );

  const meetings = ss.getSheetByName("Встречи");

  meetings.getRange("F2:F1000")
    .setNumberFormat("dd.MM.yyyy HH:mm");

  meetings.getRange("I2:I1000")
    .setNumberFormat("dd.MM.yyyy HH:mm");

  meetings.getRange("A2:B1000")
    .setHorizontalAlignment("center");

  const statuses = [
    "Запланирована",
    "Состоялась",
    "Перенесена",
    "Отменена",
    "Не явился"
  ];

  meetings.getRange("G2:G1000")
    .setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(statuses, true)
        .setAllowInvalid(false)
        .build()
    );

  const statusColors = [
    ["Запланирована", "#DCEAFE", "#1D4ED8"],
    ["Состоялась", "#DCFCE7", "#166534"],
    ["Перенесена", "#FEF3C7", "#92400E"],
    ["Отменена", "#FEE2E2", "#991B1B"],
    ["Не явился", "#F3E8FF", "#6B21A8"]
  ];

  const statusRange = meetings.getRange("G2:G1000");

  const statusRules = statusColors.map(item =>
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(item[0])
      .setBackground(item[1])
      .setFontColor(item[2])
      .setRanges([statusRange])
      .build()
  );

  const meetingRange = meetings.getRange("A2:J1000");

  const meetingDateRules = [
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(
        '=И($F2<>"";ЦЕЛОЕ($F2)=СЕГОДНЯ())'
      )
      .setBackground("#FEE2E2")
      .setFontColor("#991B1B")
      .setRanges([meetingRange])
      .build(),

    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(
        '=И($F2<>"";ЦЕЛОЕ($F2)=СЕГОДНЯ()+1)'
      )
      .setBackground("#DCFCE7")
      .setFontColor("#166534")
      .setRanges([meetingRange])
      .build(),

    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(
        '=И($F2<>"";ЦЕЛОЕ($F2)<СЕГОДНЯ())'
      )
      .setBackground("#E5E7EB")
      .setFontColor("#6B7280")
      .setRanges([meetingRange])
      .build()
  ];

  meetings.setConditionalFormatRules([
    ...meetingDateRules,
    ...statusRules
  ]);

  // КАНДИДАТЫ
  const candidates = ss.getSheetByName("Кандидаты");

  if (!candidates) {
    throw new Error('Не найден лист "Кандидаты".');
  }

  const candidateRange = candidates.getRange("A2:J1000");

  const candidateDateRules = [
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(
        '=И($E2<>"";ЦЕЛОЕ($E2)=СЕГОДНЯ())'
      )
      .setBackground("#FEE2E2")
      .setFontColor("#991B1B")
      .setRanges([candidateRange])
      .build(),

    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(
        '=И($E2<>"";ЦЕЛОЕ($E2)=СЕГОДНЯ()+1)'
      )
      .setBackground("#DCFCE7")
      .setFontColor("#166534")
      .setRanges([candidateRange])
      .build(),

    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(
        '=И($E2<>"";ЦЕЛОЕ($E2)<СЕГОДНЯ())'
      )
      .setBackground("#E5E7EB")
      .setFontColor("#6B7280")
      .setRanges([candidateRange])
      .build()
  ];

  candidates.setConditionalFormatRules(candidateDateRules);

  SpreadsheetApp.flush();

  Logger.log("Оформление CRM успешно применено!");
}


/************************************************************
 * 15. МИГРАЦИЯ ID КАНДИДАТОВ
 ************************************************************/

function migrateCandidateIds() {
  const ssId = PropertiesService.getScriptProperties()
    .getProperty(SHEET_ID_PROPERTY);

  if (!ssId) {
    throw new Error("Не найден ID рабочей таблицы.");
  }

  const ss = SpreadsheetApp.openById(ssId);
  const sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    throw new Error("Лист Кандидаты не найден.");
  }

  const header = sheet.getRange("I1").getValue();

  if (header && header !== "ID кандидата") {
    throw new Error(
      "Столбец I уже занят другим заголовком. Ничего не изменено."
    );
  }

  sheet.getRange("I1").setValue("ID кандидата");

  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    SpreadsheetApp.flush();
    Logger.log("Столбец ID создан. Кандидатов пока нет.");
    return;
  }

  const idRange = sheet.getRange(2, 9, lastRow - 1, 1);
  const ids = idRange.getValues();

  let created = 0;

  for (let i = 0; i < ids.length; i++) {
    if (!ids[i][0]) {
      ids[i][0] = "CAND-" + Utilities.getUuid()
        .substring(0, 8).toUpperCase();

      created++;
    }
  }

  idRange.setValues(ids);

  sheet.setColumnWidth(9, 155);

  sheet.getRange(1, 9)
    .setBackground("#17243A")
    .setFontColor("#D6AE59")
    .setFontWeight("bold")
    .setHorizontalAlignment("center");

  idRange.setFontFamily("Arial")
    .setFontSize(10)
    .setHorizontalAlignment("center");

  SpreadsheetApp.flush();

  Logger.log("Готово! Создано ID: " + created);
}


/************************************************************
 * 16. СОРТИРОВКА ВСТРЕЧ
 ************************************************************/

function sortMeetingsByDate() {
  const ssId = PropertiesService.getScriptProperties()
    .getProperty(SHEET_ID_PROPERTY);

  if (!ssId) {
    throw new Error("Не найден ID таблицы.");
  }

  const ss = SpreadsheetApp.openById(ssId);
  const sheet = ss.getSheetByName("Встречи");

  if (!sheet) {
    throw new Error('Не найден лист "Встречи".');
  }

  const lastRow = sheet.getLastRow();

  if (lastRow < 3) {
    Logger.log("Недостаточно встреч для сортировки.");
    return;
  }

  sheet.getRange(2, 1, lastRow - 1, 10)
    .sort([
      { column: 6, ascending: true }
    ]);

  SpreadsheetApp.flush();

  Logger.log("Встречи отсортированы по дате!");
}


/************************************************************
 * 17. ПЕРЕНОС СТАРЫХ КАНДИДАТОВ В РЕЕСТР ВСТРЕЧ
 ************************************************************/

function migrateOldCandidatesToMeetings() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    migrateCandidateIds();

    const ssId = PropertiesService.getScriptProperties()
      .getProperty(SHEET_ID_PROPERTY);

    if (!ssId) {
      throw new Error("Не найден ID таблицы.");
    }

    const ss = SpreadsheetApp.openById(ssId);
    const candidates = ss.getSheetByName(SHEET_NAME);
    const meetings = ss.getSheetByName("Встречи");
    const history = ss.getSheetByName("История");

    if (!candidates || !meetings || !history) {
      throw new Error("Не найден один из необходимых листов.");
    }

    const lastRow = candidates.getLastRow();

    if (lastRow < 2) {
      Logger.log("Старых кандидатов для переноса нет.");
      return;
    }

    const data = candidates.getRange(2, 1, lastRow - 1, 9)
      .getValues();

    const meetingLastRow = meetings.getLastRow();

    const existingIds = new Set();

    if (meetingLastRow >= 2) {
      meetings.getRange(2, 2, meetingLastRow - 1, 1)
        .getValues()
        .forEach(row => {
          if (row[0]) {
            existingIds.add(String(row[0]).trim());
          }
        });
    }

    const meetingRows = [];
    const historyRows = [];
    const idUpdates = [];

    let skipped = 0;

    function parseOldMeetingDate(value) {
      if (value instanceof Date && !isNaN(value.getTime())) {
        return value;
      }

      if (!value) return "";

      const str = String(value).trim();

      const match = str.match(
        /^(\d{2})\.(\d{2})\.(\d{4})(?:\s+(\d{1,2}):(\d{2}))?$/
      );

      if (match) {
        return new Date(
          Number(match[3]),
          Number(match[2]) - 1,
          Number(match[1]),
          Number(match[4] || 0),
          Number(match[5] || 0)
        );
      }

      const parsed = new Date(str);

      return isNaN(parsed.getTime()) ? "" : parsed;
    }

    data.forEach((row, index) => {
      const [
        createdAt,
        name,
        phone,
        specialty,
        meetingDate,
        comment,
        candidateStatus,
        source,
        candidateIdValue
      ] = row;

      if (!name && !phone && !specialty) return;

      let candidateId = String(candidateIdValue || "").trim();

      if (!candidateId) {
        candidateId = "CAND-" +
          Utilities.getUuid().substring(0, 8).toUpperCase();

        idUpdates.push({
          row: index + 2,
          id: candidateId
        });
      }

      if (existingIds.has(candidateId)) {
        skipped++;
        return;
      }

      const meetingId = "MEET-" +
        Utilities.getUuid().substring(0, 8).toUpperCase();

      const parsedDate = parseOldMeetingDate(meetingDate);

      const meetingStatus = parsedDate
        ? "Запланирована"
        : "Не назначена";

      const now = new Date();

      meetingRows.push([
        meetingId,
        candidateId,
        name || "",
        phone || "",
        specialty || "",
        parsedDate,
        meetingStatus,
        "",
        createdAt instanceof Date ? createdAt : now,
        "Перенос старого кандидата"
      ]);

      historyRows.push([
        "HIST-" + Utilities.getUuid().substring(0, 8).toUpperCase(),
        candidateId,
        now,
        "Перенос в реестр встреч",
        "",
        meetingStatus,
        "Старая запись перенесена в новый реестр",
        "CRM"
      ]);

      existingIds.add(candidateId);
    });

    idUpdates.forEach(item => {
      candidates.getRange(item.row, 9).setValue(item.id);
    });

    if (meetingRows.length) {
      const startRow = Math.max(meetings.getLastRow() + 1, 2);

      meetings.getRange(startRow, 1, meetingRows.length, 10)
        .setValues(meetingRows);

      meetings.getRange(startRow, 6, meetingRows.length, 1)
        .setNumberFormat("dd.MM.yyyy HH:mm");

      meetings.getRange(startRow, 9, meetingRows.length, 1)
        .setNumberFormat("dd.MM.yyyy HH:mm");

      const historyStart = Math.max(history.getLastRow() + 1, 2);

      history.getRange(historyStart, 1, historyRows.length, 8)
        .setValues(historyRows);
    }

    SpreadsheetApp.flush();

    sortMeetingsByDate();

    Logger.log("МИГРАЦИЯ ЗАВЕРШЕНА!");
    Logger.log("Перенесено кандидатов: " + meetingRows.length);
    Logger.log("Пропущено дублей: " + skipped);
    Logger.log("Создано недостающих ID: " + idUpdates.length);

  } finally {
    lock.releaseLock();
  }
}


/************************************************************
 * 18. ПЕРЕНОС ВСТРЕЧИ
 ************************************************************/


function rescheduleMeeting(meetingId, newDate, newTime, reason) {
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(30000);

    meetingId = String(meetingId || "").trim();
    reason = String(reason || "").trim();

    if (!meetingId) {
      throw new Error("Не указан ID встречи.");
    }

    if (!reason) {
      throw new Error("Укажи причину переноса.");
    }

    const start = parseMoscowDateTime_(newDate, newTime);
    const end = new Date(start.getTime() + 30 * 60 * 1000);

    const ssId = PropertiesService.getScriptProperties()
      .getProperty(SHEET_ID_PROPERTY);

    if (!ssId) {
      throw new Error("Не найдена рабочая таблица.");
    }

    const ss = SpreadsheetApp.openById(ssId);
    const meetings = ss.getSheetByName("Встречи");
    const candidates = ss.getSheetByName(SHEET_NAME);
    const history = ss.getSheetByName("История");

    if (!meetings || !candidates || !history) {
      throw new Error("Не найдены необходимые листы.");
    }

    const lastRow = meetings.getLastRow();

    if (lastRow < 2) {
      throw new Error("Встречи не найдены.");
    }

    const rows = meetings.getRange(2, 1, lastRow - 1, 10)
      .getValues();

    const index = rows.findIndex(row =>
      String(row[0]).trim() === meetingId
    );

    if (index === -1) {
      throw new Error("Встреча не найдена.");
    }

    const rowNumber = index + 2;
    const meeting = rows[index];

    const candidateId = String(meeting[1]).trim();
    const name = meeting[2];
    const oldDate = meeting[5];
    const eventId = String(meeting[7] || "").trim();

    if (meeting[6] === "Отменена") {
      throw new Error("Нельзя перенести отменённую встречу.");
    }

    const calendar = CalendarApp.getCalendarById(CALENDAR_ID);

    if (!calendar) {
      throw new Error("Календарь недоступен.");
    }

    let event = eventId
      ? calendar.getEventById(eventId)
      : null;

    if (!event) {
      event = calendar.createEvent(
        "Встреча с кандидатом: " + name,
        start,
        end,
        {
          description:
            "ID кандидата: " + candidateId +
            "\nID встречи: " + meetingId
        }
      );

      event.removeAllReminders();
      event.addPopupReminder(60);

    } else {
      event.setTime(start, end);
    }

    const now = new Date();

    // Обновляем реестр встреч.
    meetings.getRange(rowNumber, 6).setValue(start);
    meetings.getRange(rowNumber, 7).setValue("Перенесена");
    meetings.getRange(rowNumber, 8).setValue(event.getId());
    meetings.getRange(rowNumber, 9).setValue(now);
    meetings.getRange(rowNumber, 10).setValue(reason);

    // Обновляем дату кандидата.
    const candidateLastRow = candidates.getLastRow();

    let candidateData = {};

    if (candidateLastRow >= 2) {
      const candidateRows = candidates
        .getRange(2, 1, candidateLastRow - 1, 10)
        .getValues();

      const candidateIndex = candidateRows.findIndex(row =>
        String(row[8]).trim() === candidateId
      );

      if (candidateIndex !== -1) {
        const candidateRow = candidateRows[candidateIndex];

        candidates.getRange(candidateIndex + 2, 5).setValue(start);

        candidateData = {
          name: candidateRow[1],
          phone: candidateRow[2],
          specialty: candidateRow[3],
          source: candidateRow[7],
          comment: candidateRow[5],
          object: candidateRow[9] || "Не указан"
        };
      }
    }

    // Записываем действие в историю.
    history.appendRow([
      "HIST-" + Utilities.getUuid().substring(0, 8).toUpperCase(),
      candidateId,
      now,
      "Перенос встречи",
      oldDate,
      start,
      reason,
      "CRM"
    ]);

    SpreadsheetApp.flush();

    sortMeetingsByDate();

    // Отправляем уведомление о переносе в MAX через GitHub Actions.
    let notificationSent = false;
    let notificationErrorMessage = "";

    try {
      const notificationData = {
        name: candidateData.name || name,
        phone: candidateData.phone || meeting[3],
        specialty: candidateData.specialty || meeting[4],
        source: candidateData.source || "",
        comment: candidateData.comment || "",
        object: candidateData.object || "Не указан",

        oldMeeting: oldDate instanceof Date
          ? Utilities.formatDate(
              oldDate,
              MOSCOW_TZ,
              "dd.MM.yyyy HH:mm"
            )
          : String(oldDate || "Не указано"),

        meeting: Utilities.formatDate(
          start,
          MOSCOW_TZ,
          "dd.MM.yyyy HH:mm"
        ),

        reason: reason
      };

      Logger.log("Начинаем отправку уведомления о переносе.");
      Logger.log("Тип события: rescheduled");
      Logger.log("Данные уведомления: " + JSON.stringify(notificationData));

      sendMeetingNotificationToGitHub_(
        "rescheduled",
        notificationData
      );

      notificationSent = true;

      Logger.log("GitHub принял запрос на уведомление о переносе.");

    } catch (notificationError) {
      notificationErrorMessage =
        notificationError.message || String(notificationError);

      Logger.log("ОШИБКА ОТПРАВКИ MAX: " + notificationErrorMessage);
      Logger.log(notificationError.stack || "Stack отсутствует");
    }

    return {
      success: true,
      notificationSent: notificationSent,
      message: notificationSent
        ? "Встреча перенесена. GitHub принял запрос на уведомление."
        : "Встреча перенесена, но уведомление не отправлено: " +
          notificationErrorMessage,
      meetingId: meetingId,
      notificationError: notificationErrorMessage
    };

  } catch (error) {
    Logger.log("rescheduleMeeting error: " + error.stack);

    return {
      success: false,
      notificationSent: false,
      message: error.message || "Не удалось перенести встречу."
    };

  } finally {
    if (lock.hasLock()) {
      lock.releaseLock();
    }
  }
}



/************************************************************
 * 19. УВЕДОМЛЕНИЯ О ПЕРЕНОСЕ И ДРУГИХ СОБЫТИЯХ
 ************************************************************/

/**
 * GitHub workflow_dispatch не принимает пустые строки в inputs → HTTP 422.
 * Все значения нормализуем: trim + fallback, без "".
 */

function sanitizeWorkflowInput_(value, fallback, maxLength) {
  const text = String(value == null ? "" : value).trim();
  if (!text) return fallback || "-";

  const limit = Number(maxLength) || 500;
  return text.length > limit
    ? text.substring(0, limit)
    : text;
}

function sendMeetingNotificationToGitHub_(eventType, candidate) {
  const token = PropertiesService.getScriptProperties()
    .getProperty("GITHUB_ACTIONS_TOKEN");

  if (!token) {
    throw new Error("Не найден GITHUB_ACTIONS_TOKEN.");
  }

  candidate = candidate || {};

  const url =
    "https://api.github.com/repos/Alqoire/candidate-booking/actions/workflows/notify-max.yml/dispatches";

  const payload = {
    ref: "main",
    inputs: {
      event_type: sanitizeWorkflowInput_(eventType, "unknown"),
      full_name: sanitizeWorkflowInput_(candidate.name, "Без ФИО"),
      phone: sanitizeWorkflowInput_(candidate.phone, "-"),
      specialty: sanitizeWorkflowInput_(candidate.specialty, "-"),
      meeting: sanitizeWorkflowInput_(candidate.meeting, "-"),
      old_meeting: sanitizeWorkflowInput_(candidate.oldMeeting, "-"),
      reason: sanitizeWorkflowInput_(candidate.reason, "-"),
      source: sanitizeWorkflowInput_(candidate.source, "-"),
      comment: sanitizeWorkflowInput_(candidate.comment, "-", 4000),
      site_name: sanitizeWorkflowInput_(candidate.object, "Не указан")
    }
  };

  Logger.log("Отправка уведомления в GitHub.");
  Logger.log("Тип события: " + eventType);
  Logger.log("Кандидат: " + candidate.name);
  Logger.log("Встреча: " + candidate.meeting);

  const response = UrlFetchApp.fetch(url, {
    method: "post",
    contentType: "application/json",
    headers: {
      Authorization: "Bearer " + token,
      Accept: "application/vnd.github+json",
      "X-GitHub-Api-Version": "2022-11-28"
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const status = response.getResponseCode();
  const responseText = response.getContentText();

  Logger.log("GitHub HTTP: " + status);
  Logger.log("Ответ GitHub: " + responseText);

  if (status !== 204) {
    throw new Error(
      "GitHub вернул HTTP " + status + ": " + responseText
    );
  }

  Logger.log("GitHub принял запрос на уведомление.");

  return true;
}


function sendMeetingNotificationToGitHub_(eventType, candidate) {
  const token = PropertiesService.getScriptProperties()
    .getProperty("GITHUB_ACTIONS_TOKEN");

  if (!token) {
    throw new Error("Не найден GITHUB_ACTIONS_TOKEN.");
  }

  candidate = candidate || {};

  const url =
    "https://api.github.com/repos/Alqoire/candidate-booking/actions/workflows/notify-max.yml/dispatches";

  const payload = {
    ref: "main",
    inputs: {
      event_type: sanitizeWorkflowInput_(eventType, "unknown"),
      full_name: sanitizeWorkflowInput_(candidate.name, "Без ФИО"),
      phone: sanitizeWorkflowInput_(candidate.phone, "-"),
      specialty: sanitizeWorkflowInput_(candidate.specialty, "-"),
      meeting: sanitizeWorkflowInput_(candidate.meeting, "-"),
      old_meeting: sanitizeWorkflowInput_(candidate.oldMeeting, "-"),
      reason: sanitizeWorkflowInput_(candidate.reason, "-"),
      source: sanitizeWorkflowInput_(candidate.source, "-"),
      comment: sanitizeWorkflowInput_(candidate.comment, "-"),
      site_name: sanitizeWorkflowInput_(candidate.object, "Не указан")
    }
  };

  Logger.log("Отправка уведомления в GitHub.");
  Logger.log("Тип события: " + eventType);
  Logger.log("Кандидат: " + candidate.name);
  Logger.log("Новая встреча: " + candidate.meeting);
  Logger.log("Старая встреча: " + candidate.oldMeeting);

  const response = UrlFetchApp.fetch(url, {
    method: "post",
    contentType: "application/json",
    headers: {
      Authorization: "Bearer " + token,
      Accept: "application/vnd.github+json",
      "X-GitHub-Api-Version": "2022-11-28"
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const status = response.getResponseCode();
  const responseText = response.getContentText();

  Logger.log("GitHub HTTP: " + status);
  Logger.log("Ответ GitHub: " + responseText);

  if (status !== 204) {
    throw new Error(
      "GitHub вернул HTTP " + status + ": " + responseText
    );
  }

  Logger.log("GitHub принял запрос на уведомление.");

  return true;
}


/************************************************************
 * 20. ПОЛУЧЕНИЕ СПИСКА ОБЪЕКТОВ
 ************************************************************/

function getObjects() {
  const ssId = PropertiesService.getScriptProperties()
    .getProperty(SHEET_ID_PROPERTY);

  if (!ssId) {
    throw new Error("Не найден ID рабочей таблицы.");
  }

  const ss = SpreadsheetApp.openById(ssId);
  const sheet = ss.getSheetByName("Объекты");

  if (!sheet) {
    throw new Error("Лист «Объекты» не найден.");
  }

  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return [];
  }

  return sheet.getRange(2, 1, lastRow - 1, 1)
    .getValues()
    .flat()
    .map(value => String(value).trim())
    .filter(Boolean);
}


/************************************************************
 * 21. ДОБАВЛЕНИЕ НОВОГО ОБЪЕКТА
 ************************************************************/

function addObject(objectName) {
  objectName = String(objectName || "").trim();

  if (!objectName) {
    throw new Error("Введите название объекта.");
  }

  const ssId = PropertiesService.getScriptProperties()
    .getProperty(SHEET_ID_PROPERTY);

  if (!ssId) {
    throw new Error("Не найден ID рабочей таблицы.");
  }

  const ss = SpreadsheetApp.openById(ssId);
  const sheet = ss.getSheetByName("Объекты");

  if (!sheet) {
    throw new Error("Лист «Объекты» не найден.");
  }

  const existing = getObjects();

  const duplicate = existing.some(name =>
    name.toLowerCase() === objectName.toLowerCase()
  );

  if (duplicate) {
    return {
      success: false,
      message: "Такой объект уже существует."
    };
  }

  sheet.appendRow([objectName, new Date()]);

  return {
    success: true,
    message: "Объект добавлен.",
    objects: getObjects()
  };
}


function cancelMeeting(meetingId, reason) {
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(30000);

    meetingId = String(meetingId || "").trim();
    reason = String(reason || "").trim();

    if (!meetingId) {
      throw new Error("Не указан ID встречи.");
    }

    if (!reason) {
      throw new Error("Укажи причину отмены.");
    }

    const ssId = PropertiesService.getScriptProperties()
      .getProperty(SHEET_ID_PROPERTY);

    if (!ssId) {
      throw new Error("Не найдена рабочая таблица.");
    }

    const ss = SpreadsheetApp.openById(ssId);
    const meetings = ss.getSheetByName("Встречи");
    const candidates = ss.getSheetByName(SHEET_NAME);
    const history = ss.getSheetByName("История");

    if (!meetings || !candidates || !history) {
      throw new Error("Не найдены необходимые листы.");
    }

    const lastRow = meetings.getLastRow();

    if (lastRow < 2) {
      throw new Error("Встречи не найдены.");
    }

    const rows = meetings
      .getRange(2, 1, lastRow - 1, 10)
      .getValues();

    const index = rows.findIndex(row =>
      String(row[0]).trim() === meetingId
    );

    if (index === -1) {
      throw new Error("Встреча не найдена.");
    }

    const rowNumber = index + 2;
    const meeting = rows[index];

    const candidateId = String(meeting[1] || "").trim();
    const name = meeting[2];
    const oldDate = meeting[5];
    const eventId = String(meeting[7] || "").trim();
    const status = String(meeting[6] || "").trim();

    if (status === "Отменена") {
      throw new Error("Эта встреча уже отменена.");
    }

    const calendar = CalendarApp.getCalendarById(CALENDAR_ID);

    if (!calendar) {
      throw new Error("Календарь недоступен.");
    }

    // Удаляем событие из Google Calendar.
    if (eventId) {
      const event = calendar.getEventById(eventId);

      if (event) {
        event.deleteEvent();
      }
    }

    const now = new Date();

    // Обновляем статус встречи.
    meetings.getRange(rowNumber, 7).setValue("Отменена");
    meetings.getRange(rowNumber, 9).setValue(now);
    meetings.getRange(rowNumber, 10).setValue(reason);

    // Находим кандидата и очищаем дату его встречи.
    const candidateLastRow = candidates.getLastRow();
    let candidateData = {};

    if (candidateLastRow >= 2) {
      const candidateRows = candidates
        .getRange(2, 1, candidateLastRow - 1, 10)
        .getValues();

      const candidateIndex = candidateRows.findIndex(row =>
        String(row[8] || "").trim() === candidateId
      );

      if (candidateIndex !== -1) {
        const candidateRow = candidateRows[candidateIndex];

        candidateData = {
          name: candidateRow[1],
          phone: candidateRow[2],
          specialty: candidateRow[3],
          source: candidateRow[7],
          comment: candidateRow[5],
          object: candidateRow[9] || "Не указан"
        };

        candidates
          .getRange(candidateIndex + 2, 5)
          .clearContent();
      }
    }

    // Записываем отмену в историю.
    history.appendRow([
      "HIST-" + Utilities.getUuid().substring(0, 8).toUpperCase(),
      candidateId,
      now,
      "Отмена встречи",
      oldDate,
      "",
      reason,
      "CRM"
    ]);

    SpreadsheetApp.flush();

    sortMeetingsByDate();

    // Отправляем уведомление в MAX.
    let notificationSent = false;
    let notificationErrorMessage = "";

    try {
      const notificationData = {
        name: candidateData.name || name,
        phone: candidateData.phone || meeting[3],
        specialty: candidateData.specialty || meeting[4],
        source: candidateData.source || "",
        comment: candidateData.comment || "",
        object: candidateData.object || "Не указан",

        oldMeeting: oldDate instanceof Date
          ? Utilities.formatDate(
              oldDate,
              MOSCOW_TZ,
              "dd.MM.yyyy HH:mm"
            )
          : String(oldDate || "Не указано"),

        meeting: "",
        reason: reason
      };

      sendMeetingNotificationToGitHub_(
        "cancelled",
        notificationData
      );

      notificationSent = true;

      Logger.log("Уведомление об отмене принято GitHub.");

    } catch (notificationError) {
      notificationErrorMessage =
        notificationError.message || String(notificationError);

      Logger.log(
        "ОШИБКА ОТПРАВКИ MAX: " + notificationErrorMessage
      );
    }

    return {
      success: true,
      notificationSent: notificationSent,
      message: notificationSent
        ? "Встреча отменена. Уведомление отправлено."
        : "Встреча отменена, но уведомление не отправлено: " +
          notificationErrorMessage,
      meetingId: meetingId,
      notificationError: notificationErrorMessage
    };

  } catch (error) {
    Logger.log("cancelMeeting error: " + error.stack);

    return {
      success: false,
      notificationSent: false,
      message: error.message || "Не удалось отменить встречу."
    };

  } finally {
    if (lock.hasLock()) {
      lock.releaseLock();
    }
  }
}


/**
 * АВТОМАТИЧЕСКИЕ НАПОМИНАНИЯ В MAX ЗА 2 ЧАСА
 * Проверка встреч каждые 5 минут.
 */



/**
 * АВТОМАТИЧЕСКИЕ ГРУППОВЫЕ НАПОМИНАНИЯ В MAX ЗА 2 ЧАСА
 * Проверка встреч каждые 5 минут.
 * Все кандидаты с одинаковой датой и временем — в одном сообщении.
 */

function sendTwoHourReminders() {
  const lock = LockService.getScriptLock();

  if (!lock.tryLock(1000)) return;

  try {
    const properties = PropertiesService.getScriptProperties();
    const spreadsheetId = properties.getProperty("CANDIDATE_SHEET_ID");

    if (!spreadsheetId) {
      throw new Error("Не найден ID рабочей таблицы.");
    }

    const ss = SpreadsheetApp.openById(spreadsheetId);
    const meetingsSheet = ss.getSheetByName("Встречи");
    const candidatesSheet = ss.getSheetByName("Кандидаты");

    if (!meetingsSheet || !candidatesSheet) {
      throw new Error("Не найден лист Встречи или Кандидаты.");
    }

    const meetings = meetingsSheet.getDataRange().getValues();
    const candidates = candidatesSheet.getDataRange().getValues();

    if (meetings.length < 2 || candidates.length < 2) return;

    const now = Date.now();
    const twoHours = 2 * 60 * 60 * 1000;

    // Индекс кандидатов по ID из колонки I.
    const candidatesById = {};

    for (let i = 1; i < candidates.length; i++) {
      const candidateId = String(candidates[i][8] || "").trim();

      if (candidateId) {
        candidatesById[candidateId] = candidates[i];
      }
    }

    // Статусы, для которых напоминания не отправляются.
    const skippedStatuses = [
      "Приехал на объект",
      "На испытательном сроке",
      "Работает",
      "Отказ",
      "Не приехал",
      "Не прошёл испытательный срок"
    ];

    // Группировка по точной дате и времени встречи.
    const groups = {};

    for (let i = 1; i < meetings.length; i++) {
      const row = meetings[i];

      const meetingId = String(row[0] || "").trim();
      const candidateId = String(row[1] || "").trim();
      const meetingDate = row[5];
      const meetingStatus = String(row[6] || "").trim();

      if (!meetingId || !candidateId || !meetingDate) continue;

      if (!["Запланирована", "Перенесена"].includes(meetingStatus)) {
        continue;
      }

      const candidateRow = candidatesById[candidateId];

      if (!candidateRow) continue;

      const candidateStatus = String(candidateRow[6] || "").trim();

      if (skippedStatuses.includes(candidateStatus)) continue;

      let meetingTime;

      if (meetingDate instanceof Date) {
        meetingTime = meetingDate.getTime();
      } else {
        meetingTime = parseReminderDate_(meetingDate);
      }

      if (!meetingTime) continue;

      const timeLeft = meetingTime - now;

      // Берём только встречи в ближайшие два часа.
      if (timeLeft <= 0 || timeLeft > twoHours) continue;

      // Уникальный ключ каждой встречи для защиты от повторов.
      const reminderKey =
        "REMINDER_2H_" + meetingId + "_" + meetingTime;

      if (properties.getProperty(reminderKey)) continue;

      // Все встречи с одинаковым timestamp попадают в одну группу.
      const groupKey = String(meetingTime);

      if (!groups[groupKey]) {
        groups[groupKey] = {
          meetingTime: meetingTime,
          candidates: []
        };
      }

      groups[groupKey].candidates.push({
        meetingId: meetingId,
        reminderKey: reminderKey,
        name: String(candidateRow[1] || "Без имени"),
        phone: String(candidateRow[2] || "-"),
        specialty: String(candidateRow[3] || "-"),
        object: String(candidateRow[9] || "Не указан")
      });
    }

    // Отправляем одно уведомление на каждое время встречи.
    const groupKeys = Object.keys(groups).sort(
      (a, b) => Number(a) - Number(b)
    );

    groupKeys.forEach(function(groupKey) {
      const group = groups[groupKey];

      const meetingText = Utilities.formatDate(
        new Date(group.meetingTime),
        MOSCOW_TZ,
        "dd.MM.yyyy HH:mm"
      );

      const candidateList = group.candidates.map(
        function(candidate, index) {
          return (
            (index + 1) + ". " + candidate.name +
            "\n   Специальность: " + candidate.specialty +
            "\n   Телефон: " + candidate.phone 
          );
        }
      ).join("\n\n");

      try {
        // Отправляем одну общую заявку в GitHub.
        sendMeetingNotificationToGitHub_("reminder_group", {
          name: "Групповое напоминание",
          phone: "-",
          specialty: "Несколько кандидатов",
          meeting: meetingText,
          source: "",
          comment: candidateList,
          object: "Несколько кандидатов"
        });

        // Отмечаем каждую встречу только после успешной отправки.
        group.candidates.forEach(function(candidate) {
          properties.setProperty(
            candidate.reminderKey,
            new Date().toISOString()
          );
        });

        Logger.log(
          "Групповое напоминание отправлено. " +
          "Кандидатов: " + group.candidates.length +
          ". Время: " + meetingText
        );

      } catch (error) {
        Logger.log(
          "Ошибка группового напоминания на " +
          meetingText + ": " + error.message
        );
      }
    });

  } finally {
    lock.releaseLock();
  }
}


/**
 * Преобразование даты из таблицы, если она хранится текстом.
 * Поддерживает формат dd.MM.yyyy HH:mm
 */
function parseReminderDate_(value) {
  if (value instanceof Date) return value.getTime();

  const match = String(value).match(
    /^(\d{2})\.(\d{2})\.(\d{4})\s+(\d{2}):(\d{2})$/
  );

  if (!match) return null;

  const day = Number(match[1]);
  const month = Number(match[2]) - 1;
  const year = Number(match[3]);
  const hours = Number(match[4]);
  const minutes = Number(match[5]);

  // Формируем дату в московском часовом поясе
  const iso =
    Utilities.formatString(
      "%04d-%02d-%02dT%02d:%02d:00+03:00",
      year, month + 1, day, hours, minutes
    );

  const timestamp = new Date(iso).getTime();

  return isNaN(timestamp) ? null : timestamp;
}

/**
 * Установка автоматической проверки каждые 5 минут.
 * Повторный запуск не создаёт дублирующих триггеров.
 */
function setupTwoHourReminderTrigger() {
  const handler = "sendTwoHourReminders";

  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === handler) {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger(handler)
    .timeBased()
    .everyMinutes(5)
    .create();

  Logger.log("Триггер напоминаний установлен: каждые 5 минут.");
}


/************************************************************
 * 22. ПОЛУЧЕНИЕ ВСТРЕЧ ДЛЯ КАЛЕНДАРЯ CRM
 ************************************************************/

function getCalendarMeetings(startDate, endDate) {
  try {
    startDate = String(startDate || "").trim();
    endDate = String(endDate || "").trim();

    // Проверяем формат дат.
    const datePattern = /^\d{4}-\d{2}-\d{2}$/;

    if (!datePattern.test(startDate) || !datePattern.test(endDate)) {
      throw new Error("Некорректный формат периода календаря.");
    }

    // Границы периода по московскому времени.
    const start = new Date(startDate + "T00:00:00+03:00");
    const end = new Date(endDate + "T00:00:00+03:00");

    if (
      isNaN(start.getTime()) ||
      isNaN(end.getTime()) ||
      end <= start
    ) {
      throw new Error("Некорректный период календаря.");
    }

    const ssId = PropertiesService
      .getScriptProperties()
      .getProperty(SHEET_ID_PROPERTY);

    if (!ssId) {
      throw new Error("Не найдена рабочая таблица.");
    }

    const ss = SpreadsheetApp.openById(ssId);

    const meetingsSheet = ss.getSheetByName("Встречи");
    const candidatesSheet = ss.getSheetByName(SHEET_NAME);

    if (!meetingsSheet || !candidatesSheet) {
      throw new Error("Не найдены листы встреч или кандидатов.");
    }

    const meetingsLastRow = meetingsSheet.getLastRow();
    const candidatesLastRow = candidatesSheet.getLastRow();

    if (meetingsLastRow < 2) {
      return {
        success: true,
        meetings: []
      };
    }

    const meetings = meetingsSheet
      .getRange(2, 1, meetingsLastRow - 1, 10)
      .getValues();

    const candidates = candidatesLastRow >= 2
      ? candidatesSheet
          .getRange(2, 1, candidatesLastRow - 1, 10)
          .getValues()
      : [];

    // Индекс кандидатов по ID.
    const candidatesById = {};

    candidates.forEach(function(row) {
      const candidateId = String(row[8] || "").trim();

      if (!candidateId) return;

      candidatesById[candidateId] = row;
    });

    const result = [];

    meetings.forEach(function(row) {
      const meetingId = String(row[0] || "").trim();
      const candidateId = String(row[1] || "").trim();
      const meetingDate = row[5];

      if (!meetingId || !candidateId || !meetingDate) {
        return;
      }

      const meetingTime = meetingDate instanceof Date
        ? meetingDate.getTime()
        : parseReminderDate_(meetingDate);

      if (!meetingTime) return;

      // Конец периода не включается.
      if (
        meetingTime < start.getTime() ||
        meetingTime >= end.getTime()
      ) {
        return;
      }

      const candidate = candidatesById[candidateId] || [];

      const dateObject = new Date(meetingTime);

      result.push({
        meetingId: meetingId,
        candidateId: candidateId,

        name: String(candidate[1] || row[2] || ""),
        phone: String(candidate[2] || row[3] || ""),
        specialty: String(candidate[3] || row[4] || ""),

        date: Utilities.formatDate(
          dateObject,
          MOSCOW_TZ,
          "yyyy-MM-dd"
        ),

        time: Utilities.formatDate(
          dateObject,
          MOSCOW_TZ,
          "HH:mm"
        ),

        meeting: Utilities.formatDate(
          dateObject,
          MOSCOW_TZ,
          "dd.MM.yyyy HH:mm"
        ),

        status: String(row[6] || ""),
        object: String(candidate[9] || "Не указан"),
        comment: String(candidate[5] || ""),
        source: String(candidate[7] || ""),
        calendarEventId: String(row[7] || "")
      });
    });

    // Сортировка по времени встречи.
    result.sort(function(a, b) {
      return a.date.localeCompare(b.date) ||
        a.time.localeCompare(b.time);
    });

    return {
      success: true,
      meetings: result
    };

  } catch (error) {
    Logger.log("getCalendarMeetings error: " + error.stack);

    return {
      success: false,
      message: error.message || "Не удалось загрузить календарь.",
      meetings: []
    };
  }
}

function testGetCalendarMeetings() {
  const result = getCalendarMeetings(
    "2026-11-23",
    "2026-11-25"
  );

  Logger.log(JSON.stringify(result, null, 2));

  return result;
}


/************************************************************
 * 24. ПРОФИЛЬ КАНДИДАТА И ИСТОРИЯ
 ************************************************************/

/**
 * Преобразование значений таблицы в безопасный текст.
 */
function profileValueToText_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(
      value,
      MOSCOW_TZ,
      "dd.MM.yyyy HH:mm"
    );
  }

  return String(value == null ? "" : value);
}

/**
 * Получение полного профиля кандидата.
 * Включает данные, встречи и историю.
 */
function getCandidateProfile(candidateId) {
  try {
    candidateId = String(candidateId || "").trim();

    if (!candidateId) {
      throw new Error("Не указан ID кандидата.");
    }

    const spreadsheetId = PropertiesService
      .getScriptProperties()
      .getProperty(SHEET_ID_PROPERTY);

    if (!spreadsheetId) {
      throw new Error("Не найдена рабочая таблица.");
    }

    const ss = SpreadsheetApp.openById(spreadsheetId);

    const candidates = ss.getSheetByName(SHEET_NAME);
    const meetings = ss.getSheetByName("Встречи");
    const history = ss.getSheetByName("История");

    if (!candidates || !meetings || !history) {
      throw new Error("Не найдены необходимые листы.");
    }

    const candidateLastRow = candidates.getLastRow();

    if (candidateLastRow < 2) {
      throw new Error("Кандидат не найден.");
    }

    const candidateRows = candidates
      .getRange(2, 1, candidateLastRow - 1, 10)
      .getValues();

    const candidateIndex = candidateRows.findIndex(row =>
      String(row[8] || "").trim() === candidateId
    );

    if (candidateIndex === -1) {
      throw new Error("Кандидат не найден.");
    }

    const row = candidateRows[candidateIndex];

    const candidate = {
      candidateId: candidateId,
      row: candidateIndex + 2,
      added: profileValueToText_(row[0]),
      name: String(row[1] || ""),
      phone: String(row[2] || ""),
      specialty: String(row[3] || ""),
      meeting: profileValueToText_(row[4]),
      comment: String(row[5] || ""),
      status: String(row[6] || ""),
      source: String(row[7] || "Не указан"),
      object: String(row[9] || "Не указан")
    };

    // Получаем все встречи кандидата.
    const meetingsLastRow = meetings.getLastRow();

    const meetingRows = meetingsLastRow >= 2
      ? meetings.getRange(2, 1, meetingsLastRow - 1, 10)
          .getValues()
      : [];

    const candidateMeetings = meetingRows
      .filter(item =>
        String(item[1] || "").trim() === candidateId
      )
      .map(item => ({
        meetingId: String(item[0] || ""),
        date: profileValueToText_(item[5]),
        status: String(item[6] || ""),
        reason: String(item[9] || ""),
        created: profileValueToText_(item[8])
      }))
      .sort((a, b) => {
        const da = a.created || a.date;
        const db = b.created || b.date;
        return db.localeCompare(da);
      });

    // Получаем историю взаимодействий.
    const historyLastRow = history.getLastRow();

    const historyRows = historyLastRow >= 2
      ? history.getRange(2, 1, historyLastRow - 1, 8)
          .getValues()
      : [];

    const candidateHistory = historyRows
      .filter(item =>
        String(item[1] || "").trim() === candidateId
      )
      .map(item => ({
        id: String(item[0] || ""),
        date: profileValueToText_(item[2]),
        action: String(item[3] || ""),
        oldValue: profileValueToText_(item[4]),
        newValue: profileValueToText_(item[5]),
        comment: String(item[6] || ""),
        author: String(item[7] || "CRM"),
        sortTime: item[2] instanceof Date
          ? item[2].getTime()
          : 0
      }))
      .sort((a, b) => b.sortTime - a.sortTime)
      .map(item => {
        delete item.sortTime;
        return item;
      });

    return {
      success: true,
      candidate: candidate,
      meetings: candidateMeetings,
      history: candidateHistory
    };

  } catch (error) {
    Logger.log("getCandidateProfile error: " + error.stack);

    return {
      success: false,
      message: error.message || "Не удалось загрузить профиль."
    };
  }
}

/**
 * Добавление заметки в историю кандидата.
 */
function addCandidateNote(candidateId, note) {
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(30000);

    candidateId = String(candidateId || "").trim();
    note = String(note || "").trim();

    if (!candidateId) {
      throw new Error("Не указан ID кандидата.");
    }

    if (!note) {
      throw new Error("Напиши текст заметки.");
    }

    if (note.length > 2000) {
      throw new Error("Заметка не должна превышать 2000 символов.");
    }

    const spreadsheetId = PropertiesService
      .getScriptProperties()
      .getProperty(SHEET_ID_PROPERTY);

    if (!spreadsheetId) {
      throw new Error("Не найдена рабочая таблица.");
    }

    const ss = SpreadsheetApp.openById(spreadsheetId);

    const candidates = ss.getSheetByName(SHEET_NAME);
    const history = ss.getSheetByName("История");

    if (!candidates || !history) {
      throw new Error("Не найдены необходимые листы.");
    }

    const lastRow = candidates.getLastRow();

    if (lastRow < 2) {
      throw new Error("Кандидат не найден.");
    }

    const ids = candidates
      .getRange(2, 9, lastRow - 1, 1)
      .getValues();

    const exists = ids.some(row =>
      String(row[0] || "").trim() === candidateId
    );

    if (!exists) {
      throw new Error("Кандидат не найден.");
    }

    const now = new Date();

    history.appendRow([
      "HIST-" + Utilities.getUuid()
        .substring(0, 8).toUpperCase(),
      candidateId,
      now,
      "Заметка рекрутера",
      "",
      "",
      note,
      "CRM"
    ]);

    SpreadsheetApp.flush();

    return {
      success: true,
      message: "Заметка сохранена."
    };

  } catch (error) {
    Logger.log("addCandidateNote error: " + error.stack);

    return {
      success: false,
      message: error.message || "Не удалось сохранить заметку."
    };

  } finally {
    if (lock.hasLock()) {
      lock.releaseLock();
    }
  }
}


/**
 * АНАЛИТИКА STROY PODRYAD CRM
 * Периоды: today, week, month, lastMonth, all, custom
 */
function getRecruitmentAnalytics(filters) {
  try {
    filters = filters || {};

    const properties = PropertiesService.getScriptProperties();
    const spreadsheetId = properties.getProperty(SHEET_ID_PROPERTY);

    if (!spreadsheetId) {
      throw new Error("Не указан ID рабочей таблицы.");
    }

    const ss = SpreadsheetApp.openById(spreadsheetId);
    const candidatesSheet = ss.getSheetByName(SHEET_NAME);
    const meetingsSheet = ss.getSheetByName("Встречи");
    const historySheet = ss.getSheetByName("История");

    if (!candidatesSheet || !meetingsSheet || !historySheet) {
      throw new Error("Не найдены необходимые листы CRM.");
    }

    const now = new Date();
    const period = filters.period || "month";

    let startDate = null;
    let endDate = new Date(now);
    endDate.setHours(23, 59, 59, 999);

    if (period === "today") {
      startDate = new Date(now);
      startDate.setHours(0, 0, 0, 0);

    } else if (period === "week") {
      startDate = new Date(now);
      startDate.setDate(startDate.getDate() - 6);
      startDate.setHours(0, 0, 0, 0);

    } else if (period === "month") {
      startDate = new Date(now.getFullYear(), now.getMonth(), 1);

    } else if (period === "lastMonth") {
      startDate = new Date(now.getFullYear(), now.getMonth() - 1, 1);
      endDate = new Date(now.getFullYear(), now.getMonth(), 0);
      endDate.setHours(23, 59, 59, 999);

    } else if (period === "custom") {
      startDate = parseAnalyticsDate_(filters.startDate);
      endDate = parseAnalyticsDate_(filters.endDate);

      if (!startDate || !endDate) {
        throw new Error("Укажите корректные даты периода.");
      }

      startDate.setHours(0, 0, 0, 0);
      endDate.setHours(23, 59, 59, 999);

      if (startDate > endDate) {
        throw new Error("Начало периода позже его окончания.");
      }

    } else if (period !== "all") {
      throw new Error("Неизвестный период аналитики.");
    }

    function isInPeriod(date) {
      if (period === "all") return true;
      if (!date) return false;

      return date >= startDate && date <= endDate;
    }

    function countBy(items, getter) {
      const result = {};

      items.forEach(function(item) {
        const value = String(getter(item) || "Не указан");
        result[value] = (result[value] || 0) + 1;
      });

      return Object.keys(result)
        .map(function(name) {
          return { name: name, count: result[name] };
        })
        .sort(function(a, b) {
          return b.count - a.count;
        });
    }

    // Кандидаты: A — дата, D — специальность,
    // G — статус, H — источник, I — ID, J — объект.
    const lastCandidateRow = candidatesSheet.getLastRow();

    const candidateRows = lastCandidateRow >= 2
      ? candidatesSheet
          .getRange(2, 1, lastCandidateRow - 1, 10)
          .getValues()
      : [];

    const allCandidates = candidateRows.map(function(row) {
      return {
        added: parseAnalyticsDate_(row[0]),
        name: String(row[1] || ""),
        specialty: String(row[3] || "Не указана"),
        status: String(row[6] || "Не указан").trim(),
        source: String(row[7] || "Не указан"),
        id: String(row[8] || ""),
        object: String(row[9] || "Не указан")
      };
    });

    // Кандидаты, добавленные в выбранный период.
    const candidates = allCandidates.filter(function(candidate) {
      return isInPeriod(candidate.added);
    });

    // История взаимодействий.
    const lastHistoryRow = historySheet.getLastRow();

    const historyRows = lastHistoryRow >= 2
      ? historySheet.getRange(2, 1, lastHistoryRow - 1, 8).getValues()
      : [];

    const historyByCandidate = {};

    historyRows.forEach(function(row) {
      const candidateId = String(row[1] || "");
      if (!candidateId) return;

      const item = {
        date: parseAnalyticsDate_(row[2]),
        action: String(row[3] || ""),
        newValue: String(row[5] || "").trim()
      };

      if (!historyByCandidate[candidateId]) {
        historyByCandidate[candidateId] = [];
      }

      historyByCandidate[candidateId].push(item);
    });

    let arrivedCount = 0;
    let startedCount = 0;

    const arrivedStatuses = [
      "Приехал на объект",
      "На испытательном сроке",
      "Работает",
      "Не прошёл испытательный срок"
    ];

    candidates.forEach(function(candidate) {
      const history = historyByCandidate[candidate.id] || [];

      const hasArrived = arrivedStatuses.indexOf(candidate.status) !== -1 ||
        history.some(function(item) {
          return item.newValue === "Приехал на объект" ||
                 item.newValue === "На испытательном сроке" ||
                 item.newValue === "Работает" ||
                 item.newValue === "Не прошёл испытательный срок";
        });

      const hasStarted = candidate.status === "Работает" ||
        history.some(function(item) {
          return item.newValue === "Работает";
        });

      if (hasArrived) arrivedCount++;
      if (hasStarted) startedCount++;
    });

    // Встречи, назначенные в выбранный период.
    const lastMeetingRow = meetingsSheet.getLastRow();

    const meetingRows = lastMeetingRow >= 2
      ? meetingsSheet.getRange(2, 1, lastMeetingRow - 1, 10).getValues()
      : [];

    const meetingsInPeriod = meetingRows.filter(function(row) {
      return isInPeriod(parseAnalyticsDate_(row[5]));
    });

    const total = candidates.length;

    function percentage(value) {
      return total ? Math.round(value / total * 100) : 0;
    }

    return {
      success: true,
      period: period,

      dates: {
        start: startDate
          ? Utilities.formatDate(startDate, MOSCOW_TZ, "dd.MM.yyyy")
          : "За всё время",
        end: period === "all"
          ? "По сегодняшний день"
          : Utilities.formatDate(endDate, MOSCOW_TZ, "dd.MM.yyyy")
      },

      summary: {
        total: total,
        arrived: arrivedCount,
        started: startedCount,
        meetings: meetingsInPeriod.length,
        arrivalConversion: percentage(arrivedCount),
        startConversion: percentage(startedCount)
      },

      statuses: countBy(candidates, function(c) {
        return c.status;
      }),

      specialties: countBy(candidates, function(c) {
        return c.specialty;
      }),

      sources: countBy(candidates, function(c) {
        return c.source;
      }),

      objects: countBy(candidates, function(c) {
        return c.object;
      })

    };

  } catch (error) {
    Logger.log("getRecruitmentAnalytics error: " + error.stack);

    return {
      success: false,
      message: error.message || "Ошибка получения аналитики."
    };
  }
}


/**
 * Приведение даты к объекту Date.
 */
function parseAnalyticsDate_(value) {
  if (!value) return null;

  if (value instanceof Date && !isNaN(value.getTime())) {
    return new Date(value.getTime());
  }

  if (typeof value === "string") {
    const text = value.trim();
    if (!text) return null;

    // Формат dd.MM.yyyy или dd.MM.yyyy HH:mm
    const ru = text.match(
      /^(\d{1,2})\.(\d{1,2})\.(\d{4})(?:\s+(\d{1,2}):(\d{2}))?$/
    );

    if (ru) {
      return new Date(
        Number(ru[3]),
        Number(ru[2]) - 1,
        Number(ru[1]),
        Number(ru[4] || 0),
        Number(ru[5] || 0)
      );
    }

    const parsed = new Date(text);

    if (!isNaN(parsed.getTime())) {
      return parsed;
    }
  }

  return null;
}

function testRecruitmentAnalytics() {
  const result = getRecruitmentAnalytics({
    period: "month"
  });

  Logger.log(JSON.stringify(result, null, 2));
}


/**
 * Расписание для прораба.
 * Только чтение листа «Встречи».
 */

function getForemanSchedule(dateIso) {
  const props = PropertiesService.getScriptProperties();
  const spreadsheetId = props.getProperty(SHEET_ID_PROPERTY);
  const candidateSheetName = SHEET_NAME;

  if (!spreadsheetId) {
    throw new Error("Не найден ID таблицы");
  }

  const ss = SpreadsheetApp.openById(spreadsheetId);
  const meetingSheet = ss.getSheetByName("Встречи");
  const candidateSheet = ss.getSheetByName(candidateSheetName);

  if (!meetingSheet) {
    throw new Error('Не найден лист «Встречи»');
  }

  if (!candidateSheet) {
    throw new Error('Не найден лист «Кандидаты»');
  }

  const timezone = ss.getSpreadsheetTimeZone();
  const selectedDate = dateIso || Utilities.formatDate(
    new Date(), timezone, "yyyy-MM-dd"
  );

  // Собираем комментарии кандидатов по их ID
  const candidateRows = candidateSheet.getDataRange().getValues();
  const commentsById = {};

  for (let i = 1; i < candidateRows.length; i++) {
    const candidateId = String(candidateRows[i][8] || "").trim();
    const comment = String(candidateRows[i][5] || "").trim();

    if (candidateId) {
      commentsById[candidateId] = comment;
    }
  }

  // Получаем встречи
  const rows = meetingSheet.getDataRange().getValues();
  const meetings = [];

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];

    const meetingId = String(row[1] || "").trim();
    const name = String(row[2] || "").trim();
    const phone = String(row[3] || "").trim();
    const specialty = String(row[4] || "").trim();
    const meetingDate = row[5];
    const meetingStatus = String(row[6] || "").trim();

    if (!(meetingDate instanceof Date)) continue;

    const meetingDay = Utilities.formatDate(
      meetingDate, timezone, "yyyy-MM-dd"
    );

    if (meetingDay !== selectedDate) continue;
    if (/отмен|cancel/i.test(meetingStatus)) continue;
    if (!name && !phone) continue;

    meetings.push({
      name: name,
      phone: phone,
      specialty: specialty,
      comment: commentsById[meetingId] || "",
      time: Utilities.formatDate(
        meetingDate, timezone, "HH:mm"
      )
    });
  }

  meetings.sort((a, b) => a.time.localeCompare(b.time));

  return {
    success: true,
    date: selectedDate,
    count: meetings.length,
    meetings: meetings
  };
}

function testForemanSchedule() {
  const result = getForemanSchedule("");
  Logger.log(JSON.stringify(result, null, 2));
}

function testSheetNames() {
  const id = PropertiesService.getScriptProperties()
    .getProperty(SHEET_ID_PROPERTY);

  const ss = SpreadsheetApp.openById(id);

  Logger.log("Листы таблицы:");

  ss.getSheets().forEach(sheet => {
    Logger.log(sheet.getName());
  });

  Logger.log("SHEET_NAME = " + SHEET_NAME);
}
