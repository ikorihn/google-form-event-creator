import { CONSTANTS } from './const';

/**
 * 入力された日付の該当イベントを取得
 * @param eventDay イベントの日付
 */
function getEventId(eventDay: Date): GoogleAppsScript.Calendar.CalendarEvent {
  const property = PropertiesService.getScriptProperties();
  const calendar = CalendarApp.getCalendarById(property.getProperty('CAL_ID'));
  const events = calendar.getEventsForDay(eventDay);
  for (let event of events) {
    if (event.getTitle().indexOf(property.getProperty('EVENT_NAME')) !== -1) {
      return event;
    }
  }
}

/**
 * Form に設定している日時を取得する
 * @param sheet Spreadsheet
 */
function getDate(sheet: GoogleAppsScript.Spreadsheet.Sheet): Date {
  const firstRow = sheet.getRange(1, 1, 1, sheet.getLastColumn());
  for (let cell of firstRow.getValues()[0]) {
    if (cell.toString().search(/\d{4}\/\d{2}\/\d{2}/) !== -1) {
      const date = cell.toString().match(/\d{4}\/\d{2}\/\d{2}/);
      return new Date(date[0]);
    }
  }
}

/**
 * メールアドレス列取得
 * @param sheet Spreadsheet
 */
function getMailColumn(sheet: GoogleAppsScript.Spreadsheet.Sheet): number {
  const firstRow = sheet.getRange(1, 1, 1, sheet.getLastColumn());
  const values = firstRow.getValues()[0];
  for (let i = 0; i < values.length; i++) {
    if (values[i].toString().indexOf('メールアドレス') !== -1) {
      return i;
    }
  }
}

/**
 * フォーム送信時に実行される
 * @param e form event
 */
function onFormSubmit(e) {
  const email: string = e.response.getRespondentEmail();

  const calendarEvent = getEventId(getDate());
  calendarEvent.addGuest(email);
}

/**
 * トリガーを作成する
 */
function createTrigger() {
  ScriptApp.newTrigger('onFormSubmit')
    // .forForm(FormApp.getActiveForm())
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onFormSubmit()
    .create();
}

function main() {
  let event = getEventId(new Date('2018-12-17'));
  Logger.log(event.getTitle());
}
