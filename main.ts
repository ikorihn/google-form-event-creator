import { CONSTANTS } from './const';

/**
 * 入力された日付の該当イベントを取得
 * @param eventDay イベントの日付
 */
function getEventId(eventDay: Date): GoogleAppsScript.Calendar.CalendarEvent {
  const calendar = CalendarApp.getCalendarById(CONSTANTS.CAL_ID);
  const events = calendar.getEventsForDay(eventDay);
  for (let event of events) {
    if (event.getTitle().indexOf(CONSTANTS.EVENT_NAME) !== -1) {
      return event;
    }
  }
}

/**
 * Form に設定している日時を取得する
 */
function getDate(): Date {
  const form = FormApp.getActiveForm();
  const title = form.getTitle();
  const date = title.match(/\d{4}\/\d{2}\/\d{2}/);
  return new Date(date[0]);
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

function main() {
  let event = getEventId(new Date('2018-12-17'));
  Logger.log(event.getTitle());
}
