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
  const items = form.getItems();
  for (let item of items) {
    if (item.getType() === FormApp.ItemType.SECTION_HEADER
      && item.getTitle() === CONSTANTS.TITLE) {
      return new Date(item.getHelpText());
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

function main() {
  let event = getEventId(new Date('2018-12-17'));
  Logger.log(event.getTitle());
}
