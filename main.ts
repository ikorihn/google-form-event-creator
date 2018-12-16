import { CONSTANTS } from './const';

/**
 * 入力された日付の該当イベントを取得
 * @param eventDay イベントの日付
 */
function getEventId(eventDay: Date): GoogleAppsScript.Calendar.CalendarEvent {
  const calendar = CalendarApp.getCalendarById(CONSTANTS.CAL_ID);
  const events = calendar.getEventsForDay(eventDay);
  for (let event of events) {
    if (event.getTitle().search(CONSTANTS.EVENT_NAME) !== -1) {
      return event;
    }
  }
}

function main() {
  let event = getEventId(new Date('2018-12-17'));
  Logger.log(event.getTitle());
}
