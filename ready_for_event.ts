import { Conference } from "./conference";
import { formatLiteral, addBr } from "./util";
const PROPERTY = PropertiesService.getScriptProperties()

/**
 * 準備する
 */
function readyForConference() {
  const sheet = SpreadsheetApp.openById(PROPERTY.getProperty('MASTER_SHEET')).getSheetByName('フォーム作成');
  const rownum = sheet.getLastRow();
  let colnum = 1;

  const conference = new Conference();
  conference.date = new Date(sheet.getRange(rownum, colnum++).getValue().toString());
  conference.name = sheet.getRange(rownum, colnum++).getValue().toString();
  conference.email = sheet.getRange(rownum, colnum++).getValue().toString();
  conference.title = sheet.getRange(rownum, colnum++).getValue().toString();
  conference.description = sheet.getRange(rownum, colnum++).getValue().toString();
  conference.target = sheet.getRange(rownum, colnum++).getValue().toString();
  Logger.log(conference.toString());

  const attendeeUrl = createAttendeeForm(conference);
  createAttendeeMail(conference, attendeeUrl);
  const enqueteUrl = createEnqueteForm(conference);
  createEnqueteMail(conference, enqueteUrl);
  editCalendarEvent(conference);
  addDirectoryAtSharedDrive(conference);
}

/**
 * 日本の曜日を返す
 * @param date 
 */
function getJapaneseDayOfWeek(date: Date) {
  const japanese_week_of_day = ['日', '月', '火', '水', '木', '金', '土']
  const daynum = parseInt(Utilities.formatDate(date, 'JST', 'u'))
  return japanese_week_of_day[daynum];
}
/**
 * 日付を yyyy/MM/dd(E) にフォーマットする
 * @param date 
 */
function formatDate(date: Date): string {
  return `${Utilities.formatDate(date, "JST", "yyyy/MM/dd")}(${getJapaneseDayOfWeek(date)})`;
}

/**
 * 参加登録メールを作成する
 * @param conference 
 */
function createAttendeeMail(conference: Conference, formUrl: string) {
  const subject = `開催案内${conference.title}`;
  const body = `
<b style="color: #0000ff">${conference.title}</b>( ${formatDate(conference.date)} ) が開催されます<br>
下記より参加登録をお願いします。<br>
${formUrl}
`;

  GmailApp.createDraft(conference.email, subject, '', {
    cc: `${conference.email},${PROPERTY.getProperty('GROUP_MAIL')}`,
    htmlBody: body
  });
}

/**
 * 参加登録フォームを作成する
 * @param conference 
 * @returns フォームのURL
 */
function createAttendeeForm(conference: Conference) {
  const form = FormApp.create(`参加登録( ${formatDate(conference.date)} )`);

  form.setDescription('参加')
    .setAllowResponseEdits(true)
    .setAcceptingResponses(true)
    .setCollectEmail(true)
    .setDestination(FormApp.DestinationType.SPREADSHEET, PROPERTY.getProperty('ATTENDEE_SHEET'))
    ;

  form.addEditors([conference.email, PROPERTY.getProperty('GROUP_MAIL')]);

  const attend = form.addCheckboxItem();
  attend.setTitle(`${formatDate(conference.date)} のイベントに参加`)
    .setChoices([attend.createChoice('参加する')])
    .setRequired(true);

  return form.getPublishedUrl();
}

/**
 * アンケートメールを作成する
 * @param conference 
 */
function createEnqueteMail(conference: Conference, formUrl: string) {
  const subject = `受講後アンケート「${conference.title}」`;
  const body = addBr(formatLiteral(`本日はご参加ありがとうございました。
              |下記アンケートへのご回答よろしくお願いします。
              |${formUrl}
              |`));

  GmailApp.createDraft(conference.email, subject, '', {
    cc: `${conference.email},${PROPERTY.getProperty('GROUP_MAIL')}`,
    htmlBody: body
  });
}

/**
 * アンケートフォームを作成する
 * @param conference 
 * @returns フォームのURL
 */
function createEnqueteForm(conference: Conference): string {
  const form = FormApp.create(`受講後アンケート( ${formatDate(conference.date)} )`);

  form.setDescription('アンケートに入力してください')
    .setAllowResponseEdits(true)
    .setAcceptingResponses(true)
    .setCollectEmail(true)
    ;

  form.addEditors([conference.email, PROPERTY.getProperty('GROUP_MAIL')]);

  form.addScaleItem()
    .setTitle('満足度')
    .setHelpText('')
    .setBounds(1, 10)
    .setRequired(true);

  form.addParagraphTextItem()
    .setTitle('理由')
    .setHelpText('')
    .setRequired(false);

  return form.getPublishedUrl();
}

/**
 * 入力された日付の該当イベントを取得
 * イベントがない場合は作成する
 * @param 日付
 * @returns カレンダーイベント
 */
function getEvent(date: Date): GoogleAppsScript.Calendar.CalendarEvent {
  const calendar = CalendarApp.getCalendarById(PROPERTY.getProperty('CAL_ID'));
  const events = calendar.getEventsForDay(date);
  for (let event of events) {
    if (event.getTitle().indexOf(PROPERTY.getProperty('EVENT_NAME')) !== -1) {
      return event;
    }
  }
  // イベントがない場合
  const startTime = new Date(date.getFullYear(), date.getMonth(), date.getDate(), 12, 0, 0);
  const endTime = new Date(date.getFullYear(), date.getMonth(), date.getDate(), 13, 0, 0);
  return calendar.createEvent('', startTime, endTime);
}

/**
 * カレンダーのイベントにタイトル、内容、開催者をセットする
 * @param conference 
 */
function editCalendarEvent(conference: Conference) {
  const event = getEvent(conference.date);
  event.setTitle(`講演「${conference.title}」`);
  event.setDescription(`title
${conference.title}
  `);
  event.addGuest(conference.email);
  event.addGuest(PROPERTY.getProperty('GROUP_MAIL'));

}

/**
 * Google Drive にフォルダを作成する
 * @param conference 
 * @returns Drive のURL
 */
function addDirectoryAtSharedDrive(conference: Conference): string {
  const rootDir = DriveApp.getFolderById(PROPERTY.getProperty('SHARE_DRIVE_ID'));
  const url = rootDir.createFolder(`${Utilities.formatDate(conference.date, "JST", "yyyyyMMdd")}_${conference.title}`);
  return url.getUrl();
}

/**
 * スプレッドシートにメニューを追加する
 */
function customizeMenu() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('独自メニュー')
    .addItem('準備をする', 'readyForConference')
    .addToUi();
}

function onOpen() {
  customizeMenu();
}