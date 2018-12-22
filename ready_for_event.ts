const PROPERTY = PropertiesService.getScriptProperties()

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
  // addConfluencePage(conference);
}

function getJapaneseDayOfWeek(date: Date) {
  const japanese_week_of_day = ['日', '月', '火', '水', '木', '金', '土']
  const daynum = parseInt(Utilities.formatDate(date, 'JST', 'u'))
  return japanese_week_of_day[daynum - 1];
}

function createAttendeeMail(conference: Conference, formUrl: string) {
  const subject = `開催案内${conference.title}`;
  const body = `
<b style="color: #0000ff">${conference.title}</b>( ${Utilities.formatDate(conference.date, "JST", "yyyy/MM/dd")}(${getJapaneseDayOfWeek(conference.date)}) ) が開催されます<br>
下記より参加登録をお願いします。<br>
${formUrl}
`;

  GmailApp.createDraft(conference.email, subject, '', {
    cc: `${conference.email},${PROPERTY.getProperty('GROUP_MAIL')}`,
    htmlBody: body
  });
}

function createAttendeeForm(conference: Conference) {
  const form = FormApp.create(`参加登録( ${Utilities.formatDate(conference.date, "JST", "yyyy/MM/dd")}(${getJapaneseDayOfWeek(conference.date)}) )`);

  form.setDescription('参加')
    .setAllowResponseEdits(true)
    .setAcceptingResponses(true)
    .setCollectEmail(true)
    .setDestination(FormApp.DestinationType.SPREADSHEET, PROPERTY.getProperty('ATTENDEE_SHEET'))
    ;

  form.addEditors([conference.email, PROPERTY.getProperty('GROUP_MAIL')]);

  const attend = form.addCheckboxItem();
  attend.setTitle(`${Utilities.formatDate(conference.date, "JST", "yyyy/MM/dd")}(${getJapaneseDayOfWeek(conference.date)}) のイベントに参加`)
    .setChoices([attend.createChoice('参加する')])
    .setRequired(true);

  return form.getPublishedUrl();
}

function createEnqueteMail(conference: Conference, formUrl: string) {
  const subject = `受講後アンケート「${conference.title}」`;
  const body = `
本日はご参加ありがとうございました。<br>
下記アンケートへのご回答よろしくお願いします。<br>
${formUrl}
`;

  GmailApp.createDraft(conference.email, subject, '', {
    cc: `${conference.email},${PROPERTY.getProperty('GROUP_MAIL')}`,
    htmlBody: body
  });
}

function createEnqueteForm(conference: Conference): string {
  const form = FormApp.create(`受講後アンケート( ${Utilities.formatDate(conference.date, "JST", "yyyy/MM/dd")}(${getJapaneseDayOfWeek(conference.date)}) )`);

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
 */
function getEvent(conference: Conference): GoogleAppsScript.Calendar.CalendarEvent {
  const calendar = CalendarApp.getCalendarById(PROPERTY.getProperty('CAL_ID'));
  const events = calendar.getEventsForDay(conference.date);
  for (let event of events) {
    if (event.getTitle().indexOf(PROPERTY.getProperty('EVENT_NAME')) !== -1) {
      return event;
    }
  }
  // イベントがない場合
  const startTime = new Date(conference.date.getFullYear(), conference.date.getMonth(), conference.date.getDate(), 12, 0, 0);
  const endTime = new Date(conference.date.getFullYear(), conference.date.getMonth(), conference.date.getDate(), 13, 0, 0);
  return calendar.createEvent(conference.title, startTime, endTime);
}

function editCalendarEvent(conference: Conference) {
  const event = getEvent(conference);
  event.setTitle(`講演「${conference.title}」`);
  event.setDescription(`
title
${conference.title}
  `);
  event.addGuest(conference.email);
  event.addGuest(PROPERTY.getProperty('GROUP_MAIL'));

}

function addDirectoryAtSharedDrive(conference: Conference): string {
  const rootDir = DriveApp.getFolderById(PROPERTY.getProperty('SHARE_DRIVE_ID'));
  const url = rootDir.createFolder(`${Utilities.formatDate(conference.date, "JST", "yyyyyMMdd")}_${conference.title}`);
  return url.getUrl();
}

function addConfluencePage(conference: Conference) {
  const body = `
  `;
  const method: 'post' = 'post';
  const options = {
    method: method,
    headers: '',
    payload: body
  };

  const response = UrlFetchApp.fetch(PROPERTY.getProperty('CONFLUENCE_URL'), options);
  const result = JSON.parse(response.getContentText());
  return result["url"];
}
