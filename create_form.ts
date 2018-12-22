const PROPERTY = PropertiesService.getScriptProperties();

function createAttendeeMail(conference: Conference, formUrl: string) {
  const subject = `開催案内${conference.title}`;
  const body = `
<b style="color: #0000ff">${conference.title}</b>( ${Utilities.formatDate(conference.date, "JST", "yyyyy/MM/dd(EEE)")}) が開催されます
${formUrl}
`;

  GmailApp.createDraft(conference.email, subject, body, {
    cc: [conference.email, PROPERTY.getProperty('GROUP_MAIL')],
    htmlBody: body
  });
}

function createAttendeeForm(conference: Conference) {
  const form = FormApp.create(`参加登録( ${conference.date} )`);

  form.setDescription('参加')
    .setAllowResponseEdits(true)
    .setAcceptingResponses(true)
    .setCollectEmail(false)
    .setDestination(GoogleAppsScript.Forms.DestinationType.SPREADSHEET, PROPERTY.getProperty('ATTENDEE_SHEET'))
    ;

  form.addEditor(conference.email);

  const attend = form.addCheckboxItem();
  attend.setTitle(`${conference.date} のイベントに参加`)
    .setChoices([attend.createChoice('参加する')]);

  return form.getPublishedUrl();
}

function createEnqueteMail(conference: Conference, formUrl: string) {
  const subject = `受講後アンケート${conference.title}`;
  const body = `
受講後アンケート <b style="color: #0000ff">${conference.title}</b>
${formUrl}
`;

  GmailApp.createDraft(conference.email, subject, body, {
    cc: [conference.email, PROPERTY.getProperty('GROUP_MAIL')],
    htmlBody: body
  });
}

function createEnqueteForm(conference: Conference): string {
  const form = FormApp.create(`受講後アンケート( ${conference.date} )`);

  form.setDescription('アンケートに入力してください')
    .setAllowResponseEdits(true)
    .setAcceptingResponses(true)
    .setCollectEmail(false)
    ;

  form.addEditor(conference.email);

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

function readyForConference() {
  const sheet = SpreadsheetApp.openById(PROPERTY.getProperty('MASTER_SHEET')).getSheetByName('フォーム作成');
  const rownum = sheet.getActiveRange().getRow();
  let colnum = 1;

  const conference = new Conference();
  conference.date = new Date(sheet.getRange(rownum, colnum++).getValue().toString());
  conference.name = sheet.getRange(rownum, colnum++).getValue().toString();
  conference.email = sheet.getRange(rownum, colnum++).getValue().toString();
  conference.title = sheet.getRange(rownum, colnum++).getValue().toString();
  conference.description = sheet.getRange(rownum, colnum++).getValue().toString();
  conference.target = sheet.getRange(rownum, colnum++).getValue().toString();

  const attendeeUrl = createAttendeeForm(conference);
  createAttendeeMail(conference, attendeeUrl);
  const enqueteUrl = createEnqueteForm(conference);
  createEnqueteMail(conference, enqueteUrl);
  addDirectoryAtSharedDrive(conference);
  addConfluencePage(conference);
}