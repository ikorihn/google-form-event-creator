function createForm(title: string, date: string) {
  const form = FormApp.create(title);
  const sheet = SpreadsheetApp.getActiveSpreadsheet();

  form.setDescription('Description of form')
    .setAllowResponseEdits(true)
    .setCollectEmail(true)
    .setDestination(FormApp.DestinationType.SPREADSHEET, sheet.getId())
    ;

  form.addEditor('mailadderess');

  const attend = form.addCheckboxItem();
  attend.setTitle(`${date} のイベントに参加`)
    .setChoices([attend.createChoice('参加する')]);

}