var NOTIFY_EMAIL = 'gary@sedulousrail.co.uk';

function doPost(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['Submitted At', 'Submission ID', 'Name', 'Role', 'Filed', 'Data (JSON)']);
  }

  var body = JSON.parse(e.postData.contents);
  var id = Utilities.getUuid();

  sheet.appendRow([
    new Date(),
    id,
    body.name || '',
    body.role || '',
    false,
    JSON.stringify(body)
  ]);

  notifySubmission(body.name || '(no name given)', body.role || '', id);

  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, id: id }))
    .setMimeType(ContentService.MimeType.JSON);
}

function notifySubmission(name, role, id) {
  try {
    MailApp.sendEmail({
      to: NOTIFY_EMAIL,
      subject: 'New induction pack submitted: ' + name,
      body:
        'A new induction pack has just come in.\n\n' +
        'Name: ' + name + '\n' +
        'Role: ' + role + '\n' +
        'Submission ID: ' + id + '\n\n' +
        'Open the "Induction Submissions" sheet, or just tell Claude to check for new starters.'
    });
  } catch (err) {
    // Submission is already saved in the sheet either way; don't fail the request over a failed email.
  }
}

function doGet(e) {
  return ContentService.createTextOutput('Induction form backend is running.');
}
