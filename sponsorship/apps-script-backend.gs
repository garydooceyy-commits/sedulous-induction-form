var NOTIFY_EMAIL = 'gary@sedulousrail.co.uk';

function doPost(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['Submitted At', 'Submission ID', 'Name', 'Relationship', 'Countersigned', 'Filed', 'Data (JSON)']);
  }

  var body = JSON.parse(e.postData.contents);
  var id = Utilities.getUuid();

  sheet.appendRow([
    new Date(),
    id,
    body.name || '',
    body.relationship || '',
    false,
    false,
    JSON.stringify(body)
  ]);

  notifySubmission(body.name || '(no name given)', body.relationship || '', id);

  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, id: id }))
    .setMimeType(ContentService.MimeType.JSON);
}

function notifySubmission(name, relationship, id) {
  try {
    MailApp.sendEmail({
      to: NOTIFY_EMAIL,
      subject: 'New Contract of Sponsorship submitted: ' + name,
      body:
        'A new Contract of Sponsorship has just come in.\n\n' +
        'Name: ' + name + '\n' +
        'Contractual relationship: ' + relationship + '\n' +
        'Submission ID: ' + id + '\n\n' +
        'This is pending your countersignature as Company Rail Representative.\n' +
        'Open the "Sponsorship Submissions" sheet, or just tell Claude to check for new starters.'
    });
  } catch (err) {
    // Submission is already saved in the sheet either way; don't fail the request over a failed email.
  }
}

function doGet(e) {
  return ContentService.createTextOutput('Sponsorship form backend is running.');
}
