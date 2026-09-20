var NOTIFY_EMAIL = 'gary@sedulousrail.co.uk';

function doPost(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['Submitted At', 'Submission ID', 'Name', 'Division', 'Reason', 'Items Requested', 'Issued', 'Data (JSON)']);
  }

  var body = JSON.parse(e.postData.contents);
  var id = Utilities.getUuid();

  sheet.appendRow([
    new Date(),
    id,
    body.name || '',
    body.division || '',
    body.reason || '',
    body.requestedCount || 0,
    false,
    JSON.stringify(body)
  ]);

  notifySubmission(body.name || '(no name given)', body.division || '', body.requestedCount || 0, id);

  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, id: id }))
    .setMimeType(ContentService.MimeType.JSON);
}

function notifySubmission(name, division, itemCount, id) {
  try {
    MailApp.sendEmail({
      to: NOTIFY_EMAIL,
      subject: 'New PPE request submitted: ' + name,
      body:
        'A new PPE request has just come in.\n\n' +
        'Name: ' + name + '\n' +
        'Division: ' + division + '\n' +
        'Items requested: ' + itemCount + '\n' +
        'Submission ID: ' + id + '\n\n' +
        'Open the "PPE Requests" sheet to see the full item list, or just tell Claude to check for new starters.'
    });
  } catch (err) {
    // Submission is already saved in the sheet either way; don't fail the request over a failed email.
  }
}

function doGet(e) {
  return ContentService.createTextOutput('PPE request form backend is running.');
}
