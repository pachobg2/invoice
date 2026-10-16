/**
 * Receiving end of "Изпрати данни за тестване" (invoice program) - a Google Apps Script web app.
 * It receives the zip that the program uploads over HTTPS and e-mails it to the developer.
 * No password is stored in the program: the mail is sent by this script from the developer's own Google account.
 *
 * Set-up (once, about 2 minutes):
 *  1. Open https://script.google.com  ->  New project  ->  replace the code with this file  ->  Save.
 *  2. Deploy  ->  New deployment  ->  type "Web app".
 *       Execute as: Me          Who has access: Anyone
 *     Deploy, allow the permission request (it asks to send e-mail as you), copy the "Web app" URL (…/macros/s/…/exec).
 *  3. Put that URL into host\SendData.cs (Endpoint) and release a new version.
 * After changing this script use  Deploy -> Manage deployments -> Edit -> New version  (the URL stays the same).
 */
const TO = 'yankovpavel@gmail.com';
const MAX_BYTES = 5 * 1024 * 1024;   // refuse files bigger than 5 MB
const MAX_PER_DAY = 20;              // ... and more than 20 messages a day (protects against someone abusing the public address)

function doPost(e) {
  try {
    const props = PropertiesService.getScriptProperties();
    const day = Utilities.formatDate(new Date(), 'UTC', 'yyyy-MM-dd');
    const key = 'n_' + day;
    const n = parseInt(props.getProperty(key) || '0', 10);
    if (n >= MAX_PER_DAY) return out({ ok: false, error: 'daily limit' });

    const d = JSON.parse(e.postData.contents);
    if (d.app !== 'invoice' || typeof d.zip !== 'string' || d.zip.length > MAX_BYTES * 1.4) return out({ ok: false, error: 'bad request' });
    const bytes = Utilities.base64Decode(d.zip);
    if (bytes.length > MAX_BYTES || bytes.length < 4 || bytes[0] !== 80 || bytes[1] !== 75) return out({ ok: false, error: 'not a zip' });   // "PK"

    const name = String(d.name || 'invoice-data.zip').replace(/[^A-Za-z0-9._-]/g, '_').slice(0, 80);
    MailApp.sendEmail({
      to: TO,
      subject: 'Invoice program data for testing (' + String(d.version || '?').slice(0, 20) + ')',
      body: 'Sent from the invoice program.\n\n' + String(d.info || '').slice(0, 500),
      attachments: [Utilities.newBlob(bytes, 'application/zip', name)]
    });
    props.setProperty(key, String(n + 1));
    return out({ ok: true });
  } catch (err) {
    return out({ ok: false, error: String(err).slice(0, 200) });
  }
}

function out(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
