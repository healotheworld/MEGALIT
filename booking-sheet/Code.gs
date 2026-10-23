// Receives requests from megalit.studio/book and adds them to the "Requests" tab.
// Setup: open the sheet → Extensions → Apps Script → paste this file → Deploy →
// New deployment → Web app, Execute as "Me", Who has access "Anyone".

const SHEET_NAME = "Requests";
// Gets an email for each new request. Leave empty ("") to turn emails off.
const NOTIFY_EMAIL = "simon@megalit.studio";

// Same order as the sheet columns after "Received" and "Status".
const FIELDS = [
  "Host", "Occasion", "Theme", "Guests", "Location", "Setting", "Day or night", "Length",
  "Event date", "End date", "Services", "Name", "Email", "Phone", "Company", "Notes", "Source",
];

function doPost(e) {
  let data;
  try {
    data = JSON.parse(e.postData.contents);
  } catch (err) {
    return reply({ ok: false, error: "Invalid request" });
  }

  // Hidden field only bots fill in: pretend success, store nothing.
  if (data.website) return reply({ ok: true });
  if (!data.Name || !data.Email) return reply({ ok: false, error: "Missing name or email" });

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    requestsSheet().appendRow([new Date(), "New"].concat(FIELDS.map(f => clean(data[f]))));
  } finally {
    lock.releaseLock();
  }

  if (NOTIFY_EMAIL) {
    const lines = FIELDS.filter(f => data[f]).map(f => f + ": " + data[f]);
    const message = {
      to: NOTIFY_EMAIL,
      subject: String(data._subject || "New event request").slice(0, 200),
      body: lines.join("\n") + "\n\nAll requests: " + SpreadsheetApp.getActive().getUrl(),
    };
    if (/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(data.Email)) message.replyTo = data.Email;
    try { MailApp.sendEmail(message); } catch (err) { console.error(err); }
  }

  return reply({ ok: true });
}

// Finds the "Requests" tab, or creates it with headers and a Status dropdown on first use.
function requestsSheet() {
  const book = SpreadsheetApp.getActive();
  let sheet = book.getSheetByName(SHEET_NAME);
  if (sheet) return sheet;

  sheet = book.insertSheet(SHEET_NAME, 0);
  const headers = ["Received", "Status"].concat(FIELDS);
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold").setBackground("#eeeeee");
  sheet.setFrozenRows(1);
  sheet.getRange("B2:B").setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(["New", "Contacted", "Proposal sent", "Booked", "Declined"], true)
      .setAllowInvalid(true)
      .build()
  );
  return sheet;
}

// Run once from the Apps Script editor (select "setup", click Run) to create the tab and grant permissions.
function setup() {
  requestsSheet();
}

// Keeps text as plain text, so input like "=SUM(...)" or "+46 70..." is never run as a formula.
function clean(value) {
  const text = String(value == null ? "" : value).slice(0, 2000);
  return /^[=+\-@]/.test(text) ? "'" + text : text;
}

function reply(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
