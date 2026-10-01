// Computer skills check: automatic results
// Paste into a Google Sheet: Extensions > Apps Script. Then Deploy > New deployment > Web app
// (Execute as: Me, Who has access: Anyone). Copy the /exec link into the test's teacher setup page.

const ALLOWED_DOMAIN = "ncclondon.ac.uk"; // results are only emailed to addresses ending in this
const SHEET_NAME = "Results";

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const d = JSON.parse(e.postData.contents);
    const safe = v => (typeof v === "string" && /^[=+\-@]/.test(v)) ? "'" + v : v;
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
    if (sh.getLastRow() === 0) {
      sh.appendRow(["Received", "Teacher", "Name", "Student number", "Level / class", "Date", "Minutes",
        "Score", "Out of", "Percent", "Result", "Sections", "Questions", "Notes", "Results code"]);
      sh.setFrozenRows(1);
    }
    sh.appendRow([new Date(), d.teacher, d.name, d.id, d.cls, d.date, d.mins, d.got, d.of, d.pct,
      d.band, d.sections, d.tasks, d.notes, d.code].map(safe));

    let to = String(d.teacherEmail || "").trim().toLowerCase();
    if (!to.endsWith("@" + ALLOWED_DOMAIN)) to = Session.getEffectiveUser().getEmail();
    MailApp.sendEmail({
      to: to,
      subject: "Computer skills check: " + d.name + " (" + d.pct + "%, " + d.band + ")",
      body: d.text
    });
    return out({ ok: true });
  } catch (err) {
    return out({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return out({ ok: true, message: "Results link is working." });
}

function out(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
