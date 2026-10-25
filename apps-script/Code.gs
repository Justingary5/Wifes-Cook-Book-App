// Receives new meals and sides from the cookbook's "+ Add" form and adds them
// as rows in a "Submissions" tab, which Claude reads to write the real cookbook page.
//
// Setup (once): in the Google Sheet, Extensions > Apps Script, paste this file,
// then Deploy > New deployment > Web app, "Execute as: Me", "Who has access: Anyone".
// Put the web app URL into SUBMIT_URL in index.html.

const HEADERS = ["Submitted", "Type", "Name", "What went in it", "Serves", "Lasts (days)",
  "Goes well with", "Notes", "Status"];

function doPost(e) {
  const d = JSON.parse(e.postData.contents);
  const clip = (v, n) => String(v || "").slice(0, n);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Submissions");
  if (!sheet) {
    sheet = ss.insertSheet("Submissions");
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
  }
  sheet.appendRow([new Date(), clip(d.type, 10), clip(d.name, 80), clip(d.ingredients, 2000), clip(d.serves, 5),
    clip(d.lasts, 5), clip(d.goes, 200), clip(d.notes, 400), "New"]);
  return ContentService.createTextOutput("ok");
}
