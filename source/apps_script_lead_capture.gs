/**
 * Rank Road — lead capture + course-progress tracking endpoint.
 *
 * IMPORTANT: set this up under a personal Gmail account, not a Google
 * Workspace/business account. Workspace domains commonly block "Anyone"
 * access to Apps Script web apps at the org level even when the deployment
 * dialog says "Anyone" — the request gets bounced with a Drive-level
 * "You need access" 403 before your code ever runs. Personal Gmail has no
 * such restriction.
 *
 * This script writes to TWO tabs in the same spreadsheet:
 *   - "Leads"    — one row per student, written once at signup.
 *   - "Progress" — one row per (student, course) they've started. Updated
 *                  in place as they study, so it always shows current %.
 *                  Both tabs are created automatically on first use if they
 *                  don't already exist — no manual header setup needed.
 *
 * Setup:
 * 1. Create a new Google Sheet (e.g. "Rank Road Leads").
 * 2. In the Sheet: Extensions -> Apps Script. Delete any boilerplate code
 *    and paste this file's contents in.
 * 3. Click Deploy -> New deployment -> type "Web app".
 *      Execute as: Me
 *      Who has access: Anyone
 * 4. Copy the resulting Web App URL.
 * 5. Paste that URL into LEAD_URL in source/foundation_template.html and
 *    source/inter_template.html, then run `python build_site.py` and
 *    redeploy the site/ folder.
 * 6. Sanity check before trusting it: `curl -i <the exec URL>` should return
 *    200 (a "doGet not found" error page is fine — that's expected, since
 *    this script only defines doPost). A 403 "Access denied"/"You need
 *    access" page means the Workspace-sharing issue above is happening again.
 * 7. After ANY code edit here, you must redeploy for the live URL to pick it
 *    up: Deploy -> Manage deployments -> pencil icon -> Version: New version
 *    -> Deploy. Saving the file alone does not update the live endpoint.
 */

function getOrCreateSheet_(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
  }
  return sheet;
}

function nowReadable_() {
  return Utilities.formatDate(new Date(), "GMT+5:30", "d MMM yyyy, h:mm a");
}

function doPost(e) {
  const data = JSON.parse(e.postData.contents);
  if (data.type === "progress") return handleProgress_(data);
  return handleLead_(data);
}

function handleLead_(data) {
  const sheet = getOrCreateSheet_("Leads", ["Name", "Phone", "Email", "Timestamp"]);
  sheet.appendRow([data.name || "", data.phone || "", data.email || "", nowReadable_()]);
  return ContentService.createTextOutput(JSON.stringify({ok: true}))
    .setMimeType(ContentService.MimeType.JSON);
}

function handleProgress_(data) {
  const sheet = getOrCreateSheet_("Progress",
    ["Name", "Phone", "Email", "Course", "Started At", "Last Active", "% Complete"]);
  const phone = data.phone || "";
  const course = data.course || "";
  const ts = nowReadable_();
  const pctText = data.pct != null ? data.pct + "%" : "";

  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][1]) === phone && String(values[i][3]) === course) {
      const row = i + 1;
      sheet.getRange(row, 6).setValue(ts);       // Last Active
      sheet.getRange(row, 7).setValue(pctText);  // % Complete
      return ContentService.createTextOutput(JSON.stringify({ok: true}))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }
  // No existing row for this student+course — this is a new journey.
  sheet.appendRow([data.name || "", phone, data.email || "", course, ts, ts, pctText]);
  return ContentService.createTextOutput(JSON.stringify({ok: true}))
    .setMimeType(ContentService.MimeType.JSON);
}
