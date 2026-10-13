/**
 * Pick of the Week — saves "how each bet was sorted" choices into a tab of this spreadsheet,
 * so a choice you make on the website is saved for EVERYONE (not just your own browser).
 *
 * SETUP (once, about 5 minutes):
 *
 *  1. In your Google Sheet: Extensions > Apps Script.
 *     Click the "+" next to "Files" > Script, name it  UnitOverrides  and paste ALL of this in.
 *     (Your GroupMe script can stay exactly as it is.)
 *
 *  2. At the bottom of the sheet, add a new tab named exactly:  Unit Overrides
 *     and type these headers in row 1:   Key | Bucket | Name | Week | Pick | Updated
 *
 *  3. In Apps Script click  Deploy > New deployment > (gear icon) > Web app.
 *        Execute as:      Me
 *        Who has access:  Anyone
 *     Click Deploy, approve the permissions, and COPY the "Web app URL".
 *     (Paste that URL into a browser tab: it should say "Unit overrides endpoint is live.")
 *
 *  4. Back in the sheet: File > Share > Publish to web. Choose the "Unit Overrides" tab and
 *     "Comma-separated values (.csv)", click Publish, and COPY that link.
 *
 *  5. In index.html, paste the two links into the CONFIG near the top:
 *        UNIT_OVERRIDES_WRITE_URL: "<the Web app URL from step 3>",
 *        UNIT_OVERRIDES_CSV_URL:   "<the published CSV link from step 4>",
 *
 *  6. Open your site with  ?admin=1  on the end of the address (once). The dropdowns appear
 *     in Units > "How each bet was sorted".
 *
 * Note: this script only ever touches the "Unit Overrides" tab. Anyone who found the Web app URL
 * could change which bucket a bet is sorted into (nothing else), so don't post that URL publicly.
 */

const OVERRIDES_TAB = 'Unit Overrides';
const ALLOWED_BUCKETS = ['NFL Spreads', 'NFL Totals', 'CFB Spreads', 'CFB Totals', 'Player Props', 'Other', ''];

// open the Web app URL in a browser to check it's live
function doGet() {
  return ContentService.createTextOutput('Unit overrides endpoint is live.');
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const d = JSON.parse(e.postData.contents);       // {key, bucket, name, week, pick}
    const key = String(d.key || '').trim();
    const bucket = String(d.bucket === undefined ? '' : d.bucket);
    if (!key || ALLOWED_BUCKETS.indexOf(bucket) === -1) throw new Error('bad request');

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sh = ss.getSheetByName(OVERRIDES_TAB);
    if (!sh) {
      sh = ss.insertSheet(OVERRIDES_TAB);
      sh.appendRow(['Key', 'Bucket', 'Name', 'Week', 'Pick', 'Updated']);
    }

    const last = sh.getLastRow();
    const keys = last > 1 ? sh.getRange(2, 1, last - 1, 1).getValues().map(r => String(r[0])) : [];
    const at = keys.indexOf(key);

    if (bucket === '') {
      // "Auto" = remove the override and let the site guess again
      if (at !== -1) sh.deleteRow(at + 2);
    } else {
      const row = [key, bucket, d.name || '', d.week || '', d.pick || '', new Date()];
      if (at === -1) sh.appendRow(row);
      else sh.getRange(at + 2, 1, 1, 6).setValues([row]);
    }
    return ContentService.createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}
