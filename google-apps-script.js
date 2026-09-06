/**
 * Opteva intake — Google Sheets → the platform.
 *
 * Replaces the old GitHub-commit step. Onboarding submissions used to be
 * written as JSON files into the opteva-clients repository, where nothing read
 * them, so a brand someone submitted never reached the studio. This posts the
 * submission to /api/intake instead, which creates the brand directly.
 *
 * SETUP (Project Settings → Script Properties):
 *   OPTEVA_INTAKE_URL     https://demo.opteva.ai/api/intake
 *   OPTEVA_INTAKE_SECRET  must match INTAKE_SECRET in Vercel
 *
 * The secret lives in Script Properties, never in this file — anyone you share
 * the Sheet with can read the code, and that secret is the only thing between a
 * stranger and a brand in your live account.
 *
 * TRIGGER: Triggers → Add trigger → onFormSubmit → From spreadsheet → On form submit
 */

/**
 * Send one submission to Opteva.
 * Returns { ok, status, body }. Never throws, so a form submission is not lost
 * to an exception if the platform is briefly unreachable.
 */
function sendToOpteva(record) {
  var props  = PropertiesService.getScriptProperties();
  var url    = props.getProperty('OPTEVA_INTAKE_URL');
  var secret = props.getProperty('OPTEVA_INTAKE_SECRET');

  if (!url || !secret) {
    Logger.log('Not configured: set OPTEVA_INTAKE_URL and OPTEVA_INTAKE_SECRET.');
    return { ok: false, status: 0, body: 'not configured' };
  }

  var res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    headers: { 'x-opteva-secret': secret },
    payload: JSON.stringify(record),
    muteHttpExceptions: true   // read the status ourselves rather than throwing
  });

  var status = res.getResponseCode();
  var body   = res.getContentText();

  // 201 = created. 200 = already received: the endpoint is idempotent on
  // clientId, so a retry or a double submit does not create a second brand.
  if (status === 201 || status === 200) {
    Logger.log('Opteva OK: ' + body);
    // The platform reports a logo it could not save rather than failing the
    // whole submission, so surface that here instead of burying it in the body.
    if (body.indexOf('logoWarning') !== -1) {
      Logger.log('NOTE: the brand was created but its logo was not saved. See logoWarning above.');
    }
    return { ok: true, status: status, body: body };
  }

  Logger.log('Opteva FAILED (' + status + '): ' + body);
  return { ok: false, status: status, body: body };
}

/**
 * Runs on each form submission.
 * Columns are matched by header name, so adding a question to the form does not
 * break this mapping the way column positions would.
 */
function onFormSubmit(e) {
  var sheet  = e.range.getSheet();
  var header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var row    = e.range.getValues()[0];

  function cell() {
    for (var a = 0; a < arguments.length; a++) {
      var want = String(arguments[a]).trim().toLowerCase();
      for (var i = 0; i < header.length; i++) {
        if (String(header[i]).trim().toLowerCase() === want) {
          return String(row[i] || '').trim();
        }
      }
    }
    return '';
  }

  var bizName = cell('Business Name', 'bizName', 'Business', 'Company');
  if (!bizName) {
    Logger.log('Skipped row ' + e.range.getRow() + ': no business name.');
    return;
  }

  // Stable per-row id, so a rerun over the sheet updates nothing rather than
  // creating duplicate brands.
  var record = {
    clientId:   'sheet-' + slugify(bizName) + '-' + e.range.getRow(),
    bizName:    bizName,
    firstName:  cell('First Name', 'firstName'),
    lastName:   cell('Last Name', 'lastName'),
    email:      cell('Email', 'Email Address'),
    website:    cell('Website', 'URL'),
    industry:   cell('Industry', 'Business Type'),
    audience:   cell('Audience', 'Target Audience', 'Who is your audience?'),
    brandVoice: cell('Brand Voice', 'Voice', 'Tone'),
    color1:     cell('Color 1', 'Primary Color', 'Brand Color'),
    color2:     cell('Color 2', 'Secondary Color'),
    instagram:  cell('Instagram'),
    facebook:   cell('Facebook'),
    tiktok:     cell('TikTok'),
    linkedin:   cell('LinkedIn'),
    youtube:    cell('YouTube'),
    xtwitter:   cell('X / Twitter', 'X', 'Twitter'),
    pinterest:  cell('Pinterest'),
    gmb:        cell('Google Business', 'GMB'),
    goals:      cell('Goals', 'What are your goals?'),
    frequency:  cell('Frequency', 'Posting Frequency'),
    logoUrl:    cell('Logo Drive URL', 'Logo URL', 'Logo'),
    notes:      cell('Notes'),
    scrapeUrls: cell('Web URLs to Scrape')
  };

  // Canary for column drift. If the form gains a question and the header row is
  // not updated, every value after it lands under the wrong name. A colour that
  // is not a colour is the earliest cheap sign of that.
  var hexish = /^#?[0-9a-f]{3,8}$/i;
  if (record.color1 && !hexish.test(record.color1)) {
    Logger.log('WARNING: "Color 1" does not look like a colour — got "' +
               String(record.color1).slice(0, 60) + '". The sheet header row ' +
               'probably no longer matches the form, so fields are landing in ' +
               'the wrong columns. Fix the headers before trusting this row.');
    record.color1 = '';
  }
  if (record.color2 && !hexish.test(record.color2)) {
    Logger.log('WARNING: "Color 2" does not look like a colour — got "' +
               String(record.color2).slice(0, 60) + '".');
    record.color2 = '';
  }

  // Send the logo FILE, not the Drive link. Drive serves a viewer page rather
  // than the image and blocks hotlinking, so a Drive URL stored as a logo shows
  // a broken image everywhere it is used. We have Drive access here and the
  // platform does not, so reading the bytes is this script's job.
  attachLogo(record);

  var result = sendToOpteva(record);

  // Write the outcome back to the sheet, so a failure is visible to a person
  // rather than only in the execution log. Add a column headed "Opteva status".
  for (var i = 0; i < header.length; i++) {
    if (String(header[i]).trim().toLowerCase() === 'opteva status') {
      sheet.getRange(e.range.getRow(), i + 1)
           .setValue(result.ok ? 'sent ' + new Date().toISOString()
                               : 'FAILED ' + result.status);
      break;
    }
  }
}

/**
 * Read the logo out of Drive and attach it to the record as base64.
 *
 * Never throws: a logo that cannot be read must not cost someone their whole
 * submission, so every failure is logged and the record is sent without it.
 * The platform reports the same outcome back, so a missing logo is visible in
 * two places rather than silently absent.
 */
function attachLogo(record) {
  var url = record.logoUrl;
  if (!url) return;

  var id = driveFileId(url);
  if (!id) {
    // Not a Drive link. If it is a normal image URL the platform can use it
    // directly, so leave it alone.
    return;
  }

  try {
    var file = DriveApp.getFileById(id);
    var blob = file.getBlob();
    var type = String(blob.getContentType() || '');

    if (type.indexOf('image/') !== 0) {
      Logger.log('Logo skipped: "' + file.getName() + '" is ' + (type || 'an unknown type') +
                 ', not an image. Ask for a PNG, JPG or SVG.');
      return;
    }

    var bytes = blob.getBytes();
    // Kept well under the platform's own 5MB limit: base64 inflates by about a
    // third, and the request has to survive Vercel's body size cap too.
    if (bytes.length > 3 * 1024 * 1024) {
      Logger.log('Logo skipped: "' + file.getName() + '" is ' +
                 (bytes.length / 1048576).toFixed(1) + 'MB. Ask for a version under 3MB.');
      return;
    }

    record.logoBase64 = Utilities.base64Encode(bytes);
    record.logoContentType = type;
    Logger.log('Logo attached: ' + file.getName() + ' (' + type + ', ' +
               Math.round(bytes.length / 1024) + 'KB)');
  } catch (err) {
    // Most often the file is not shared with the account running this script.
    Logger.log('Logo could not be read from Drive: ' + err +
               '. Open the file in Drive and check this account can see it. ' +
               'The brand will be created without a logo.');
  }
}

/**
 * Pull the file id out of any of the Drive URL shapes people paste.
 * Returns '' for anything that is not a Drive link.
 */
function driveFileId(url) {
  var u = String(url || '');
  if (!/^https?:\/\/(drive|docs)\.google\.com\//i.test(u)) return '';

  var m = u.match(/\/file\/d\/([a-zA-Z0-9_-]+)/);       // .../file/d/ID/view
  if (m) return m[1];
  m = u.match(/[?&]id=([a-zA-Z0-9_-]+)/);                 // ...open?id=ID, uc?id=ID
  if (m) return m[1];
  m = u.match(/\/d\/([a-zA-Z0-9_-]+)/);                  // .../d/ID/...
  if (m) return m[1];

  Logger.log('Could not find a file id in the Drive link: ' + u.slice(0, 120));
  return '';
}

function slugify(s) {
  return String(s).toLowerCase().replace(/[^a-z0-9]+/g, '-')
                  .replace(/^-|-$/g, '').slice(0, 40);
}

/** Run this once by hand to check the wiring before trusting the form. */
function testIntake() {
  var out = sendToOpteva({
    clientId: 'test-' + new Date().getTime(),
    bizName:  'Intake Test — delete me',
    email:    'test@example.com',
    industry: 'Testing'
  });
  Logger.log(JSON.stringify(out));
}

/**
 * Send a row that is ALREADY in the sheet — no need to re-submit the form.
 *
 * Row 1 is the header, so the first submission is row 2. Open the sheet and
 * read the row number down the left-hand edge.
 *
 *   sendRow(5)     // send row 5
 *
 * Safe to run twice: the row number is part of the clientId, and the endpoint
 * is idempotent on it, so a second run reports the existing brand instead of
 * creating a duplicate.
 */
function sendRow(rowNumber) {
  var sheet  = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastCol = sheet.getLastColumn();

  // The editor's Run button passes no argument, which is how this is usually
  // run. Rather than fail, fall back to the row the cursor is on: click a cell
  // in the sheet, come back, press Run. Requiring a second function only helps
  // if the person managed to add it, which is one more thing to get wrong.
  if (typeof rowNumber !== 'number' || isNaN(rowNumber)) {
    rowNumber = sheet.getActiveRange().getRow();
    Logger.log('No row given, so using the row your cursor is on: row ' + rowNumber + '.');
  }
  if (rowNumber < 2) {
    Logger.log('Row 1 is the header — pass the row number of an actual submission.');
    return;
  }
  if (rowNumber > sheet.getLastRow()) {
    Logger.log('Row ' + rowNumber + ' is empty. Last row with data is ' + sheet.getLastRow() + '.');
    return;
  }

  // Reuse the form-submit path exactly, so a manual send and a real submission
  // cannot drift apart.
  onFormSubmit({ range: sheet.getRange(rowNumber, 1, 1, lastCol) });
}

/**
 * The same as running sendRow with nothing: sends the row the cursor is on.
 * Kept as a second name because it is the one that reads obviously in the
 * editor's function list.
 */
function sendSelectedRow() {
  sendRow();
}

/** Lists the rows in the sheet so you can see which number to send. */
function listRows() {
  var sheet  = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var last   = sheet.getLastRow();
  if (last < 2) { Logger.log('No submissions yet.'); return; }

  var rows = sheet.getRange(2, 1, last - 1, sheet.getLastColumn()).getValues();
  for (var r = 0; r < rows.length; r++) {
    var parts = [];
    for (var c = 0; c < header.length && c < 4; c++) {
      if (rows[r][c]) parts.push(header[c] + ': ' + rows[r][c]);
    }
    Logger.log('row ' + (r + 2) + '  —  ' + parts.join('  |  '));
  }
}


// ═══════════════════════════════════════════════════════════════════════════
// WEB APP — receives the onboarding form (optevaonboarding.vercel.app)
// ═══════════════════════════════════════════════════════════════════════════
//
// The form is a plain HTML page, not a Google Form, so it POSTs JSON here.
// Because the row is written by this script (appendRow) rather than by a Google
// Form, the "On form submit" trigger never fires for it. So this handler does
// what the trigger would have done: write the row, then send that row to
// Opteva through the SAME onFormSubmit() path a manual sendRow() uses.
//
// Deploy: Deploy > Manage deployments > (pencil) > Version: New version > Deploy.
// The form posts to the deployment URL that is hard-coded in
// opteva_onboarding/index.html (SHEET_URL). Keep that deployment updated.

var STATUS_HEADER = 'Opteva status';

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(15000);
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var data  = JSON.parse(e.postData.contents);

    // URL-safe client slug from the business name, unique within the sheet.
    var slug = (data.bizName || 'client')
      .toLowerCase()
      .replace(/[^a-z0-9\s]/g, '')
      .trim()
      .replace(/\s+/g, '-')
      .substring(0, 40);
    var existingSlugs = sheet.getLastRow() > 1
      ? sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().flat().map(String)
      : [];
    var finalSlug = slug;
    var counter = 2;
    while (existingSlugs.indexOf(finalSlug) !== -1) {
      finalSlug = slug + '-' + counter;
      counter++;
    }

    // Logo: save the uploaded file to Drive. onFormSubmit() reads the bytes
    // back out of Drive and sends the FILE to Opteva (a Drive link alone is
    // not an image anywhere it would be displayed).
    var logoUrl = '';
    if (data.logo && String(data.logo).indexOf('base64,') !== -1) {
      try {
        var base64Data = data.logo.split('base64,')[1];
        var mimeType   = data.logo.split(';')[0].split(':')[1] || 'image/png';
        var ext        = mimeType.split('/')[1] || 'png';
        var fileName   = data.logoFileName || (finalSlug + '-logo.' + ext);
        var blob       = Utilities.newBlob(Utilities.base64Decode(base64Data), mimeType, fileName);
        var folder     = getOrCreateFolder('Opteva Client Logos');
        var file       = folder.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        logoUrl = file.getUrl();
      } catch (logoErr) {
        logoUrl = '';
        Logger.log('Logo upload to Drive failed: ' + logoErr);
      }
    }

    var headers = [
      'Client ID (slug)',
      'Submitted At',
      'First Name', 'Last Name', 'Email', 'Phone',
      'Business Name', 'Website', 'Industry', 'Audience', 'Brand Voice',
      'Services', 'Goal Detail', 'Content Rules',
      'Color 1', 'Color 2',
      'Instagram', 'Facebook', 'TikTok', 'LinkedIn', 'YouTube',
      'X / Twitter', 'Pinterest', 'Google Business', 'Booking Link',
      'Web URLs to Scrape',
      'Goals', 'Current Tools', 'Posting Frequency', 'Notes',
      'Asset Source', 'Asset Folder', 'Customer Photos',
      'Logo File Name', 'Logo Drive URL',
      'Studio URL', STATUS_HEADER
    ];
    if (sheet.getLastRow() === 0) sheet.appendRow(headers);
    ensureStatusHeader(sheet);
    ensureExtraHeaders(sheet);

    var values = {
      'Client ID':          finalSlug,
      'Client ID (slug)':   finalSlug,
      'Submitted At':       new Date(),
      'First Name':         data.firstName || '',
      'Last Name':          data.lastName  || '',
      'Email':              data.email     || '',
      'Phone':              data.phone     || '',
      'Business Name':      data.bizName   || '',
      'Website':            data.website   || '',
      'Industry':           data.industry  || '',
      'Audience':           data.audience  || '',
      'Brand Voice':        data.brandVoice|| '',
      'Services':           data.services  || '',
      'Goal Detail':        data.heroService || '',
      'Content Rules':      data.offLimits || '',
      'Color 1':            data.color1    || '',
      'Color 2':            data.color2    || '',
      'Instagram':          data.instagram || '',
      'Facebook':           data.facebook  || '',
      'TikTok':             data.tiktok    || '',
      'LinkedIn':           data.linkedin  || '',
      'YouTube':            data.youtube   || '',
      'X / Twitter':        data.xtwitter  || '',
      'Pinterest':          data.pinterest || '',
      'Google Business':    data.gmb       || '',
      'Booking Link':       data.bookingLink || '',
      'Web URLs to Scrape': (data.webUrls || []).join(', '),
      'Goals':              data.goals     || '',
      'Current Tools':      data.currentTools || '',
      'Posting Frequency':  data.frequency || '',
      'Notes':              data.notes     || '',
      'Asset Source':       data.assetSource || '',
      'Asset Folder':       data.assetFolder || '',
      'Customer Photos':    data.customerPhotos || '',
      'Logo File Name':     data.logoFileName || '',
      'Logo Drive URL':     logoUrl,
      'Studio URL':         'https://demo.opteva.ai',
      'GitHub Status':      '',
      'Opteva status':      ''
    };

    // Write the row by header NAME, so the row lines up with whatever the
    // sheet's header row actually is (older rows were written by position
    // and drifted).
    var header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var row = header.map(function (h) {
      var key = String(h).trim();
      return values.hasOwnProperty(key) ? values[key] : '';
    });
    sheet.appendRow(row);
    SpreadsheetApp.flush();
    var newRow = sheet.getLastRow();

    // Send to Opteva — same code path as a manual sendRow(newRow).
    var intakeStatus = '';
    try {
      onFormSubmit({ range: sheet.getRange(newRow, 1, 1, sheet.getLastColumn()) });
      intakeStatus = readStatus(sheet, newRow);
    } catch (sendErr) {
      intakeStatus = 'FAILED ' + sendErr;
      writeStatus(sheet, newRow, intakeStatus);
      Logger.log('Intake send threw: ' + sendErr);
    }

    return ContentService
      .createTextOutput(JSON.stringify({
        success: true,
        clientId: finalSlug,
        row: newRow,
        opteva: intakeStatus
      }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

// The status column used to be headed "GitHub Status". onFormSubmit() writes
// to a column headed "Opteva status", so rename the old header if that is what
// the sheet has, and add the column if it has neither.
function ensureStatusHeader(sheet) {
  var lastCol = sheet.getLastColumn();
  if (lastCol === 0) return;
  var header = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var idxOpteva = -1, idxGithub = -1;
  for (var i = 0; i < header.length; i++) {
    var h = String(header[i]).trim().toLowerCase();
    if (h === STATUS_HEADER.toLowerCase()) idxOpteva = i;
    if (h === 'github status') idxGithub = i;
  }
  if (idxOpteva !== -1) return;
  if (idxGithub !== -1) {
    sheet.getRange(1, idxGithub + 1).setValue(STATUS_HEADER);
  } else {
    sheet.getRange(1, lastCol + 1).setValue(STATUS_HEADER);
  }
}

// Columns the form gained after the sheet was first laid out. Added at the
// end of the header row if missing, so nothing the client typed is dropped.
var EXTRA_HEADERS = ['Services', 'Goal Detail', 'Content Rules', 'Booking Link',
                     'Asset Source', 'Asset Folder', 'Customer Photos'];

function ensureExtraHeaders(sheet) {
  var lastCol = sheet.getLastColumn();
  if (lastCol === 0) return;
  var header = sheet.getRange(1, 1, 1, lastCol).getValues()[0]
                    .map(function (h) { return String(h).trim().toLowerCase(); });
  var col = lastCol;
  for (var i = 0; i < EXTRA_HEADERS.length; i++) {
    if (header.indexOf(EXTRA_HEADERS[i].toLowerCase()) === -1) {
      col++;
      sheet.getRange(1, col).setValue(EXTRA_HEADERS[i]);
    }
  }
}

function statusColumn(sheet) {
  var header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  for (var i = 0; i < header.length; i++) {
    if (String(header[i]).trim().toLowerCase() === STATUS_HEADER.toLowerCase()) return i + 1;
  }
  return 0;
}

function readStatus(sheet, rowNumber) {
  var col = statusColumn(sheet);
  return col ? String(sheet.getRange(rowNumber, col).getValue()) : '';
}

function writeStatus(sheet, rowNumber, text) {
  var col = statusColumn(sheet);
  if (col) sheet.getRange(rowNumber, col).setValue(text);
}

// Drive folder for uploaded logos.
function getOrCreateFolder(name) {
  var folders = DriveApp.getFoldersByName(name);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(name);
}

// ── GET: look up one client by slug, or export every row (used by
//        opteva_onboarding/sync_clients.py). Kept for compatibility. ──
function doGet(e) {
  try {
    var clientId = e && e.parameter ? e.parameter.client : '';
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var data  = sheet.getDataRange().getValues();
    var header = data.length ? data[0].map(function (h) { return String(h).trim(); }) : [];

    function rowObject(row, rowNumber) {
      var o = { row: rowNumber };
      for (var c = 0; c < header.length; c++) {
        if (header[c]) o[header[c]] = row[c] instanceof Date ? row[c].toISOString() : row[c];
      }
      return o;
    }

    if (!clientId) {
      var clients = [];
      for (var r = 1; r < data.length; r++) {
        var ro = rowObject(data[r], r + 1);
        clients.push({
          row: ro.row,
          clientId:   String(ro['Client ID'] || ro['Client ID (slug)'] || ''),
          firstName:  ro['First Name'],  lastName:   ro['Last Name'],
          email:      ro['Email'],       bizName:    ro['Business Name'],
          website:    ro['Website'],     industry:   ro['Industry'],
          audience:   ro['Audience'],    brandVoice: ro['Brand Voice'],
          color1:     ro['Color 1'],     color2:     ro['Color 2'],
          instagram:  ro['Instagram'],   facebook:   ro['Facebook'],
          tiktok:     ro['TikTok'],      linkedin:   ro['LinkedIn'],
          youtube:    ro['YouTube'],     xtwitter:   ro['X / Twitter'],
          pinterest:  ro['Pinterest'],   gmb:        ro['Google Business'],
          goals:      ro['Goals'],       currentTools: ro['Current Tools'],
          frequency:  ro['Posting Frequency'], logoUrl: ro['Logo Drive URL'],
          optevaStatus: ro[STATUS_HEADER] || ro['GitHub Status'] || ''
        });
      }
      return ContentService
        .createTextOutput(JSON.stringify({ debug: true, clients: clients }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim().toLowerCase() === String(clientId).trim().toLowerCase()) {
        return ContentService
          .createTextOutput(JSON.stringify(rowObject(data[i], i + 1)))
          .setMimeType(ContentService.MimeType.JSON);
      }
    }
    return ContentService
      .createTextOutput(JSON.stringify({ error: 'Client not found' }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
