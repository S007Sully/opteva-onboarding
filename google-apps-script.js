// OPTEVA ONBOARDING — Google Apps Script
// ----------------------------------------
// 1. Open your Opteva Onboarding Google Sheet
// 2. Click Extensions > Apps Script
// 3. Replace ALL existing code with this file
// 4. Click Save
// 5. Click Deploy > Manage deployments > edit the existing deployment > New version > Deploy
//    (or Deploy > New deployment if first time)
//    - Type: Web app
//    - Execute as: Me
//    - Who has access: Anyone
// 6. Copy the Web app URL — paste it into both:
//    - opteva_onboarding/index.html  (SHEET_URL constant)
//    - opteva_clients/index.html     (SHEET_API constant)

// ── Receive new onboarding submission ──
function doPost(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var data = JSON.parse(e.postData.contents);

    // Generate a URL-safe client slug from business name
    var slug = (data.bizName || 'client')
      .toLowerCase()
      .replace(/[^a-z0-9\s]/g, '')
      .trim()
      .replace(/\s+/g, '-')
      .substring(0, 40);

    // Make slug unique if it already exists
    var existingSlugs = sheet.getRange(2, 1, Math.max(sheet.getLastRow() - 1, 1), 1).getValues().flat();
    var finalSlug = slug;
    var counter = 2;
    while (existingSlugs.indexOf(finalSlug) !== -1) {
      finalSlug = slug + '-' + counter;
      counter++;
    }

    // Save logo to Google Drive if provided
    var logoUrl = '';
    if (data.logo && data.logo.indexOf('base64,') !== -1) {
      try {
        var base64Data = data.logo.split('base64,')[1];
        var mimeType   = data.logo.split(';')[0].split(':')[1] || 'image/png';
        var ext        = mimeType.split('/')[1] || 'png';
        var fileName   = (data.logoFileName || (finalSlug + '-logo.' + ext));
        var blob       = Utilities.newBlob(Utilities.base64Decode(base64Data), mimeType, fileName);
        var folder     = getOrCreateFolder('Opteva Client Logos');
        var file       = folder.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        logoUrl = file.getUrl();
      } catch (logoErr) {
        logoUrl = 'Upload failed: ' + logoErr.toString();
      }
    }

    // Add header row if sheet is empty
    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        'Client ID (slug)',
        'Submitted At',
        'First Name', 'Last Name', 'Email', 'Phone',
        'Business Name', 'Website', 'Industry', 'Audience', 'Brand Voice',
        'Color 1', 'Color 2',
        'Instagram', 'Facebook', 'TikTok', 'LinkedIn', 'YouTube',
        'X / Twitter', 'Pinterest', 'Google Business',
        'Web URLs to Scrape',
        'Goals', 'Posting Frequency', 'Notes',
        'Logo File Name', 'Logo Drive URL'
      ]);
    }

    sheet.appendRow([
      finalSlug,
      new Date(),
      data.firstName || '',
      data.lastName  || '',
      data.email     || '',
      data.phone     || '',
      data.bizName   || '',
      data.website   || '',
      data.industry  || '',
      data.audience  || '',
      data.brandVoice|| '',
      data.color1    || '',
      data.color2    || '',
      data.instagram || '',
      data.facebook  || '',
      data.tiktok    || '',
      data.linkedin  || '',
      data.youtube   || '',
      data.xtwitter  || '',
      data.pinterest || '',
      data.gmb       || '',
      (data.webUrls || []).join(', '),
      data.goals     || '',
      data.frequency || '',
      data.notes     || '',
      data.logoFileName || '',
      logoUrl
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ success: true, clientId: finalSlug }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── Helper: get or create a Drive folder by name ──
function getOrCreateFolder(name) {
  var folders = DriveApp.getFoldersByName(name);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(name);
}

// ── Look up a client by slug (used by the client content tool) ──
function doGet(e) {
  var headers = {
    'Access-Control-Allow-Origin': '*',
    'Content-Type': 'application/json'
  };

  try {
    var clientId = e.parameter.client;

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var data  = sheet.getDataRange().getValues();

    // Debug: return all slugs when no client param provided
    if (!clientId) {
      var slugs = [];
      for (var r = 1; r < data.length; r++) {
        slugs.push({ row: r + 1, col_A: String(data[r][0]), col_B: String(data[r][1]).substring(0,20) });
      }
      return ContentService
        .createTextOutput(JSON.stringify({ debug: true, rows: slugs }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // Row 0 = headers, slug is column 0 (case-insensitive, trimmed)
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim().toLowerCase() === clientId.trim().toLowerCase()) {
        var row = data[i];
        var client = {
          clientId:   row[0],
          firstName:  row[2],
          lastName:   row[3],
          email:      row[4],
          bizName:    row[6],
          website:    row[7],
          industry:   row[8],
          audience:   row[9],
          brandVoice: row[10],
          color1:     row[11],
          color2:     row[12],
          instagram:  row[13],
          facebook:   row[14],
          tiktok:     row[15],
          linkedin:   row[16],
          youtube:    row[17],
          xtwitter:   row[18],
          pinterest:  row[19],
          gmb:        row[20],
          goals:      row[22],
          frequency:  row[23]
        };
        return ContentService
          .createTextOutput(JSON.stringify(client))
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
