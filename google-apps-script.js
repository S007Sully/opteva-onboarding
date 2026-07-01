// OPTEVA ONBOARDING — Google Apps Script
// ----------------------------------------
// 1. Open a new Google Sheet
// 2. Click Extensions > Apps Script
// 3. Delete everything and paste this entire file
// 4. Click Save (floppy disk icon)
// 5. Click Deploy > New deployment
//    - Type: Web app
//    - Execute as: Me
//    - Who has access: Anyone
// 6. Click Deploy, authorize when prompted
// 7. Copy the Web app URL
// 8. Paste it into index.html where it says PASTE_YOUR_APPS_SCRIPT_URL_HERE

function doPost(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var data = JSON.parse(e.postData.contents);

    // Add header row if sheet is empty
    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        'Submitted At',
        'First Name', 'Last Name', 'Email', 'Phone',
        'Business Name', 'Website', 'Industry', 'Audience', 'Brand Voice',
        'Color 1', 'Color 2',
        'Instagram', 'Facebook', 'TikTok', 'LinkedIn', 'YouTube',
        'X / Twitter', 'Pinterest', 'Google Business',
        'Web URLs to Scrape',
        'Goals', 'Posting Frequency', 'Notes',
        'Logo File Name'
      ]);
    }

    sheet.appendRow([
      new Date(),
      data.firstName || '',
      data.lastName || '',
      data.email || '',
      data.phone || '',
      data.bizName || '',
      data.website || '',
      data.industry || '',
      data.audience || '',
      data.brandVoice || '',
      data.color1 || '',
      data.color2 || '',
      data.instagram || '',
      data.facebook || '',
      data.tiktok || '',
      data.linkedin || '',
      data.youtube || '',
      data.xtwitter || '',
      data.pinterest || '',
      data.gmb || '',
      (data.webUrls || []).join(', '),
      data.goals || '',
      data.frequency || '',
      data.notes || '',
      data.logoFileName || ''
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
