function autoFillTripReport(e) {
  const formResponse = e.namedValues;

  const TEMPLATE_ID = ;     
  const ROOT_FOLDER_ID = ;     
  const SHEET_NAME = 'Form Responses 1';

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);

  try {
    const interactionDateStr = formResponse["When was your customer interaction?"]?.[0] || '';
    const date = new Date(interactionDateStr || e.values[0]); // fallback to timestamp
    const monthFolderName = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM');

    const rootFolder = DriveApp.getFolderById(ROOT_FOLDER_ID);
    const monthlyFolder = getOrCreateFolder(monthFolderName, rootFolder);

    const customerName = formResponse["Customer (company, agency name, etc)"]?.[0] || 'Unknown';
    const docName = `Trip Report - ${customerName} - ${monthFolderName}`;

    const copy = DriveApp.getFileById(TEMPLATE_ID).makeCopy(docName, monthlyFolder);
    const doc = DocumentApp.openById(copy.getId());
    const body = doc.getBody();

    const fieldMap = {
      '{{SubmitterEmail}}': 'Email Address',
      '{{MeetingTopic}}': 'Briefly Describe the meeting topic',
      '{{Date}}': 'When was your customer interaction?',
      '{{Region}}': 'What Region was your trip to?',
      '{{PublicSector}}': 'Was this a Public Sector Customer',
      '{{Customer}}': 'Customer (company, agency name, etc)',
      '{{MeetingPurpose}}': 'Briefly summarize the purpose of the meeting.',
      '{{CustomerInsights}}': 'What were your Key Customer insights?',
      '{{ActionItems}}': 'What Action Items came out of the meeting?',
      '{{CustomersPresent}}': 'What customers were in attendance?',
      '{{RedHatters}}': 'What Red Hatters were in attendance?',
      '{{Partners}}': 'What Partners were in attendance?',
      '{{Observations}}': 'Any additional notes.'
    };

    for (const placeholder in fieldMap) {
      const fieldLabel = fieldMap[placeholder];
      const value = formResponse[fieldLabel]?.[0] || '';
      body.replaceText(placeholder, value);
    }

    doc.saveAndClose();

    writeBackTripReportURL(e, copy.getUrl(), SHEET_NAME);
    Logger.log(`✅ Document created and URL written: ${copy.getUrl()}`);

  } catch (error) {
    Logger.log(`❌ Error creating trip report: ${error}`);
  }
}

function getOrCreateFolder(name, parentFolder) {
  const folders = parentFolder.getFoldersByName(name);
  return folders.hasNext() ? folders.next() : parentFolder.createFolder(name);
}

function writeBackTripReportURL(e, docUrl, sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  // Get the first "Trip Report URL" column or create it
  let urlCol = -1;
  for (let i = 0; i < headers.length; i++) {
    if (headers[i].toString().trim().startsWith("Trip Report URL")) {
      urlCol = i + 1;
      break;
    }
  }

  if (urlCol === -1) {
    urlCol = headers.length + 1;
    sheet.getRange(1, urlCol).setValue("Trip Report URL");
  }

  const submittedTimestamp = e.values[0];
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === submittedTimestamp) {
      sheet.getRange(i + 1, urlCol).setValue(docUrl);
      return;
    }
  }

  Logger.log(`⚠️ Timestamp not matched: ${submittedTimestamp}`);
}
