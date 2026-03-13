function autoFillTripReport(e) {
  const formResponse = e.namedValues || {};

  const TEMPLATE_ID = '';  // Fill in your Google Doc template ID
  const ROOT_FOLDER_ID = '';  // Fill in your Google Drive root folder ID
  const RESPONSE_SPREADSHEET_ID = '';  // Fill in your Google Sheets response spreadsheet ID
  const SHEET_NAME = 'Form Responses 1';

  try {
    Logger.log("RUNNING PROJECT: " + ScriptApp.getScriptId());

    const ss = SpreadsheetApp.openById(RESPONSE_SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) throw new Error(`Sheet not found: ${SHEET_NAME}`);

    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      Utilities.sleep(400);

      // Timestamp used to match the exact response row
      const submittedTs = getSubmittedTimestamp_(e);
      Logger.log("DEBUG submittedTs: " + submittedTs);

      const submittedRow = findResponseRowByTimestamp_(sheet, submittedTs);
      Logger.log("DEBUG submittedRow: " + submittedRow);

      // Folder/month
      const interactionDateStr = (formResponse["When was your customer interaction?"]?.[0] || '').toString();
      const dateForFolder = new Date(interactionDateStr || submittedTs || new Date());
      const monthFolderName = Utilities.formatDate(dateForFolder, Session.getScriptTimeZone(), 'yyyy-MM');

      const rootFolder = DriveApp.getFolderById(ROOT_FOLDER_ID);
      const monthlyFolder = getOrCreateFolder(monthFolderName, rootFolder);

      // Filename values
      const customerNameRaw = (formResponse["Customer (company, agency name, etc)"]?.[0] || 'Unknown').toString();
      const regionRaw = (formResponse["What Region was your trip to?"]?.[0] || '').toString();

      // If region doesn't come through for any reason, fall back to the sheet cell
      const regionFromSheet = getCellByHeader_(sheet, submittedRow, "What Region was your trip to?");
      const regionFinal = sanitizeFilePart_(regionRaw.trim() || String(regionFromSheet || "Unknown Region"));

      const customerFinal = sanitizeFilePart_(customerNameRaw);

      const docName = `Trip Report - ${customerFinal} - ${regionFinal} - ${monthFolderName}`;
      Logger.log(`DEBUG docName: "${docName}"`);

      // Create doc copy (never edits template)
      const copy = DriveApp.getFileById(TEMPLATE_ID).makeCopy(docName, monthlyFolder);
      copy.setName(docName);

      Logger.log("CREATED FILE ID: " + copy.getId());
      Logger.log("CREATED FILE NAME: " + copy.getName());

      // Fill placeholders IN THE COPY
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
        let value = (formResponse[fieldLabel]?.[0] || '').toString();

        // Fallbacks for key fields
        if (!value) {
          if (fieldLabel === "What Region was your trip to?") value = regionFinal;
          if (fieldLabel === "Customer (company, agency name, etc)") value = customerNameRaw;
          if (fieldLabel === "When was your customer interaction?") value = interactionDateStr || '';
        }

        body.replaceText(placeholder, value);
      }

      doc.saveAndClose();

      // Write URL back to the matched row
      const urlCol = ensureTripReportUrlColumn_(sheet);
      sheet.getRange(submittedRow, urlCol).setValue(copy.getUrl());
      SpreadsheetApp.flush();

      Logger.log(`✅ URL written row ${submittedRow}, col ${urlCol}: ${copy.getUrl()}`);
    } finally {
      lock.releaseLock();
    }
  } catch (error) {
    Logger.log(`❌ Error: ${error && error.stack ? error.stack : error}`);
    throw error;
  }
}

function getOrCreateFolder(name, parentFolder) {
  const folders = parentFolder.getFoldersByName(name);
  return folders.hasNext() ? folders.next() : parentFolder.createFolder(name);
}

function getSubmittedTimestamp_(e) {
  // Spreadsheet onFormSubmit trigger
  if (e && e.values && e.values[0]) {
    const d = new Date(e.values[0]);
    if (!isNaN(d.getTime())) return d;
  }
  // Form onFormSubmit trigger
  if (e && e.response && typeof e.response.getTimestamp === 'function') {
    return e.response.getTimestamp();
  }
  return new Date();
}

function findResponseRowByTimestamp_(sheet, ts) {
  const target = ts.getTime();
  const toleranceMs = 2 * 60 * 1000; // 2 minutes

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) throw new Error("No responses to match in sheet.");

  const colA = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (let i = colA.length - 1; i >= 0; i--) {
    const v = colA[i][0];
    let t = null;

    if (v instanceof Date) t = v.getTime();
    else if (v) {
      const d = new Date(v);
      if (!isNaN(d.getTime())) t = d.getTime();
    }

    if (t !== null && Math.abs(t - target) <= toleranceMs) {
      return i + 2;
    }
  }

  Logger.log("⚠️ Timestamp row not found; falling back to lastRow=" + lastRow);
  return lastRow;
}

function ensureTripReportUrlColumn_(sheet) {
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

  let urlCol = -1;
  for (let i = 0; i < headers.length; i++) {
    if (String(headers[i]).trim().startsWith("Trip Report URL")) {
      urlCol = i + 1;
      break;
    }
  }

  if (urlCol === -1) {
    urlCol = headers.length + 1;
    sheet.getRange(1, urlCol).setValue("Trip Report URL");
    SpreadsheetApp.flush();
  }
  return urlCol;
}

function getCellByHeader_(sheet, row, headerText) {
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const idx = headers.findIndex(h => String(h).trim() === String(headerText).trim());
  if (idx === -1) return '';
  return sheet.getRange(row, idx + 1).getValue();
}

function sanitizeFilePart_(s) {
  return String(s)
    .replace(/[\/\\\?\%\*\:\|\\"\<\>]/g, '-')
    .replace(/\s+/g, ' ')
    .trim();
}
