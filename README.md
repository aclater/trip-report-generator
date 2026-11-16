
# 🧳 Google Form → Trip Report Generator

This Google Apps Script project automatically generates trip report documents from Google Form submissions. When a form is submitted, the script:

1. Copies a Google Docs template
2. Replaces placeholders with form data
3. Saves the new report into a monthly folder on Google Drive
4. Writes the link to the document back into the linked Google Sheet

---

## Quick start

You need a Google Form with a linked Google Sheet, a Google Docs template (see [Template Format](#-template-format-google-doc)), and a Drive folder for the reports.

1. Open the **Google Sheet** linked to your form and go to `Extensions → Apps Script`.
2. Replace the editor's contents with [`code.gs`](code.gs) from this repository.
3. Set the three constants at the top of `autoFillTripReport` to your template, folder and sheet IDs ([how to find them](#-how-to-find-your-ids)):
   ```javascript
   const TEMPLATE_ID = '';             // Your Google Docs template ID
   const ROOT_FOLDER_ID = '';          // Your Google Drive root folder ID
   const RESPONSE_SPREADSHEET_ID = ''; // Your Google Sheets response spreadsheet ID
   ```
   If your response sheet tab is not named `Form Responses 1`, also update `const SHEET_NAME = 'Form Responses 1';`. Save the script.
4. Click the **Triggers** icon (clock), then **+ Add Trigger**, and set:
   - Function: `autoFillTripReport`
   - Event Source: **From spreadsheet**
   - Event Type: **On form submit**

   Save and authorize the script.
5. Submit the form once. A report appears in a monthly folder such as `Trip Reports/2025-07/`, and its link in the sheet's `Trip Report URL` column.

---

## 📋 What You’ll Need

1. **A Google Form** collecting trip report fields
2. **A linked Google Sheet** to receive form submissions
3. **A Google Docs template** with placeholders like `{{Customer}}`
4. **A "Trip Reports" folder** in Google Drive to store documents

---

## 📁 Folder & File Setup


### 🔍 How to Find Your IDs

- **Template Doc ID**:  
  Open your Google Docs template in the browser. The URL will look like this:  
  `https://docs.google.com/document/d/1ABC1234567yourDocIDgoeshere/edit`  
  The part between `/d/` and `/edit` is your **Template Doc ID**.

- **Root Folder ID**:
  Open the destination folder in Google Drive. The URL will look like this:
  `https://drive.google.com/drive/folders/1XYZ123folderIDgoeshere`
  The part after `/folders/` is your **Root Folder ID**.

- **Response Spreadsheet ID**:
  Open the Google Sheet linked to your form. The URL will look like this:
  `https://docs.google.com/spreadsheets/d/1ABC1234567yourSheetIDgoeshere/edit`
  The part between `/d/` and `/edit` is your **Response Spreadsheet ID**.


| Resource | Value |
|---------|-------|
| **Template Doc ID** | `<YOUR_TEMPLATE_DOC_ID>` |
| **Root Drive Folder ID** | `<YOUR_ROOT_FOLDER_ID>` |
| **Response Spreadsheet ID** | `<YOUR_RESPONSE_SPREADSHEET_ID>` |
| **Response Sheet Name** | `Form Responses 1` (default when linking a form) |

---

## 📝 Template Format (Google Doc)

Use this structure in your Google Docs template:

```
Customer / Partner Connection Document

Submitter: {{SubmitterEmail}}
Topic: {{MeetingTopic}}

Date : {{Date}}
Region: {{Region}}
Public Sector: {{PublicSector}}

Customer: {{Customer}}
After-Action Summary & Outcomes

Summary
{{MeetingPurpose}}

Key Customer Insights
{{CustomerInsights}}

Action Items:
{{ActionItems}}

Activity Notes & Details 
Customer Attendees
{{CustomersPresent}}
Red Hat Attendees
{{RedHatters}}
Partner Attendees
{{Partners}}
Notes
{{Observations}}
```

---

## ⚙️ Form Field Mapping

| Template Placeholder | Form Question |
|----------------------|----------------|
| `{{SubmitterEmail}}` | Email Address |
| `{{MeetingTopic}}` | Briefly Describe the meeting topic |
| `{{Date}}` | When was your customer interaction? |
| `{{Region}}` | What Region was your trip to? |
| `{{PublicSector}}` | Was this a Public Sector Customer |
| `{{Customer}}` | Customer (company, agency name, etc) |
| `{{MeetingPurpose}}` | Briefly summarize the purpose of the meeting. |
| `{{CustomerInsights}}` | What were your Key Customer insights? |
| `{{ActionItems}}` | What Action Items came out of the meeting? |
| `{{CustomersPresent}}` | What customers were in attendance? |
| `{{RedHatters}}` | What Red Hatters were in attendance? |
| `{{Partners}}` | What Partners were in attendance? |
| `{{Observations}}` | Any additional notes. |

---

## ✅ What Happens After Submission

- The form submission triggers the script.
- A Google Doc is created in a folder like `Trip Reports/2025-07/`
- Placeholders are replaced with the submission content.
- A link to the created document is written into the first `Trip Report URL` column of the sheet.

---

## 🛠 Troubleshooting

- **Missing values in report?**
  - Ensure field names in the form exactly match what's in the mapping.
- **No document created?**
  - Check `Apps Script > Executions` tab for error logs.
- **Multiple `Trip Report URL` columns?**
  - Manually delete all but one. The script always writes to the first matching column.
- **Script times out or fails to lock?**
  - Check `Apps Script > Executions` for a `LockService` timeout error. This can happen if multiple submissions arrive simultaneously — re-submitting the form should resolve it.

---

## 📬 Future Ideas

- Auto-email the generated document to the submitter
- Generate AI summaries using Gemini or PaLM (optionally)
- Summarize all monthly reports into a single doc
- Sync metadata to CRM or Jira

---

## 🔒 Notes

- This script runs under the permissions of the user who owns the spreadsheet.
- Be mindful of API quotas if used heavily or in a shared environment.

---

## 📎 Related Files

- Sample Google Form (ask user to create)
- Sample Google Doc template (see above)
- Script: `autoFillTripReport()` (included in this repo)

---

## 🧠 Questions?

Drop issues into this repo or contact the script maintainer.
