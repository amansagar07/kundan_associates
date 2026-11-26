//1-Click Bulk Client Folder Generator (50 Clients × 8 Services)
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('⚡ Automation⚡')
    .addItem('1. Generate Client Folders & Grant View Access', 'createClientFolderStructure')
    .addToUi();
}

function createClientFolderStructure() {
  const PARENT_FOLDER_ID = '1np63FeqV6ogmHcekSIwNKbHcpe7jELva'.trim();

  const SERVICE_FOLDERS = [
    '01_Name_of_the_Company',
    '02_GST',
    '03_Income_Tax',
    '04_Labour_Law_Compliances',
    '05_MSME_DPIIT_FSSAI',
    '06_Trademark_Registrations',
    '07_Digital_Signature_Certificate',
    '08_Miscellaneous_Registration'
  ];

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('_client_master_');
  
  if (!sheet) {
    SpreadsheetApp.getUi().alert('Error: Please create a tab named "_client_master_" first.');
    return;
  }

  let parentFolder;
  try {
    parentFolder = DriveApp.getFolderById(PARENT_FOLDER_ID);
  } catch (e) {
    SpreadsheetApp.getUi().alert(
      'Cannot access folder.\n\n' +
      'Active User: ' + Session.getActiveUser().getEmail() + '\n' +
      'Technical Details: ' + e.message
    );
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert('No client names found starting at row 2 in _client_master_.');
    return;
  }

  const dataRange = sheet.getRange(2, 1, lastRow - 1, 4);
  const data = dataRange.getValues();
  let processedCount = 0;

  for (let i = 0; i < data.length; i++) {
    const clientName = data[i][0].toString().trim();
    const spocEmail = data[i][1].toString().trim();
    const existingUrl = data[i][2].toString().trim();
    const status = data[i][3].toString().trim();

    if (!clientName) continue;

    let clientFolder;
    let clientFolderUrl = existingUrl;

    // 1. Create folder if not already created
    if (status !== 'Created') {
      clientFolder = parentFolder.createFolder(clientName);
      clientFolderUrl = clientFolder.getUrl();

      for (let j = 0; j < SERVICE_FOLDERS.length; j++) {
        clientFolder.createFolder(SERVICE_FOLDERS[j]);
      }

      sheet.getRange(i + 2, 3).setValue(clientFolderUrl);
      sheet.getRange(i + 2, 4).setValue('Created');
    } else if (existingUrl) {
      // Extract Folder ID from existing URL so we can still apply View Access to Client 1, 2, 3
      const match = existingUrl.match(/[-\w]{25,}/);
      if (match) {
        clientFolder = DriveApp.getFolderById(match[0]);
      }
    }

    // 2. Grant View Access to Client SPOC Email (silent, no spam email sent yet)
    if (clientFolder && spocEmail && spocEmail.includes('@')) {
      try {
        Drive.Permissions.insert(
          {
            role: 'reader',
            type: 'user',
            value: spocEmail
          },
          clientFolder.getId(),
          {
            supportsAllDrives: true,
            sendNotificationEmails: false
          }
        );
      } catch (permErr) {
        // Fallback to standard DriveApp method if Drive API v2 isn't enabled
        try {
          clientFolder.addViewer(spocEmail);
        } catch (fallbackErr) {
          SpreadsheetApp.getUi().alert(
            'Folder created for ' + clientName + ', but could not share with ' + spocEmail + '.\n' +
            'Reason: ' + fallbackErr.message + '\n\n' +
            'Check Shared Drive Settings > "Allow people outside Kundan & Associates to access files".'
          );
        }
      }
    }

    SpreadsheetApp.flush();
    processedCount++;
  }

  SpreadsheetApp.getUi().alert('Done! Processed folders & View permissions for ' + processedCount + ' clients.');
}



//Tamper-Proof Automated Time Tracking
const TIME_TRACKER_CONFIG = {
  SHEET_NAME: '01_Daily Task Tracker',
  END_DATE_COL: 15,      // Col O
  TASK_STATUS_COL: 19,   // Col S
  HOURS_SPENT_COL: 21,   // Col U
  STARTED_AT_COL: 28,    // Col AB (hidden audit timestamp)
  COMPLETED_AT_COL: 29   // Col AC (hidden audit timestamp)
};

function setupTimeTrackingColumns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(TIME_TRACKER_CONFIG.SHEET_NAME);
  if (!sheet) {
    SpreadsheetApp.getUi().alert('Error: Sheet "' + TIME_TRACKER_CONFIG.SHEET_NAME + '" not found.');
    return;
  }

  sheet.getRange(1, TIME_TRACKER_CONFIG.STARTED_AT_COL).setValue('_Started_At');
  sheet.getRange(1, TIME_TRACKER_CONFIG.COMPLETED_AT_COL).setValue('_Completed_At');

  const maxRows = sheet.getMaxRows();
  if (maxRows > 1) {
    sheet.getRange(2, TIME_TRACKER_CONFIG.STARTED_AT_COL, maxRows - 1, 2)
      .setNumberFormat('dd-mm-yyyy hh:mm:ss');
    sheet.getRange(2, TIME_TRACKER_CONFIG.HOURS_SPENT_COL, maxRows - 1, 1)
      .setNumberFormat('0.00');
  }
}

function handleTimeTrackingEdit(e) {
  if (!e || !e.range) return;

  const sheet = e.range.getSheet();
  if (sheet.getName() !== TIME_TRACKER_CONFIG.SHEET_NAME) return;

  const row = e.range.getRow();
  const col = e.range.getColumn();

  if (row < 2 || col !== TIME_TRACKER_CONFIG.TASK_STATUS_COL || e.range.getNumRows() > 1) return;

  const newStatus = (e.value || '').toString().trim();
  if (!newStatus) return;

  const now = new Date();
  const startedCell = sheet.getRange(row, TIME_TRACKER_CONFIG.STARTED_AT_COL);
  const completedCell = sheet.getRange(row, TIME_TRACKER_CONFIG.COMPLETED_AT_COL);
  const hoursCell = sheet.getRange(row, TIME_TRACKER_CONFIG.HOURS_SPENT_COL);
  const endDateCell = sheet.getRange(row, TIME_TRACKER_CONFIG.END_DATE_COL);

  if (newStatus === 'In Progress') {
    if (!startedCell.getValue()) {
      startedCell.setValue(now);
    }
    if (completedCell.getValue()) {
      completedCell.clearContent();
      hoursCell.clearContent();
      endDateCell.clearContent();
    }
  } else if (newStatus === 'Under Review') {
    if (!startedCell.getValue()) {
      startedCell.setValue(now);
    }
    if (completedCell.getValue()) {
      completedCell.clearContent();
      hoursCell.clearContent();
      endDateCell.clearContent();
    }
  } else if (newStatus === 'Completed') {
    let startTime = startedCell.getValue();
    if (!startTime) {
      startTime = now;
      startedCell.setValue(startTime);
    }

    completedCell.setValue(now);
    if (!endDateCell.getValue()) {
      endDateCell.setValue(now);
    }

    const startObj = new Date(startTime);
    const elapsedHours = Math.max(0, (now.getTime() - startObj.getTime()) / (1000 * 60 * 60));
    hoursCell.setValue(Number(elapsedHours.toFixed(2)));
  }
}


//Automated Email Alerts (Assignment & Review)
const EMAIL_ALERT_CONFIG = {
  TRACKER_SHEET_NAME: '01_Daily Task Tracker',
  EMPLOYEE_SHEET_NAME: '_employee_master_',
  DEFAULT_REVIEWER_NAME: 'Kundan Kumar',
  FALLBACK_REVIEWER_EMAIL: 'pcskundan@cskundan.com',
  COLS: {
    TASK_ID: 1,       // Col A
    CLIENT_NAME: 2,   // Col B
    COMPANY_POC: 3,   // Col C
    SERVICE: 4,       // Col D
    DELIVERABLE: 5,   // Col E
    ASSIGNED_TO: 6,   // Col F
    ASSIST: 7,        // Col G
    REVIEWER: 8,      // Col H
    REVIEW_STATUS: 9, // Col I (Checkbox)
    DUE_DATE: 16,     // Col P
    TASK_STATUS: 19,  // Col S
    PROJECT_LINK: 20  // Col T
  }
};

function handleEmailAlertsEdit(e) {
  if (!e || !e.range) return;

  const sheet = e.range.getSheet();
  if (sheet.getName() !== EMAIL_ALERT_CONFIG.TRACKER_SHEET_NAME) return;

  const row = e.range.getRow();
  const col = e.range.getColumn();
  if (row < 2 || e.range.getNumRows() > 1) return;

  const c = EMAIL_ALERT_CONFIG.COLS;
  const rowVals = sheet.getRange(row, 1, 1, c.PROJECT_LINK).getValues()[0];

  // Auto-default Reviewer (Col H) to Kundan Kumar if blank on an active row edit
  let reviewer = (rowVals[c.REVIEWER - 1] || '').toString().trim();
  const hasTaskContext =
    rowVals[c.CLIENT_NAME - 1] ||
    rowVals[c.SERVICE - 1] ||
    rowVals[c.DELIVERABLE - 1] ||
    rowVals[c.ASSIGNED_TO - 1];

  if (!reviewer && hasTaskContext && col !== c.REVIEWER) {
    reviewer = EMAIL_ALERT_CONFIG.DEFAULT_REVIEWER_NAME;
    sheet.getRange(row, c.REVIEWER).setValue(reviewer);
  }

  const taskId = rowVals[c.TASK_ID - 1] || ('Row ' + row);
  const clientName = rowVals[c.CLIENT_NAME - 1] || 'Unspecified Client';
  const poc = rowVals[c.COMPANY_POC - 1] || '-';
  const service = rowVals[c.SERVICE - 1] || '-';
  const deliverable = rowVals[c.DELIVERABLE - 1] || '-';
  const assignedTo = (rowVals[c.ASSIGNED_TO - 1] || '').toString().trim();
  const assist = (rowVals[c.ASSIST - 1] || '').toString().trim();
  const dueDate = formatAlertDate_(rowVals[c.DUE_DATE - 1]);
  const projectLink = (rowVals[c.PROJECT_LINK - 1] || '').toString().trim();

  const empMap = getActiveEmployeeEmailMap_();

  // 1. Primary Assignee (Col F) or Assist (Col G) assigned/changed
  const newVal = (e.value || '').toString().trim();
  if ((col === c.ASSIGNED_TO || col === c.ASSIST) && newVal) {
    const targetEmail = empMap[newVal.toLowerCase()];
    if (targetEmail) {
      const roleTitle = col === c.ASSIGNED_TO ? 'Primary Assignee' : 'Supporting (Assist)';
      const subject = `[Task Assigned - ${roleTitle}] ${taskId} | ${clientName}`;
      const body =
        `Hi ${newVal},\n\n` +
        `You have been assigned as ${roleTitle} on the following task:\n\n` +
        `• Task ID: ${taskId}\n` +
        `• Client Name: ${clientName}\n` +
        `• Company POC: ${poc}\n` +
        `• Service: ${service}\n` +
        `• Deliverable: ${deliverable}\n` +
        `• Assigned To: ${assignedTo || '-'}\n` +
        `• Assist: ${assist || '-'}\n` +
        `• Reviewer: ${reviewer || EMAIL_ALERT_CONFIG.DEFAULT_REVIEWER_NAME}\n` +
        `• Due Date: ${dueDate}\n` +
        (projectLink ? `• Project Folder: ${projectLink}\n` : '') +
        `\nPlease update Task Status to "In Progress" when you start.\n\n` +
        `Regards,\nKundan & Associates Automation`;

      safeSendAlertEmail_(targetEmail, subject, body);
    }
  }

  // 2. Task Status (Col S) changed to "Under Review" -> Notify Reviewer (defaults to Kundan Kumar)
  if (col === c.TASK_STATUS && newVal === 'Under Review') {
    const effectiveReviewer = reviewer || EMAIL_ALERT_CONFIG.DEFAULT_REVIEWER_NAME;
    const reviewerEmail =
      empMap[effectiveReviewer.toLowerCase()] ||
      empMap[EMAIL_ALERT_CONFIG.DEFAULT_REVIEWER_NAME.toLowerCase()] ||
      EMAIL_ALERT_CONFIG.FALLBACK_REVIEWER_EMAIL;

    const subject = `[Under Review] ${taskId} | ${clientName} — ${deliverable}`;
    const body =
      `Hi ${effectiveReviewer},\n\n` +
      `The following task has been marked "Under Review" and is awaiting review:\n\n` +
      `• Task ID: ${taskId}\n` +
      `• Client Name: ${clientName}\n` +
      `• Company POC: ${poc}\n` +
      `• Service: ${service}\n` +
      `• Deliverable: ${deliverable}\n` +
      `• Assigned To: ${assignedTo || '-'}\n` +
      `• Assist: ${assist || '-'}\n` +
      `• Due Date: ${dueDate}\n` +
      (projectLink ? `• Project Folder: ${projectLink}\n` : '') +
      `\nOnce reviewed, please tick Review Status (Col I) as TRUE.\n\n` +
      `Regards,\nKundan & Associates Automation`;

    safeSendAlertEmail_(reviewerEmail, subject, body);
  }

  // 3. Review Status (Col I) checkbox marked TRUE -> Notify Assigned To (and Assist if present)
  const isCheckedTrue = e.value === 'TRUE' || e.value === true;
  if (col === c.REVIEW_STATUS && isCheckedTrue) {
    const effectiveReviewer = reviewer || EMAIL_ALERT_CONFIG.DEFAULT_REVIEWER_NAME;
    const assigneeEmail = assignedTo ? empMap[assignedTo.toLowerCase()] : '';
    const assistEmail = assist ? empMap[assist.toLowerCase()] : '';

    if (assigneeEmail) {
      const subject = `[Review Approved - Ready for Delivery] ${taskId} | ${clientName}`;
      const body =
        `Hi ${assignedTo},\n\n` +
        `Good news — ${effectiveReviewer} has reviewed and marked Review Status as done for the task below.\n` +
        `You can now deliver it to the client and update Task Status to "Completed".\n\n` +
        `• Task ID: ${taskId}\n` +
        `• Client Name: ${clientName}\n` +
        `• Company POC: ${poc}\n` +
        `• Service: ${service}\n` +
        `• Deliverable: ${deliverable}\n` +
        `• Reviewed By: ${effectiveReviewer}\n` +
        `• Assist: ${assist || '-'}\n` +
        `• Due Date: ${dueDate}\n` +
        (projectLink ? `• Project Folder: ${projectLink}\n` : '') +
        `\nRegards,\nKundan & Associates Automation`;

      safeSendAlertEmail_(assigneeEmail, subject, body, assistEmail || '');
    }
  }
}

function getActiveEmployeeEmailMap_() {
  const map = {};
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const empSheet = ss.getSheetByName(EMAIL_ALERT_CONFIG.EMPLOYEE_SHEET_NAME);
  if (!empSheet) return map;

  const lastRow = empSheet.getLastRow();
  if (lastRow < 2) return map;

  const rows = empSheet.getRange(2, 1, lastRow - 1, 3).getValues();
  for (let i = 0; i < rows.length; i++) {
    const name = (rows[i][0] || '').toString().trim().toLowerCase();
    const email = (rows[i][1] || '').toString().trim();
    const status = (rows[i][2] || 'Active').toString().trim().toLowerCase();

    if (name && email && email.includes('@') && status === 'active') {
      map[name] = email;
    }
  }
  return map;
}

function safeSendAlertEmail_(recipient, subject, body, ccEmail) {
  try {
    const options = {
      to: recipient,
      subject: subject,
      body: body
    };
    if (ccEmail && ccEmail !== recipient) {
      options.cc = ccEmail;
    }
    MailApp.sendEmail(options);
  } catch (err) {
    Logger.log('Alert email error (' + recipient + '): ' + err.message);
  }
}

function formatAlertDate_(val) {
  if (!val) return 'Not set';
  if (Object.prototype.toString.call(val) === '[object Date]' && !isNaN(val)) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), 'dd-MM-yyyy');
  }
  return val.toString();
}
