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
