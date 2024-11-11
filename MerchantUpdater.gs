function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Custom Tools')
    .addItem('Refresh Data', 'copyData')
    .addItem('Set Notification Email', 'setEmailAddress')
    .addItem('Set Source and Target Sheets', 'setupSheetConfiguration')
    .addItem('Set Daily Timer', 'setupDailyTrigger')
    .addItem('Set Up TSV Configuration', 'setupConfigSheet')
    .addItem('Set Google Drive Folder ID', 'setFolderId')
    .addItem('Run TSV Export', 'exportTabsToTSV')
    .addToUi();

  // Automatically run copyData when the sheet is opened
  copyData();
}

function copyData() {
  const documentProperties = PropertiesService.getDocumentProperties();
  const sourceSpreadsheetId = documentProperties.getProperty('SOURCE_SPREADSHEET_ID');
  const sourceSheetName = documentProperties.getProperty('SOURCE_SHEET_NAME');
  const targetSpreadsheetId = documentProperties.getProperty('TARGET_SPREADSHEET_ID');
  const targetSheetName = documentProperties.getProperty('TARGET_SHEET_NAME');

  if (!sourceSpreadsheetId || !sourceSheetName || !targetSpreadsheetId || !targetSheetName) {
    Logger.log('Sheet configuration not set. Please use "Set Source and Target Sheets" in the menu.');
    return;
  }

  try {
    const sourceSpreadsheet = SpreadsheetApp.openById(sourceSpreadsheetId);
    const sourceSheet = sourceSpreadsheet.getSheetByName(sourceSheetName);

    const targetSpreadsheet = SpreadsheetApp.openById(targetSpreadsheetId);
    const targetSheet = targetSpreadsheet.getSheetByName(targetSheetName);

    const sourceRange = sourceSheet.getDataRange();
    const sourceValues = sourceRange.getValues();

    targetSheet.clearContents();
    targetSheet.getRange(1, 1, sourceValues.length, sourceValues[0].length).setValues(sourceValues);

    targetSheet.setFrozenRows(1);
    targetSheet.getRange(2, 1, sourceValues.length - 1, sourceValues[0].length).sort({ column: 3, ascending: true });

    const sourceSheetUrl = `https://docs.google.com/spreadsheets/d/${sourceSpreadsheetId}`;
    const targetSheetUrl = `https://docs.google.com/spreadsheets/d/${targetSpreadsheetId}`;
    notifySuccess(sourceValues.length, sourceValues[0].length, sourceSheetUrl, targetSheetUrl);

    // Add a delay to ensure all data processes are complete
    Utilities.sleep(5000); // Wait for 5 seconds before exporting TSV

    exportTabsToTSV(); // Run TSV export after data refresh
  } catch (error) {
    notifyFailure(error.message);
  }
}

function setEmailAddress() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('Set Notification Email', 'Enter your email address:', ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() === ui.Button.OK) {
    const email = response.getResponseText().trim();
    if (email) {
      PropertiesService.getDocumentProperties().setProperty('NOTIFICATION_EMAIL', email);
      ui.alert('Email address saved successfully.');
    } else {
      ui.alert('Invalid email address.');
    }
  }
}

function setupSheetConfiguration() {
  const ui = SpreadsheetApp.getUi();

  const sourceSpreadsheetResponse = ui.prompt(
    'Set Source Spreadsheet',
    'Enter the ID of the source spreadsheet:\n\nExample: For the URL\nhttps://docs.google.com/spreadsheets/d/1ABCDEF1234567890XYZ/edit#gid=0\n\nEnter: 1ABCDEF1234567890XYZ',
    ui.ButtonSet.OK_CANCEL
  );
  if (sourceSpreadsheetResponse.getSelectedButton() !== ui.Button.OK) return;
  const sourceSpreadsheetId = sourceSpreadsheetResponse.getResponseText().trim();

  const sourceSheetResponse = ui.prompt('Set Source Sheet', 'Enter the name of the source sheet:', ui.ButtonSet.OK_CANCEL);
  if (sourceSheetResponse.getSelectedButton() !== ui.Button.OK) return;
  const sourceSheetName = sourceSheetResponse.getResponseText().trim();

  const targetSpreadsheetResponse = ui.prompt(
    'Set Target Spreadsheet',
    'Enter the ID of the target spreadsheet:\n\nExample: For the URL\nhttps://docs.google.com/spreadsheets/d/1GHIJKL9876543210UVW/edit#gid=0\n\nEnter: 1GHIJKL9876543210UVW',
    ui.ButtonSet.OK_CANCEL
  );
  if (targetSpreadsheetResponse.getSelectedButton() !== ui.Button.OK) return;
  const targetSpreadsheetId = targetSpreadsheetResponse.getResponseText().trim();

  const targetSheetResponse = ui.prompt('Set Target Sheet', 'Enter the name of the target sheet:', ui.ButtonSet.OK_CANCEL);
  if (targetSheetResponse.getSelectedButton() !== ui.Button.OK) return;
  const targetSheetName = targetSheetResponse.getResponseText().trim();

  const documentProperties = PropertiesService.getDocumentProperties();
  documentProperties.setProperty('SOURCE_SPREADSHEET_ID', sourceSpreadsheetId);
  documentProperties.setProperty('SOURCE_SHEET_NAME', sourceSheetName);
  documentProperties.setProperty('TARGET_SPREADSHEET_ID', targetSpreadsheetId);
  documentProperties.setProperty('TARGET_SHEET_NAME', targetSheetName);

  ui.alert('Sheet configuration saved successfully.');
}

function setupDailyTrigger() {
  const ui = SpreadsheetApp.getUi();
  const timeResponse = ui.prompt(
    'Set Daily Trigger Time',
    'Enter the time in 24-hour format (e.g., 08 for 8 AM, 14 for 2 PM):',
    ui.ButtonSet.OK_CANCEL
  );

  if (timeResponse.getSelectedButton() !== ui.Button.OK) return;
  const time = parseInt(timeResponse.getResponseText().trim(), 10);

  if (isNaN(time) || time < 0 || time > 23) {
    ui.alert('Invalid time. Please enter a valid hour between 0 and 23.');
    return;
  }

  createDailyTrigger(time);
  ui.alert(`Daily trigger set for ${time}:00.`);
}

function createDailyTrigger(hour) {
  deleteExistingTriggers();
  ScriptApp.newTrigger('copyData')
    .timeBased()
    .atHour(hour)
    .everyDays(1)
    .create();
  Logger.log(`Daily trigger set for ${hour}:00.`);
}

function deleteExistingTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'copyData') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

function setupConfigSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TSV_Config');
  if (!sheet) {
    const configSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet('TSV_Config');
    configSheet.getRange('A1:C1').setValues([['Tab Name', 'TSV File Name', 'TSV File URL']]);
    configSheet.setColumnWidths(1, 3, 300);
    SpreadsheetApp.getUi().alert('TSV_Config sheet created! Enter your configuration details.');
  } else {
    SpreadsheetApp.getUi().alert('TSV_Config sheet already exists!');
  }
}

function setFolderId() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    'Set Google Drive Folder ID',
    'Enter the ID of the folder where TSV files will be saved:\n\nExample: For the folder URL\nhttps://drive.google.com/drive/folders/1ABCDEF1234567890XYZ\n\nEnter: 1ABCDEF1234567890XYZ',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() === ui.Button.OK) {
    const folderId = response.getResponseText().trim();
    if (folderId) {
      PropertiesService.getDocumentProperties().setProperty('GOOGLE_DRIVE_FOLDER_ID', folderId);
      ui.alert('Folder ID saved successfully.');
    } else {
      ui.alert('Invalid Folder ID. Please try again.');
    }
  }
}

function exportTabsToTSV() {
  const folderId = PropertiesService.getDocumentProperties().getProperty('GOOGLE_DRIVE_FOLDER_ID');
  if (!folderId) {
    Logger.log('Google Drive Folder ID not set. Please use "Set Google Drive Folder ID" in the menu.');
    return;
  }

  const configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TSV_Config');
  if (!configSheet) {
    Logger.log('Please set up the TSV_Config sheet first.');
    return;
  }

  const configData = configSheet.getRange(2, 1, configSheet.getLastRow() - 1, 2).getValues();
  const folder = DriveApp.getFolderById(folderId);

  configData.forEach(([tabName, fileName], index) => {
    if (!tabName || !fileName) return;

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(tabName);
    if (!sheet) {
      Logger.log(`Tab "${tabName}" not found.`);
      return;
    }

    const data = sheet.getDataRange().getValues();
    const tsvContent = data.map(row => row.join('\t')).join('\n');
    const blob = Utilities.newBlob(tsvContent, 'text/tab-separated-values', fileName);

    let existingFile = folder.getFilesByName(fileName);
    if (existingFile.hasNext()) {
      let file = existingFile.next();
      file.setTrashed(true); // Trash the old file to maintain consistent URL
    }

    const newFile = folder.createFile(blob);
    const fileUrl = newFile.getUrl();

    // Update the TSV File URL in TSV_Config
    configSheet.getRange(index + 2, 3).setValue(fileUrl);
  });
}

function notifySuccess(rowCount, columnCount, sourceSheetUrl, targetSheetUrl) {
  const email = PropertiesService.getDocumentProperties().getProperty('NOTIFICATION_EMAIL');
  if (email) {
    const message = `The data refresh was completed successfully.
    
Details:
- Total rows updated: ${rowCount}
- Total columns updated: ${columnCount}
- Source Sheet: ${sourceSheetUrl}
- Target Sheet: ${targetSheetUrl}`;

    MailApp.sendEmail(email, 'Data Refresh Successful', message);
  } else {
    Logger.log('No notification email set.');
  }
}

function notifyFailure(errorMessage) {
  const email = PropertiesService.getDocumentProperties().getProperty('NOTIFICATION_EMAIL');
  if (email) {
    const message = `The data refresh failed with the following error:
    
${errorMessage}`;

    MailApp.sendEmail(email, 'Data Refresh Failed', message);
  }
  Logger.log(errorMessage);
}
