Google Sheets Automation Script
This Google Apps Script automates data refreshing between Google Sheets, exports tabs to TSV files in Google Drive, and sends email notifications upon success or failure. It is designed to run headlessly, compatible with triggers, and maintains all functionalities accessible via a custom menu in Google Sheets.

Table of Contents
Features
Prerequisites
Installation
Setup Instructions
1. Set Notification Email
2. Set Source and Target Sheets
3. Set Up TSV Configuration
4. Set Google Drive Folder ID
5. Set Daily Timer (Optional)
Usage
Manual Refresh
Automated Refresh
Error Handling
Logging
Contributing
License
Features
Data Refresh: Copies data from a source sheet to a target sheet, with sorting and formatting.
TSV Export: Exports specified tabs to TSV files in a designated Google Drive folder.
Email Notifications: Sends email notifications upon successful completion or failure of operations.
Custom Menu: Provides a user-friendly menu for configuration and manual operations.
Automated Triggers: Supports automated daily data refreshes via triggers.
Prerequisites
A Google Account with access to Google Sheets and Google Drive.
Basic understanding of Google Apps Script.
Necessary permissions to access and modify the source and target spreadsheets.
Installation
Open Your Google Spreadsheet:

Navigate to the Google Sheet where you want to add this script.
Access Apps Script Editor:

Click on Extensions > Apps Script to open the script editor.
Add the Script:

Replace any existing code with the script provided below.
javascript
Copy code
// [Insert the entire script here]
Save the Script:

Click on the floppy disk icon or press Ctrl+S to save the script.
Authorize the Script:

Upon running the script for the first time, you will be prompted to authorize it.
Follow the on-screen instructions to grant the necessary permissions.
Setup Instructions
After installing the script, you need to configure it using the custom menu added to your Google Sheet.

1. Set Notification Email
Configure the email address to receive notifications.

Go to Custom Tools > Set Notification Email.
Enter your email address when prompted.
Click OK to save.
2. Set Source and Target Sheets
Define the source and target spreadsheets and sheets for data copying.

Go to Custom Tools > Set Source and Target Sheets.
Enter the ID of the source spreadsheet when prompted.
Example: For URL https://docs.google.com/spreadsheets/d/1ABCDEF1234567890XYZ/edit#gid=0, the ID is 1ABCDEF1234567890XYZ.
Enter the name of the source sheet.
Enter the ID of the target spreadsheet.
Enter the name of the target sheet.
Confirm that the configuration was saved successfully.
3. Set Up TSV Configuration
Create and configure the TSV_Config sheet for exporting tabs to TSV files.

Go to Custom Tools > Set Up TSV Configuration.
If the TSV_Config sheet doesn't exist, it will be created.
In the TSV_Config sheet, enter the following details starting from row 2:
Tab Name: Name of the sheet/tab to export.
TSV File Name: Desired name for the TSV file (including the .tsv extension).
TSV File URL: This will be populated automatically after export.
4. Set Google Drive Folder ID
Specify the Google Drive folder where TSV files will be saved.

Go to Custom Tools > Set Google Drive Folder ID.
Enter the ID of the Google Drive folder.
Example: For URL https://drive.google.com/drive/folders/1ABCDEF1234567890XYZ, the ID is 1ABCDEF1234567890XYZ.
Confirm that the folder ID was saved successfully.
5. Set Daily Timer (Optional)
Set up an automated daily trigger to refresh data at a specific time.

Go to Custom Tools > Set Daily Timer.
Enter the hour of the day (0-23) in 24-hour format when prompted.
Example: Enter 8 for 8 AM or 14 for 2 PM.
Confirm that the daily trigger was set successfully.
Usage
Manual Refresh
To manually refresh data and export TSV files:

Go to Custom Tools > Refresh Data.
The script will:
Copy data from the source sheet to the target sheet.
Sort and format the target sheet.
Export specified tabs to TSV files in the designated Google Drive folder.
Upon completion, you will receive an email notification if configured.
Automated Refresh
If you have set up the daily timer, the script will automatically run the copyData function at the specified time each day.

Error Handling
The script includes comprehensive error handling and notifications.

Missing Configuration:
If required configurations (e.g., folder ID, TSV configuration) are missing, the script logs the error and sends a failure notification email.
Exceptions:
Any exceptions during execution are caught, and a detailed error message is sent via email.
Logging
The script uses Logger.log() to record significant events and errors.
To view logs:
Open the script editor.
Go to View > Logs.
Contributing
Contributions are welcome! If you have suggestions or improvements, feel free to create a pull request or open an issue on GitHub.

License
This project is licensed under the MIT License.

Note: Ensure that all IDs and sensitive information are kept secure and not shared publicly. Always test the script in a controlled environment before deploying it for production use.
