# PowerShell---Send-PS-Variables-to-Google-Apps-Script
Use a PowerShell script to send local PowerShell values to a Google Sheet using Google Apps Script. This example uploads computer and battery information to your sheet.

1. Download the PS1 script.
2. Create a new Google Sheet. Click Tools > Script Editor. Remove all the code.gs code and paste the GAS code from the repo (or from the PS1 comments).
3. Run the setup() function, authenticate, and Run the setup() function again.
4. Click Publish > Deploy as a web app. Allow it to run as your account and by anyone, even anonymous. Click Publish.
5. Copy the 'Current web app URL' and paste that as the $GoogleWebAppURL value in the PS1 script.
6. Run the PS1 script. Your Google Sheet gets updated with the matching variable columns.

If you create new PowerShell variables you want passed in to the Sheet, a few things: 
* Either update the var headerRow in the GAS with the new PS variable(s) or manually update the sheet to have a new column.
* Create a new column name variable and a new column value variable. Then add them to the $postParams array.

Thanks to [/u/computir](https://www.reddit.com/u/computir) for the original script!
