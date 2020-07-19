#!ps

<#
https://www.reddit.com/r/PowerShell/comments/8gk0l1/script_to_send_powershell_variables_to_google/
Create a new spreadsheet in Google Drive
Give the new sheet a name
Leave the sheet name as Sheet1 or ensure that you update the SHEET_NAME variable in the new code.gs script below.
Create column names ensuring that you use the same spacing and character case for your variables
If you create a column titled Timestamp, the script will automatically populate that value
Click Tools > Script Editor
Remove all the code.gs code and paste the code from the next comment block
Save the Project (Recommended: Use the name of your spreadsheet for the project code)
Select the function 'setup' and run it by clicking 'Play'
A dialog will pop up and you will have to grant access
Run the 'setup' function again
Click Publish > Deploy as a web app
Allow it to run as your account and by anyone, even anonymous. 
Click Publish
Copy the 'Current web app URL' and paste that as the $GoogleWebAppURL value, keeping the single quotes at the beginning and ending
Assign variables as needed below to match your column names as above
If you created a column called Timestamp, you don't have to create a variable for it in your Powershell script

#>

<#

// Code.gs replacement script
/*   
   Copyright 2011 Martin Hawksey

   Licensed under the Apache License, Version 2.0 (the "License");
   you may not use this file except in compliance with the License.
   You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
*/

// Usage
//  1. Enter sheet name where data is to be written below
        var SHEET_NAME = "Sheet1";

//  2. Run > setup
//
//  3. Publish > Deploy as web app 
//    - enter Project Version name and click 'Save New Version' 
//    - set security level and enable service (most likely execute as 'me' and access 'anyone, even anonymously) 
//
//  4. Copy the 'Current web app URL' and post this in your form/script action 
//
//  5. Insert column names on your destination sheet matching the parameter names of the data you are passing in (exactly matching case)

var SCRIPT_PROP = PropertiesService.getScriptProperties(); // new property service

// If you don't want to expose either GET or POST methods you can comment out the appropriate function
function doGet(e){
  return handleResponse(e);
}

function doPost(e){
  return handleResponse(e);
}

function handleResponse(e) {
  // shortly after my original solution Google announced the LockService[1]
  // this prevents concurrent access overwritting data
  // [1] http://googleappsdeveloper.blogspot.co.uk/2011/10/concurrency-and-google-apps-script.html
  // we want a public lock, one that locks for all invocations
  var lock = LockService.getPublicLock();
  lock.waitLock(30000);  // wait 30 seconds before conceding defeat.

  try {
    // next set where we write the data - you could write to multiple/alternate destinations
    var doc = SpreadsheetApp.openById(SCRIPT_PROP.getProperty("key"));
    var sheet = doc.getSheetByName(SHEET_NAME);

    // we'll assume header is in row 1 but you can override with header_row in GET/POST data
    var headRow = e.parameter.header_row || 1;
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var nextRow = sheet.getLastRow()+1; // get next row
    var row = []; 
    // loop through the header columns
    for (i in headers){
      if (headers[i] == "Timestamp"){ // special case if you include a 'Timestamp' column
        row.push(new Date());
      } else { // else use header name to get data
        row.push(e.parameter[headers[i]]);
      }
    }
    // more efficient to set values as [][] array than individually
    sheet.getRange(nextRow, 1, 1, row.length).setValues([row]);
    // return json success results
    return ContentService
          .createTextOutput(JSON.stringify({"result":"success", "row": nextRow}))
          .setMimeType(ContentService.MimeType.JSON);
  } catch(e){
    // if error return this
    return ContentService
          .createTextOutput(JSON.stringify({"result":"error", "error": e}))
          .setMimeType(ContentService.MimeType.JSON);
  } finally { //release lock
    lock.releaseLock();
  }
}

function setup() {
    var doc = SpreadsheetApp.getActiveSpreadsheet();
  
    //  Add header row
  var firstCell = doc.getActiveSheet().getRange(1, 1).getValue();
  if (firstCell != 'Timestamp') {
    var headerRow = ["Timestamp", "Computer", "IP Address", "OS Version", "Uptime", "Power Status", "Battery Charge Status", "Battery Full Lifetime", "Battery Life Percent", "Battery Life Remaining"];
    doc.appendRow(headerRow);
  }
  
    SCRIPT_PROP.setProperty("key", doc.getId());
}

#>

# Update Battery variables
# https://www.powershellmagazine.com/2012/10/18/pstip-get-system-power-information/

Add-Type -Assembly System.Windows.Forms;
$Battery = [System.Windows.Forms.SystemInformation]::PowerStatus;

# $GoogleWebAppURL = 'https://script.google.com/macros/s/xM7Ib7iXgXzx9XXXvuMXXzDkXXXtxUXXxxxWpxX1XXxWI1q3t-XXXXX/exec'
$GoogleWebAppURL = 'https://script.google.com/macros/s/AKfycbybizBXPbJyVVrXY83cIAff95UxMqZNgwcJw4kQcb32U9gXT0w/exec'

$Column1Name = 'Computer'
$Column1Value = $env:COMPUTERNAME

$Column2Name = 'IP Address'
$Column2Value = (Test-Connection -ComputerName ::1 -Count 1|select ipv4address).ipv4address

$Column3Name = 'OS Version'
$Column3Value = "$([System.Environment]::OSVersion.Version.Major)`.$([System.Environment]::OSVersion.Version.Minor)`.$([System.Environment]::OSVersion.Version.Build)`.$([System.Environment]::OSVersion.Version.Revision)"

$Column4Name = 'Uptime'
$Column4Value = (Get-CimInstance -ClassName win32_operatingsystem | select lastbootuptime).lastbootuptime

$Column5Name = 'Power Status'
$Column5Value = $Battery.PowerLineStatus
$Column6Name = 'Battery Charge Status'
$Column6Value = $Battery.BatteryChargeStatus
$Column7Name = 'Battery Full Lifetime'
$Column7Value = $Battery.BatteryFullLifetime
$Column8Name = 'Battery Life Percent'
$Column8Value = $Battery.BatteryLifePercent
$Column9Name = 'Battery Life Remaining'
$Column9Value = $Battery.BatteryLifeRemaining

$postParams = @{$Column1Name=$Column1Value;$Column2Name=$Column2Value;$Column3Name=$Column3Value;$Column4Name=$Column4Value;$Column5Name=$Column5Value;$Column6Name=$Column6Value;$Column7Name=$Column7Value;$Column8Name=$Column8Value;$Column9Name=$Column9Value}
Invoke-WebRequest -UseBasicParsing -Uri $GoogleWebAppURL -Method POST -Body $postParams

Write-Host `"$Column1Name=$Column1Value"&"$Column2Name=$Column2Value"&"$Column3Name=$Column3Value"&"$Column4Name=$Column4Value"&"$Column5Name=$Column5Value"&"$Column6Name=$Column6Value"&"$Column7Name=$Column7Value"&"$Column8Name=$Column8Value"&"$Column9Name=$Column9Value`"

# Other methods that may work (couldn't test since I'm on a desktop at the moment)
# https://devblogs.microsoft.com/scripting/using-windows-powershell-to-determine-if-a-laptop-is-on-battery-power/
# http://blog.technotesdesk.com/powershell-get-info-on-laptop-battery
# https://github.com/auberginehill/get-battery-info
