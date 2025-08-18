// Debugging
var debug = true;
var sendAllowed = false;

// Tabs
var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
var tasksTabName = '📌 Tasks';
var emailTabName = 'ℹ️ Info';
var taskSheet = spreadsheet.getSheetByName(tasksTabName);
var emailSheet = spreadsheet.getSheetByName(emailTabName);

// Data Locations
var taskNameCol = 3 // Col holding name of task
var startCol = 4    // First col with an RA name
var emailRow = 2    // Row with first email address
var emailCol = 2    // Col holding email addresses

// Symbols
var upcomingSymbol = '⏰';
var missingSymbol = '❌';  

// Email Variables
var taskSubject = "[REMINDER] FlightTracker Tasks";
var taskCC = ['email'];

var introductionText = 'Please review the items below to stay on track with your RA responsibilities:';
var missingText = 'Missing Tasks:';
var upcomingText = 'Upcoming Tasks:';
var closingText = 'Be sure to check your Flight Tracker and complete anything outstanding.';
var signatureText = 'Thanks,<br>NAME';


// DO NOT EDIT BELOW THIS LINE
function sendTaskReminder() {
  let lastTaskCol = taskSheet.getLastColumn();
  for (let col = startCol; col <= lastTaskCol; col++) { 
    let emailAddress = emailSheet.getRange(col - startCol + emailRow, emailCol).getValue();
    let upcomingTasks = formatTaskList(col, upcomingSymbol);
    let missingTasks = formatTaskList(col, missingSymbol);

    if (upcomingTasks || missingTasks) {
      let name = emailSheet.getRange(col - startCol + emailRow, emailCol - 1).getValue();
      let body = formatBody(name, upcomingTasks, missingTasks);

      debug && console.log(emailAddress); 
      debug && console.log(body);
      sendAllowed && sendEmail(emailAddress, taskSubject, body, taskCC);
    }
  }
}


function formatBody(name, upcomingTasks, lateTasks) {
  return (
    `${name},<br><br>` +
    (introductionText && introductionText + '<br><br>') +
    (lateTasks && `<span style="font-weight:bold;">${missingText}</span><br>` +
      `<ul style="margin:0 0 16px 20px; padding:0;">${lateTasks}</ul>`) +
    (upcomingTasks && `<span style="font-weight:bold;">${upcomingText}</span><br>` +
      `<ul style="margin:0 0 16px 20px; padding:0;">${upcomingTasks}</ul>`) +
    (closingText && closingText + '<br><br>') +
    (signatureText && signatureText)
  );
}


function formatTaskList(col, symbol) {
  let tasks = getTasks(col, symbol);
  if (tasks.length > 0) {
    return '<ul>' +
      tasks.map(([date, task]) => {
        let formattedDate;
        try {
          let parsed = new Date(date);
          formattedDate = isNaN(parsed)
            ? date  // fallback to raw
            : parsed.toLocaleDateString('en-US', { weekday: 'long', month: 'long', day: 'numeric' });
        } catch {
          formattedDate = date;
        }
        return `<li>${task} – ${formattedDate}</li>`;
      }).join('') +
    '</ul>';
  }
  return '';
};


function getTasks(targetCol, targetSymbol) {
  let tasks = [];
  let lastRow = taskSheet.getLastRow();
  for (let row = 1; row <= lastRow; row++) {
    let cellValue = taskSheet.getRange(row, targetCol).getValue();
    if (cellValue === targetSymbol) {
      let taskCell = taskSheet.getRange(row, taskNameCol);  
      let richTextValue = taskCell.getRichTextValue(); 
      let taskName = richTextValue.getLinkUrl() 
        ? `<a href="${richTextValue.getLinkUrl()}">${richTextValue.getText()}</a>` 
        : richTextValue.getText();

      let date = taskSheet.getRange(row, taskNameCol - 1).getDisplayValue(); // column to the left
      tasks.push([date, taskName]);
    }
  }
  return tasks;
};


function sendEmail(recipient, subject, body, cc=null){
    var options = {
      htmlBody: body
    };
    if (cc !== null && cc.length > 0) {
        options.cc = cc.join(',');
    }
    GmailApp.sendEmail(recipient, subject, '', options);    
};
