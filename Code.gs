/**
 * After editing SuperMarkIt's scripts, File > Manage Versions, Create New
 * Back here: Resources > Libraries > Select new version

Works with the klausrheum/supermarkit github project.

*/


/**
 * A special function that runs when the spreadsheet is open, used to add a
 * custom menu to the spreadsheet.
 */
function onOpen() {
  installReportbookMenu();
}


function installReportbookMenu () {
  var masterUser = "classroom@hope.edu.kh";
  var spreadsheet = SpreadsheetApp.getActive();
  var adminMenuItems = [
    {name: '⚠ Import Courses', functionName: 'updateReportbookClassrooms'},
    {name: '⚠ Import Teachers', functionName: 'getTeachersFromTracker'},
    {name: '⚠ Generate Empty Reportbooks (Sync 🗹)', functionName: 'createMissingReportbooks'},
    {name: '⚠ Import Students', functionName: 'updateRbStudents'},
    null,
    {name: 'Import Grades', functionName: 'importGrades'},
    {name: 'Hide Admin Columns', functionName: 'hideCols'},
    {name: '⚠ Update Individual Reports tab', functionName: 'SuperMarkIt.updateReportbooks'},
    null,
    {name: '⚠ Generate Empty Portfolios', functionName: 'createPortfolios'},
    {name: "⚠ Update 🗹 Portfolios from 🗹 Courses", functionName: 'exportPortfolios'},    
    null,
    {name: '⚠ Push Extra-Curr, Pull Portfolio Admin', functionName: 'backupAllPastoralAdmin'},
    null,
    {name: '⚠ Generate PDFs for 🗹 Portfolios', functionName: 'generateSelectedPortfolioPDFs'},
    {name: '⚠ Generate PDFs for 🗹 Portfolios and email to guardians', functionName: 'generateAndSendSelectedPortfolioPDFs'},
    null,
    {name: '🕱 Delete ALL SUBJECTS from 🗹 Portfolios', functionName: 'keepKillPortfolioSheets'},
    {name: '🕱 Archive ALL Courses', functionName: 'archiveAllCourses'}
  ];
  
  var userMenuItems = [
    {name: 'Import Grades', functionName: 'importGrades'},
    null,
    {name: 'Hide Admin Columns', functionName: 'hideCols'}
  ];
  
  if (Session.getActiveUser().getEmail() == masterUser) {
    spreadsheet.addMenu('Reportbook', adminMenuItems);
  } else {
    spreadsheet.addMenu('Reportbook', userMenuItems);
  }
}

function hideCols() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Reportbooks');
  // index position 0  1  2   3   4
  var colsToHide = [4, 5, 6, 7, 8, 10, 13, 19, 20, 21, 22, 24, 28 ];

  for (var i = 0; i < colsToHide.length; i++) {
    var columnIndex = colsToHide[i];
    sh.hideColumns(columnIndex);    
  };
}


function archiveAllCourses() {
  SuperMarkIt.archiveAllCourses();
}


function updateReportbookClassrooms() {
  SuperMarkIt.updateReportbookClassrooms();
}


function getTeachersFromTracker() {
  SuperMarkIt.getTeachersFromTracker();
}


function createMissingReportbooks() {
  SuperMarkIt.createMissingReportbooks();
}


function generateSelectedPortfolioPDFs() {
  SuperMarkIt.generateSelectedPortfolioPDFs(false);
}


function importGrades() {
  
  // Y2025 ICT JKw
  var rbId; // = "1BijeGY49S0amD3u-eePjz8iWBwH1sEc7QE_yADzVzgQ";  
  var courseId; // = "16052292479";
  var sheet = SpreadsheetApp.getActiveSheet();
  if (sheet.getName() != "Reportbooks") {
    SpreadsheetApp.getUi().alert("ERROR: Select a course from the Reportbooks tab");
  } else {
    var selection = SpreadsheetApp.getSelection();
    //SpreadsheetApp.getUi().alert(currentRange);
    //Logger.log('Active Sheet: ' + selection.getActiveSheet().getName());
    // Current Cell: D1
    //Logger.log('Current Cell: ' + selection.getCurrentCell().getA1Notation());
    // Active Range: D1:E4
    //Logger.log('Active Range: ' + selection.getActiveRange().getA1Notation());
    // Active Ranges: A1:B4, D1:E4
    var ranges =  selection.getActiveRangeList().getRanges();
    Logger.log("selection contains %s cells", ranges.length);
    for (var i = 0; i < ranges.length; i++) {
      var row = ranges[i].getRow();
      var column = ranges[i].getColumn();
      Logger.log('row %s, column %s', row, column);
      
      var courseIdColumn = 5;
      var rbIdColumn = 1;
      var subjectColumn = 2;
      var timestampColumn = 17;
      
      if (column != 2) {
        SpreadsheetApp.getUi().alert("ERROR: Select a course from column B");
      } else {
        var rbId = sheet.getRange(row, rbIdColumn).getValue();
        var subjectName = sheet.getRange(row, subjectColumn).getValue();
        var courseId = sheet.getRange(row, courseIdColumn).getValue();
        if (rbId && courseId) {
          Logger.log("importGrades: rbId=%s, courseId=%s", rbId, courseId);
          SuperMarkIt.importGrades(rbId, courseId);
          
          var message = Utilities.formatString("IMPORT GRADES: " + subjectName);
          SuperMarkIt.logMe( message );

          sheet.getRange(row, timestampColumn).setValue(new Date()); 

        }
      }
    }
  }
}


function updateRbStudents() {
  var message = Utilities.formatString("Function %s executed", "updateRbStudents");
  SuperMarkIt.logToSheet( message );
  
  var sheet = SpreadsheetApp.getActiveSheet();
  if (sheet.getName() != "Reportbooks") {
    SpreadsheetApp.getUi().alert("ERROR: Select a course from the Reportbooks tab");
  } else {
    var selection = SpreadsheetApp.getSelection();
    var ranges =  selection.getActiveRangeList().getRanges();
    Logger.log("selection contains %s cells", ranges.length);
    for (var i = 0; i < ranges.length; i++) {
      var row = ranges[i].getRow();
      var column = ranges[i].getColumn();
      Logger.log('row %s, column %s', row, column);
      
      var courseIdColumn = 5;
      var rbIdColumn = 1;
      var timestampColumn = 17;
      
      if (column != 2) {
        SpreadsheetApp.getUi().alert("ERROR: Select a course from column B");
      } else {
        var rbId = sheet.getRange(row, rbIdColumn).getValue();
        var courseId = sheet.getRange(row, courseIdColumn).getValue();
        if (rbId && courseId) {
          Logger.log("updateRbStudents: rbId=%s, courseId=%s", rbId, courseId);
          var courseStudents = SuperMarkIt.listStudents(courseId);
          SuperMarkIt.updateRbStudents(rbId, courseStudents);
        }
      }
    }
  }
}


function keepKillPortfolioSheets() {
  var ui = SpreadsheetApp.getUi();
  
  if (Session.getActiveUser().getEmail() == masterUser) {
    var result = ui.alert(
      'NUCLEAR OPTION!!!',
      'Delete ALL generated subject tabs from all ticked Portfolios?',
      ui.ButtonSet.YES_NO);
    
    if (result == ui.Button.YES) {
      // User clicked "Yes".
      SuperMarkIt.keepKillPortfolioSheets(true);
    } else {
      // User clicked "No" or X in the title bar.
      
      ui.alert('Cancelled', 'Deletion cancelled.', ui.ButtonSet.OK);
    }

  } else {
    ui.alert("Sorry, only the master user can run this script");
  }
}
    
function generateAndSendSelectedPortfolioPDFs() {
  var ui = SpreadsheetApp.getUi();
  var result = ui.alert(
    'Email Guardians',
     'Portfolio PDFs will be generated for ALL ticked Portfolios and emailed out!',
     ui.ButtonSet.YES_NO);
  
  if (result == ui.Button.YES) {
    // User clicked "Yes".
    SuperMarkIt.generateSelectedPortfolioPDFs(true);
  } else {
    // User clicked "No" or X in the title bar.
    
    ui.alert('Cancelled', 'Emails cancelled.', ui.ButtonSet.OK);
  }

}
    
function createFilterView() {
  var name = "Thressa Brand";
  var ss = SpreadsheetApp.getActive();
  var sheet = SpreadsheetApp.getActiveSheet();
  Logger.log(ss.getColumnFilterCriteria(18));
  if (sheet.getName() == "Reportbooks") {
   
  }
}

function createPortfolios() {
  var ui = SpreadsheetApp.getUi();
  var result = ui.alert(
    'Create Spreadsheets for Student Portfolios',
     'Portfolio Documents will be generated for ALL Portfolio that do not have them (this will take a while)',
     ui.ButtonSet.YES_NO);
  
  if (result == ui.Button.YES) {
    // User clicked "Yes".
    SuperMarkIt.createStudents();  
  } else {
    // User clicked "No" or X in the title bar.
    
    ui.alert('Cancelled', 'Export cancelled.', ui.ButtonSet.OK);
  }
}

function exportPortfolios() {
  var message = Utilities.formatString("Function %s executed", "exportPortfolios")
  SuperMarkIt.logToSheet(message);
  
  var ui = SpreadsheetApp.getUi();
  var result = ui.alert(
    'WARNING: Export Reportbooks to Portfolios?',
     'Portfolio tabs will be (re)generated for ALL ticked Reportbook / Portfolio combinations?',
     ui.ButtonSet.YES_NO);
  
  if (result == ui.Button.YES) {
    // User clicked "Yes".
    SuperMarkIt.exportAllRBs();
  } else {
    // User clicked "No" or X in the title bar.
    
    ui.alert('Cancelled', 'Export cancelled.', ui.ButtonSet.OK);
  }
  
}


function backupAllPastoralAdmin() {
  var message = Utilities.formatString("Function %s executed", "backupAllPastoralAdmin")

  SuperMarkIt.backupAllPastoralAdmin();
}

function showAlert(title, prompt) {
  var ui = SpreadsheetApp.getUi(); // Same variations.
  
  var result = ui.alert(
     title,
     prompt,
      ui.ButtonSet.YES_NO);

  // Process the user's response.
  if (result == ui.Button.YES) {
    // User clicked "Yes".
    //ui.alert('Confirmation received.');
  } else {
    // User clicked "No" or X in the title bar.
    //ui.alert('Permission denied.');
  }
  
  return result == ui.Button.YES;
}