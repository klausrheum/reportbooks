/**
 * After editing SuperMarkIt's scripts, File > Manage Versions, Create New
 * Back here: Resources > Libraries > Select new version

Works with the klausrheum/supermarkit github project.

*/


/**
 * A special function that runs when the spreadsheet is open, used to add a
 * custom menu to the spreadsheet.
 */
SuperMarkIt.top.FILES.RBTRACKER = SpreadsheetApp.getActiveSpreadsheet().getId();
Logger.log(SuperMarkIt.top.FILES.RBTRACKER);
SuperMarkIt.top.META.SEM = SpreadsheetApp.getActiveSpreadsheet().getName().slice(-7); // ie 'Dec2020' from 'Reportbooks Dec2020'
Logger.log(SuperMarkIt.top.META.SEM);

var masterUser = "classroom@hope.edu.kh";

function onOpen() {
  installReportbookMenu();
}

function testTop() {
  SuperMarkIt.TESTtop();
}

function installReportbookMenu () {
  SuperMarkIt.top.FILES.RBTRACKER = SpreadsheetApp.getActiveSpreadsheet().getId();
  
  var spreadsheet = SpreadsheetApp.getActive();
  var adminMenuItems = [
    {name: '⚠ Import Courses', functionName: 'updateReportbookClassrooms'},
    {name: '⚠ Import Teachers', functionName: 'getTeachersFromTracker'},
    {name: '⚠ Generate Reportbooks (if Sync 🗹)', functionName: 'createMissingReportbooks'},
    {name: '⚠ Import Students (if Sync 🗹)', functionName: 'updateRbStudents'},
    null,
    {name: 'Import Grades', functionName: 'importGrades'},
    {name: 'Hide Admin Columns', functionName: 'hideCols'},
    {name: '⚠ Update Individual Reports tab', functionName: 'updateReportbooks'},
    null,
    {name: '⚠ Generate Empty Portfolios', functionName: 'createPortfolios'},
    {name: "⚠ Update 🗹 Portfolios from 🗹 Courses", functionName: 'exportPortfolios'},    
    null,
    {name: '⚠ Push Extra-Curr, Pull Portfolio Admin', functionName: 'backupAllPastoralAdmin'},
    null,
    {name: '⚠ Generate PDFs for 🗹 Portfolios', functionName: 'generateSelectedPortfolioPDFs'},
    null,
    {name: '⚠⚠⚠ Generate PDFs for 🗹 Portfolios and email to guardians', functionName: 'generateAndSendSelectedPortfolioPDFs'},
    null,
    {name: '🕱 Delete ALL SHEETS from 🗹 Portfolios', functionName: 'keepKillPortfolioSheets'},
    {name: '🔍 Find SHEETS from 🗹 Portfolios matching Admin KEEP/KILL', functionName: 'keepKillPortfolioSheetsMatchingRegexFind'},
    {name: '🗑️ Delete SHEETS from 🗹 Portfolios matching Admin KEEP/KILL', functionName: 'keepKillPortfolioSheetsMatchingRegexDelete'},
    {name: '📁 Archive ALL Courses', functionName: 'archiveAllCourses'},
    null,
    {name: 'testTop', functionName: 'testTop'}   
  ];
  
  var selectMenuItems = [
    {name: '🗹 Reportbooks Y2020 PP', functionName: 'tickBoxes_Reportbooks_Y2020_PP'},
    null,
    {name: '□ Reportbooks NONE', functionName: 'tickReportbooksNone'},
    {name: '🗹 Reportbooks SR', functionName: 'tickReportbooksSR'},
    {name: '🗹 Reportbooks Y2020', functionName: 'tickReportbooksPP2020'},
    {name: '🗹 Reportbooks Y2021', functionName: 'tickReportbooksPP2021'},
    {name: '🗹 Reportbooks Y2022', functionName: 'tickReportbooksPP2022'},
    {name: '🗹 Reportbooks Y2023', functionName: 'tickReportbooksPP2023'},
    {name: '🗹 Reportbooks Y2024', functionName: 'tickReportbooksPP2024'},
    {name: '🗹 Reportbooks Y2025', functionName: 'tickReportbooksPP2025'},
    {name: '🗹 Reportbooks Y2026', functionName: 'tickReportbooksPP2026'},
    null,
    {name: '□ Portfolios NONE', functionName: 'tickPortfoliosNone'},
    {name: '🗹 Portfolios SR',   functionName: 'tickPortfoliosSR'},
    {name: '🗹 Portfolios PP06', functionName: 'tickPortfoliosPP06'},
    {name: '🗹 Portfolios PP07', functionName: 'tickPortfoliosPP07'},
    {name: '🗹 Portfolios PP08', functionName: 'tickPortfoliosPP08'},
    {name: '🗹 Portfolios PP09', functionName: 'tickPortfoliosPP09'},
    {name: '🗹 Portfolios PP10', functionName: 'tickPortfoliosPP10'},
    {name: '🗹 Portfolios PP11', functionName: 'tickPortfoliosPP11'},
    {name: '🗹 Portfolios PP12', functionName: 'tickPortfoliosPP12'}
  ];
  
  var userMenuItems = [
    {name: 'Import Grades', functionName: 'importGrades'},
    null,
    {name: 'Hide Admin Columns', functionName: 'hideCols'}
  ];
  
  if (Session.getActiveUser().getEmail() == masterUser) {
    spreadsheet.addMenu('Admin', adminMenuItems);
    spreadsheet.addMenu('🗹', selectMenuItems);
  } else {
    spreadsheet.addMenu('Reportbook', userMenuItems);
  }
}

function tickReportbooksNone() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Reportbooks");
  var fields = {'Y0000': ['']};
  tickBoxes(sheet, fields, 'Export'); 
}

function tickBoxes_Reportbooks_TEST() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Reportbooks");
  var fields = {'Y0000': ['Y2020', 'Y2022'], 'PP SR': ['PP']};
  tickBoxes(sheet, fields, 'Export'); 
}

function tickReportbooksPP2020() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Reportbooks");
  var fields = {'Y0000': ['Y2020'], 'PP SR': ['PP']};
  tickBoxes(sheet, fields, 'Export'); 
}

function tickReportbooksPP2021() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Reportbooks");
  var fields = {'Y0000': ['Y2021'], 'PP SR': ['PP']};
  tickBoxes(sheet, fields, 'Export'); 
}

function tickReportbooksPP2022() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Reportbooks");
  var fields = {'Y0000': ['Y2022'], 'PP SR': ['PP']};
  tickBoxes(sheet, fields, 'Export'); 
}

function tickReportbooksPP2023() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Reportbooks");
  var fields = {'Y0000': ['Y2023'], 'PP SR': ['PP']};
  tickBoxes(sheet, fields, 'Export'); 
}

function tickReportbooksPP2024() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Reportbooks");
  var fields = {'Y0000': ['Y2024'], 'PP SR': ['PP']};
  tickBoxes(sheet, fields, 'Export');
}

function tickReportbooksPP2025() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Reportbooks");
  var fields = {'Y0000': ['Y2025'], 'PP SR': ['PP']};
  tickBoxes(sheet, fields, 'Export'); 
}

function tickReportbooksPP2026() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Reportbooks");
  var fields = {'Y0000': ['Y2026'], 'PP SR': ['PP']};
  tickBoxes(sheet, fields, 'Export'); 
}

function tickReportbooksSR() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Reportbooks");
  var fields = {'Y0000': [], 'PP SR': ['SR']};
  tickBoxes(sheet, fields, 'Export'); 
}


function tickPortfoliosNone() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Portfolios");
  var fields = {'AA00': ['X']};
  tickBoxes(sheet, fields, 'Export'); 
}

function tickPortfoliosPP06() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Portfolios");
  var fields = {'AA00': ['PP06']};
  tickBoxes(sheet, fields, 'Export'); 
}

function tickPortfoliosPP07() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Portfolios");
  var fields = {'AA00': ['PP07']};
  tickBoxes(sheet, fields, 'Export'); 
}

function tickPortfoliosPP08() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Portfolios");
  var fields = {'AA00': ['PP08']};
  tickBoxes(sheet, fields, 'Export'); 
}

function tickPortfoliosPP09() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Portfolios");
  var fields = {'AA00': ['PP09']};
  tickBoxes(sheet, fields, 'Export'); 
}

function tickPortfoliosPP10() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Portfolios");
  var fields = {'AA00': ['PP10']};
  tickBoxes(sheet, fields, 'Export'); 
}

function tickPortfoliosPP11() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Portfolios");
  var fields = {'AA00': ['PP11']};
  tickBoxes(sheet, fields, 'Export'); 
}

function tickPortfoliosPP12() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Portfolios");
  var fields = {'AA00': ['PP12']};
  tickBoxes(sheet, fields, 'Export'); 
}

function tickPortfoliosSR() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Portfolios");
  var fields = {'AA00': ['SR06', 'SR07', 'SR08', 'SR09', 'SR10']};
  tickBoxes(sheet, fields, 'Export'); 
}


function tickBoxes(sheet, fields, tickField) {
  // check updateColumn exists
  var tickCol = null;
  var headers = sheet.getRange("A1:1").getValues()[0];
  for (var i = 0; i < headers.length; i++) {
    var header = headers[i];
    Logger.log(header);
    if (headers[i] === tickField) {
      tickCol = i+1;
    }
  }
  if (!tickCol) {
    throw "Could not find the column header '" + tickField + "'";
  }
      
  var opt = {} 
  var data = SuperMarkIt.getRows(sheet, opt);
  
  var yesNos = [];
  for (var i = 0; i < data.length; i++) {
    var thisRecord = data[i];
    var match = true;
    var matchStr = "";
    for (var key in fields) {
      var val = thisRecord[key];
      var reqKeys = fields[key];
      
      if (reqKeys.length > 0) {
        var found = (reqKeys.indexOf(val) > -1);
        match = match && found;
      }
      matchStr += '\nIs ' + val + " in [" + reqKeys + ']? ' + found;      
    }
    Logger.log(matchStr);
    Logger.log('Final decision: ' + match + "\n");
    if (match) {
      Logger.log('Setting cell ' + i + ' to Y'); 
      yesNos.push(['Y']);
    } else {
      Logger.log('Setting cell ' + i + ' to N');
      yesNos.push(['N']);
    }
  }
  //Logger.log('setValues(\n' + yesNos);
  var numRows = sheet.getMaxRows()-1;
  var tickRange = sheet.getRange(2, tickCol, numRows);
  tickRange.setValues(yesNos);
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

function updateReportbooks() {
  SuperMarkIt.updateReportbooks();
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
    SpreadsheetApp.getUi().alert("ERROR: Check 'Sync' column on the Reportbooks tab");
  } else {
//    var selection = SpreadsheetApp.getSelection();
//    var ranges =  selection.getActiveRangeList().getRanges();
//    Logger.log("selection contains %s cells", ranges.length);
    var rbIdColumn = 1;
    var courseNameColumn = 2;
    var courseIdColumn = 5;
    var timestampColumn = 17;
    var syncColumn = 21;
    var startCol = 1;
    var startRow = 2;
    var numRows = sheet.getMaxRows()-1;
    var numCols = sheet.getMaxColumns();

    var range = sheet.getRange(startRow, startCol, numRows, numCols);
    var values = range.getValues();
        
    for (var i = 0; i < values.length; i++) {
      // sheet.setActiveRange(sheet.getRange(i+2, 2));
      // Logger.log(values[i]);
      var rbId = values[i][rbIdColumn-1];
      var courseId = values[i][courseIdColumn-1];
      var sync = values[i][syncColumn-1];
      var courseName = values[i][courseNameColumn-1];
      
      // Logger.log("Import Students: %s sync=%s", courseName, sync);
      if (rbId && courseId && sync) {
        var message = "updateRbStudents: " + courseName + " " + sync;
        Logger.log(message);
        SuperMarkIt.logMe(message);
  
        var courseStudents = SuperMarkIt.listStudents(courseId);
        SuperMarkIt.updateRbStudents(rbId, courseStudents);
      }
    }
  }
}


function keepKillPortfolioSheets() {
  var ui = SpreadsheetApp.getUi();
  var sheetsRemoved = [];
  
  if (Session.getActiveUser().getEmail() == masterUser) {
    var result = ui.alert(
      'NUCLEAR OPTION!!!',
      'Delete ALL generated subject tabs from all ticked Portfolios?',
      ui.ButtonSet.YES_NO);
    
    if (result == ui.Button.YES) {
      // User clicked "Yes".
      sheetsRemoved = SuperMarkIt.keepKillPortfolioSheets(true);
    } else {
      // User clicked "No" or X in the title bar.
      
      ui.alert('Cancelled', 'Deletion cancelled.', ui.ButtonSet.OK);
    }

  } else {
    ui.alert("Sorry, only the master user can run this script");
  }
  
  return sheetsRemoved;
}
    

function keepKillPortfolioSheetsMatchingRegexFind() {
  var forReal = false;
  keepKillPortfolioSheetsMatchingRegex(forReal);
}

function keepKillPortfolioSheetsMatchingRegexDelete() {
  var forReal = true;
  keepKillPortfolioSheetsMatchingRegex(forReal);  
}

function keepKillPortfolioSheetsMatchingRegex(forReal) {
    if (forReal === undefined) {
    forReal = false;
  }
  var sheetsRemoved = [];
  var ui = SpreadsheetApp.getUi();
  
  if (Session.getActiveUser().getEmail() == masterUser) {
    
    var keepPatterns = SpreadsheetApp.getActiveSpreadsheet().getRangeByName('keepRegex').getValues();
    var killPatterns = SpreadsheetApp.getActiveSpreadsheet().getRangeByName('killRegex').getValues();
    
    if (keepPatterns && killPatterns) {

      var keepRegexes = [];
      var killRegexes = [];
      for (var i=0; i<keepPatterns.length; i++) {
        var keepPatternString = keepPatterns[i].toString();
        if (keepPatternString.length > 0) {
          keepRegexes.push(new RegExp(keepPatterns[i]));
        }
      }
      for (var i=0; i<killPatterns.length; i++) {
        var killPatternString = killPatterns[i].toString();
        if (killPatternString.length > 0) {
          killRegexes.push(new RegExp(killPatterns[i]));
        }
      }
//      SuperMarkIt.logMe("keepRegexes: " + keepRegexes);
//      SuperMarkIt.logMe("killRegexes: " + killRegexes);
      
      var displayText = 'all sheets containing \n' + killPatterns + ' \n from all ticked Portfolios, but KEEP any containing \n' + keepPatterns + '\n  (see Admin tab)';
     
      if (forReal) {
        displayText = 'DELETE ' + displayText;
      } else {
        displayText = 'DISPLAY ' + displayText;
      }
      
      var result = ui.alert(
        'CONFIRM?!',
        displayText,
        ui.ButtonSet.YES_NO);
      
      if (result == ui.Button.YES) {
        // User clicked "Yes".
        sheetsRemoved = SuperMarkIt.keepKillPortfolioSheets(keepRegexes, killRegexes, forReal);
      } else {
        // User clicked "No" or X in the title bar.
        
        ui.alert('Cancelled', 'Deletion cancelled.', ui.ButtonSet.OK);
      }
    }
  } else {
    ui.alert("Sorry, only the master user can run this script");
  }
  
  return sheetsRemoved;
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