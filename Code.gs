/**
 * @OnlyCurrentDoc
 *
 * The above comment directs Apps Script to limit the scope of file
 * access for this add-on. It specifies that this add-on will only
 * attempt to read or modify the files in which the add-on is used,
 * and not all of the user's files. The authorization request message
 * presented to users will reflect this limited scope.
 */
 
/**
 * Creates a menu entry in the Google Docs UI when the document is opened.
 * This method is only used by the regular add-on, and is never called by
 * the mobile add-on version.
 *
 * @param {object} e The event parameter for a simple onOpen trigger. To
 *     determine which authorization mode (ScriptApp.AuthMode) the trigger is
 *     running in, inspect e.authMode.
 */
function onOpen(e) {
  SpreadsheetApp.getUi().createAddonMenu()
      .addItem('Start', 'showSidebar')
      .addToUi();
}

/**
 * Runs when the add-on is installed.
 * This method is only used by the regular add-on, and is never called by
 * the mobile add-on version.
 *
 * @param {object} e The event parameter for a simple onInstall trigger. To
 *     determine which authorization mode (ScriptApp.AuthMode) the trigger is
 *     running in, inspect e.authMode. (In practice, onInstall triggers always
 *     run in AuthMode.FULL, but onOpen triggers may be AuthMode.LIMITED or
 *     AuthMode.NONE.)
 */
function onInstall(e) {
  onOpen(e);
}

/**
 * Opens a sidebar in the document containing the add-on's user interface.
 * This method is only used by the regular add-on, and is never called by
 * the mobile add-on version.
 */
function showSidebar() {
  var ui = HtmlService.createHtmlOutputFromFile('Sidebar')
      .setTitle('gsbook');
  SpreadsheetApp.getUi().showSidebar(ui);
}

/**
 * Conditional formatting on Grade sheet
 */
function onEdit(e) {
  if (e) { 
    var ss = e.source.getActiveSheet();
    var r = e.source.getActiveRange(); 
    if ((r.getRow() > 3) && (ss.getName() == "Grade")) {
      var value = r.getValue()
      if (value == 1) {
        r.setBackgroundColor("red");
      } else if (value == 2) {
        r.setBackgroundColor("orange");
      } else if (value == 3) { 
        r.setBackgroundColor("yellow")
      } else if (value == 4) {
        r.setBackgroundColor("lime")
      } else {
        r.setBackgroundColor("white")
      }
    }
  }
}

/**
 * Creates Grade sheet from existing Roster and Assessment sheets
 */
function create() {
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var gradeSheet = activeSpreadsheet.getSheetByName("Grade");
  var desiredMaxNumStudents = 100;

  if (gradeSheet == null) {
    gradeSheet = activeSpreadsheet.insertSheet("Grade");
    gradeSheet.deleteRows(desiredMaxNumStudents+20, gradeSheet.getMaxRows()-desiredMaxNumStudents-30) // delete extra rows

    var assessmentValues = activeSpreadsheet.getSheetByName("Assessment")
        .getDataRange()
        .getValues();    
    
    // Set assessment data
    offset_col = 1
    for (var i=0; i<assessmentValues.length; i++) {
      var nameStandard = assessmentValues[i][0]
      var numVersions = assessmentValues[i][1] + 1  // add one as version M
      var typeStandard = assessmentValues[i][2]
      
      if (gradeSheet.getMaxColumns() < offset_col+numVersions) {
        gradeSheet.insertColumnsAfter(gradeSheet.getMaxColumns(), numVersions)
      }
      
      gradeSheet.getRange(1, offset_col+1, 1, numVersions)
          .merge()
          .setValue(nameStandard)
          .setFontWeight('bold')
          .setHorizontalAlignment('center');
      gradeSheet.getRange(2, offset_col+1, 1, numVersions)
          .merge()
          .setValue(typeStandard)
          .setFontWeight('bold')
          .setHorizontalAlignment('center');
      for (var j=1; j<numVersions; j++) {
        gradeSheet.getRange(3, offset_col+1+j)
            .setValue(j)
            .setFontWeight('bold')
            .setHorizontalAlignment('center');
      }
      gradeSheet.getRange(3, offset_col+1)
          .setValue('M')
          .setFontWeight('bold')
          .setHorizontalAlignment('center');
      
      // Set formula for version M
      gradeSheet.getRange(4, offset_col+1, desiredMaxNumStudents, 1)
          .setFormulaR1C1("=MAX(R[0]C[1]:R[0]C[" + (numVersions-1) + "])")
      
      offset_col += numVersions
    }    
    
    for (var i=1; i<gradeSheet.getMaxColumns(); i++) {
      gradeSheet.setColumnWidth(i+1,30)
    }
    
    gradeSheet.setColumnWidth(1,100);
    gradeSheet.setFrozenColumns(1);
    gradeSheet.getRange(3,1).setValue('Name')
  } else {
    throw "Grade sheet already existed.";
  }
  return "Grade sheet successfully created.";
}


/**
 * Creates Grade report from Grade sheet
 */
function makeReport() {
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var reportSheet = activeSpreadsheet.getSheetByName("Report");
  var gradeSheet = activeSpreadsheet.getSheetByName("Grade");
  var assessmentSheet = activeSpreadsheet.getSheetByName("Assessment");
  
  if (gradeSheet != null) {
    if (reportSheet == null) {
      reportSheet = activeSpreadsheet.insertSheet("Report");
    }
    reportSheet.clear();
    if (reportSheet.getMaxColumns() < 20) {
      reportSheet.insertColumnsBefore(1, 10)
    }
    
    var assessmentValues = activeSpreadsheet.getSheetByName("Assessment")
        .getDataRange()
        .getValues();   
    
    // Prefetch grade values: [[name, score, score...],...]
    var gradeValues = gradeSheet.getRange(4, 1, gradeSheet.getMaxRows()-3, gradeSheet.getMaxColumns()-1)
        .getValues()
    
    //Populate report: [[name, maxOfStandard, maxOfStandard...]...]
    var report = []
    for (var k=0; k<gradeValues.length; k++) {
      if (gradeValues[k][0] != null && gradeValues[k][0] != "") {
        var studentReport = [gradeValues[k][0]]
        for (var i=0; i<assessmentValues.length; i++) {
          var numVersions = assessmentValues[i][1]+1 //one extra for Version M
          var maxOfStandard = gradeValues[k][i*numVersions+1]
          studentReport.push([maxOfStandard])
        }  
        report.push(studentReport)
      }
    }
    
    //Write report to sheet
    info = [['Name','Standard','NumVersions','Type','Score','ForA','ForB','ForC','ForD']]
    extras = 4 // forA, forB, forC, forD
    for (var k=0; k<report.length; k++) {
      reportSheet.getRange((assessmentValues.length+extras)*k+1, 1, 1, info[0].length)
          .setValues(info)
          .setBackground("navy")
          .setFontColor("white");
      reportSheet.getRange((assessmentValues.length+extras)*k+2, 1, (assessmentValues.length+extras-1), 1)
          .setValue(report[k][0])
          .merge()
          .setVerticalAlignment('top');
      reportSheet.getRange((assessmentValues.length+extras)*k+2, 2, assessmentValues.length, 3)
          .setValues(assessmentValues);
      reportSheet.getRange((assessmentValues.length+extras)*k+2, 5, assessmentValues.length, 1)
          .setValues(report[k].slice(1, report[k].length));     
      
      //Compute extras from report per student: [name, maxOfStandard...]
      var forA = [];
      var forB = [];
      var forC = [];
      var forD = [];
      for (var i=0; i<assessmentValues.length; i++) {
        var score = (report[k][i+1]!=null)? report[k][i+1]:0
        if (assessmentValues[i][2]=='core') {
          forA.push([(score >= 3)? 'OK':3])
          forB.push([(score >= 3)? 'OK':3])
          forC.push([(score >= 3)? 'OK':3])
          forD.push([(score >= 2)? 'OK':2])
        } else if (assessmentValues[i][2]=='advance') {
          forA.push([(score >= 3)? 'OK':3])
          forB.push([(score >= 2)? 'OK':2])
          forC.push(['-'])
          forD.push(['-'])
        } else {
          forA.push(['?'])
          forB.push(['?'])
          forC.push(['?'])
          forD.push(['?'])
        }
      }
      reportSheet.getRange((assessmentValues.length+extras)*k+2, 6, assessmentValues.length, 1)
          .setValues(forA)
          .setWrap(true);      
      reportSheet.getRange((assessmentValues.length+extras)*k+2, 7, assessmentValues.length, 1)
          .setValues(forB)
          .setWrap(true);
      reportSheet.getRange((assessmentValues.length+extras)*k+2, 8, assessmentValues.length, 1)
          .setValues(forC)
          .setWrap(true);
      reportSheet.getRange((assessmentValues.length+extras)*k+2, 9, assessmentValues.length, 1)
          .setValues(forD)
          .setWrap(true);
    }
    // Beautification
    reportSheet.getDataRange().setHorizontalAlignment('left');
    reportSheet.deleteColumn(3);
    reportSheet.deleteColumns(12, reportSheet.getMaxColumns()-12);
    reportSheet.setColumnWidth(1, 150);  
    reportSheet.setColumnWidth(2, 100);  
    reportSheet.setColumnWidth(3, 100);  
    reportSheet.setColumnWidth(4, 75);  
    reportSheet.setColumnWidth(5, 50);
    reportSheet.setColumnWidth(6, 50);
    reportSheet.setColumnWidth(7, 50);
    reportSheet.setColumnWidth(8, 50);
    reportSheet.getRange(1, 9).setValue('Updated:' + (new Date()));
  } else {
    throw "Grade sheet does not exist.";
  }
  return "Report sheet successfully created.";
}

function moveData(e) {
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var rawSheet = activeSpreadsheet.getSheetByName("Raw");
  var gradeSheet = activeSpreadsheet.getSheetByName("Grade");

  var rawValues = rawSheet.getDataRange().getValues();  
  
  for (var k=0; k<rawValues.length; k++) {
    for (var j=0; j<rawValues[0].length; j++) {
      tmp = [[rawValues[k][j],null,null,null,null,null,null,null,null,null]]
      gradeSheet.getRange(4+k, 1+2+j*11, 1, tmp[0].length).setValues(tmp);
    }
  }
}





