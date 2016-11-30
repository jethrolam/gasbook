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
 * Initialize Grade and Report sheet from Assessment sheets
 */
function create() {
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var gradeSheet = activeSpreadsheet.getSheetByName("Grade");
  var reportSheet = activeSpreadsheet.getSheetByName("Report");
  var debug = false;

  if (gradeSheet != null) {
    if (!debug) {
      throw "Cannot overwrite existing Grade sheet"
    } else {
      gradeSheet.clear();
    }
  } else {
    gradeSheet = activeSpreadsheet.insertSheet("Grade");
  }

  if (reportSheet != null) {
    if (!debug) {
      throw "Cannot overwrite existing Report sheet"
    } else {
      reportSheet.clear();
    }
  } else {
    reportSheet = activeSpreadsheet.insertSheet("Report");
  }

  // Init
  var numStudents = 40;
  var gradeHeaderBackgroundColor = 'maroon';
  var gradeHeaderFontColor = 'white';
  var reportHeaderBackgroundColor = 'navy';
  var reportHeaderFontColor = 'white';

  // Read from Standard sheet
  var standardValues = activeSpreadsheet.getSheetByName("Standard")
    .getDataRange()
    .getValues();
  var numStandards = standardValues.length;
  var numScores = 0
  for (var i = 0; i < numStandards; i++) {
    numScores += standardValues[i][1] + 1 // add one as version M
  }

  // Crop Grade sheet size
  var margins = 10
  var gradeSheetMaxColumns = numScores + 2 + margins;
  var gradeSheetMaxRows = numStudents + 3 + margins;
  if (gradeSheet.getMaxRows() > gradeSheetMaxRows) {
    gradeSheet.deleteRows(gradeSheetMaxRows, gradeSheet.getMaxRows() - gradeSheetMaxRows) // delete extra rows
  } else if (gradeSheet.getMaxRows() < gradeSheetMaxRows) {
    gradeSheet.insertRowsAfter(gradeSheet.getMaxRows(), gradeSheetMaxRows - gradeSheet.getMaxRows()) // insert extra rows
  }
  if (gradeSheet.getMaxColumns() > gradeSheetMaxColumns) {
    gradeSheet.deleteColumns(gradeSheetMaxColumns, gradeSheet.getMaxColumns() - gradeSheetMaxColumns) // delete extra cols
  } else if (gradeSheet.getMaxColumns() < gradeSheetMaxColumns) {
    gradeSheet.insertColumnsAfter(gradeSheet.getMaxColumns(), gradeSheetMaxColumns - gradeSheet.getMaxColumns()) // insert extra cols
  }

  // Set Grade sheet headers and formula
  offset_col = 2
  versionMA1 = [] // keep track of A1 notation of verison M columns, one for each standard
  for (var i = 0; i < numStandards; i++) {
    var nameStandard = standardValues[i][0]
    var numVersions = standardValues[i][1] + 1 // add one as version M
    var typeStandard = standardValues[i][2]

    if (gradeSheet.getMaxColumns() < offset_col + numVersions) {
      gradeSheet.insertColumnsAfter(gradeSheet.getMaxColumns(), numVersions)
    }

    gradeSheet.getRange(1, offset_col + 1, 1, numVersions)
      .merge()
      .setValue(nameStandard)
      .setHorizontalAlignment('center')
      .setBackground(gradeHeaderBackgroundColor)
      .setFontColor(gradeHeaderFontColor);
    gradeSheet.getRange(2, offset_col + 1, 1, numVersions)
      .merge()
      .setValue(typeStandard)
      .setHorizontalAlignment('center')
      .setBackground(gradeHeaderBackgroundColor)
      .setFontColor(gradeHeaderFontColor);
    for (var j = 1; j < numVersions; j++) {
      gradeSheet.getRange(3, offset_col + 1 + j)
        .setValue(j)
        .setHorizontalAlignment('center')
        .setBackground(gradeHeaderBackgroundColor)
        .setFontColor(gradeHeaderFontColor);
    }
    gradeSheet.getRange(3, offset_col + 1)
      .setValue('M')
      .setHorizontalAlignment('center')
      .setBackground(gradeHeaderBackgroundColor)
      .setFontColor(gradeHeaderFontColor);

    // Record A1 notation for version M columns
    versionMA1.push([gradeSheet.getRange(4, offset_col + 1, numStudents, 1)
      .getA1Notation()
      .split('4:')[0]
    ]);

    // Set formula for version M
    gradeSheet.getRange(4, offset_col + 1, numStudents, 1)
      .setFormulaR1C1("=MAX(R[0]C[1]:R[0]C[" + (numVersions - 1) + "])")

    offset_col += numVersions
  }

  // Grade sheet beautification
  for (var i = 1; i < gradeSheet.getMaxColumns() - 2; i++) {
    gradeSheet.setColumnWidth(i + 2, 30)
  }
  gradeSheet.setColumnWidth(1, 100);
  gradeSheet.setColumnWidth(2, 100);
  gradeSheet.setFrozenColumns(2);
  gradeSheet.getRange(3, 1).setValue('Name')
    .setBackground(gradeHeaderBackgroundColor)
    .setFontColor(gradeHeaderFontColor);
  gradeSheet.getRange(3, 2).setValue('Email')
    .setBackground(gradeHeaderBackgroundColor)
    .setFontColor(gradeHeaderFontColor);
  for (var k = 0; k < numStudents; k++) {
    gradeSheet.getRange(4 + k, 1).setValue('<student ' + (k + 1) + '>')
    gradeSheet.getRange(4 + k, 2).setValue('<email ' + (k + 1) + '>')
  }

  // Report Header
  info = [
    ['Name', 'Standard', 'NumVersions', 'Type', 'Score', 'ForA', 'ForB', 'ForC', 'ForD', 'ToDoA', 'ToDoB', 'ToDoC', 'ToDoD']
  ]

  // Crop Report sheet size
  var margins = 10
  var reportSheetMaxColumns = info[0].length + margins;
  var reportSheetMaxRows = numStudents * (1 + numStandards) + margins;
  if (reportSheet.getMaxRows() > reportSheetMaxRows) {
    reportSheet.deleteRows(reportSheetMaxRows, reportSheet.getMaxRows() - reportSheetMaxRows) // delete extra rows
  } else if (reportSheet.getMaxRows() < reportSheetMaxRows) {
    reportSheet.insertRowsAfter(reportSheet.getMaxRows(), reportSheetMaxRows - reportSheet.getMaxRows()) // insert extra rows
  }
  if (reportSheet.getMaxColumns() > reportSheetMaxColumns) {
    reportSheet.deleteColumns(reportSheetMaxColumns, reportSheet.getMaxColumns() - reportSheetMaxColumns) // delete extra cols
  } else if (reportSheet.getMaxRows() < reportSheetMaxRows) {
    reportSheet.insertColumnsAfter(reportSheet.getMaxColumns(), reportSheetMaxColumns - reportSheet.getMaxColumns()) // insert extra cols
  }

  // Create Report headers and formula
  extras = 3 // extra rows between students
  for (var k = 0; k < numStudents; k++) {
    reportSheet.getRange((extras + numStandards) * k + 1, 1, 1, info[0].length)
      .setValues(info)
      .setBackground(reportHeaderBackgroundColor)
      .setFontColor(reportHeaderFontColor);
    reportSheet.getRange((extras + numStandards) * k + 2, 1)
      .setFormula(gradeSheet.getName() + "!" +
        gradeSheet.getRange(k + 4, 1)
        .getA1Notation());
    reportSheet.getRange((extras + numStandards) * k + 3, 1)
      .setFormula(gradeSheet.getName() + "!" +
        gradeSheet.getRange(k + 4, 2)
        .getA1Notation());
    reportSheet.getRange((extras + numStandards) * k + 2, 2, numStandards, 3)
      .setValues(standardValues);

    // Set Version M here using formula
    for (var i = 0; i < numStandards; i++) {
      var studentVersionMA1 = versionMA1[i] + (k + 4)
      reportSheet.getRange((extras + numStandards) * k + 2 + i, 5)
        .setFormula('=' + gradeSheet.getName() + '!' + studentVersionMA1);
    }

    // Compute forABCD
    reportSheet.getRange((extras + numStandards) * k + 2, 6, numStandards, 1)
      .setFormulaR1C1('=if(C[-1]>=3, "OK", 3)');
    reportSheet.getRange((extras + numStandards) * k + 2, 7, numStandards, 1)
      .setFormulaR1C1('=if(C[-3]="core", if(C[-2]>=3, "OK", 3), if(C[-2]>=2, "OK", 2))');
    reportSheet.getRange((extras + numStandards) * k + 2, 8, numStandards, 1)
      .setFormulaR1C1('=if(C[-4]="core", if(C[-3]>=3, "OK", 3), "-")');
    reportSheet.getRange((extras + numStandards) * k + 2, 9, numStandards, 1)
      .setFormulaR1C1('=if(C[-5]="core", if(C[-4]>=2, "OK", 2), "-")');

    // Compute ToDoABCD
    reportSheet.getRange((extras + numStandards) * k + 2, 10, numStandards, 1)
      .setFormulaR1C1('=if(or(C[-4]="OK",C[-4]="-"), "-", if(C[-3]="OK",C[-8],"-"))');
    reportSheet.getRange((extras + numStandards) * k + 2, 11, numStandards, 1)
      .setFormulaR1C1('=if(or(C[-4]="OK",C[-4]="-"), "-", if(C[-3]="OK",C[-9],"-"))');
    reportSheet.getRange((extras + numStandards) * k + 2, 12, numStandards, 1)
      .setFormulaR1C1('=if(or(C[-4]="OK",C[-4]="-"), "-", if(C[-3]="OK",C[-10],"-"))');
    reportSheet.getRange((extras + numStandards) * k + 2, 13, numStandards, 1)
      .setFormulaR1C1('=if(or(C[-4]="OK",C[-4]="-"), "-", C[-11])');

  }

  // Beautification
  reportSheet.getDataRange().setHorizontalAlignment('left');
  reportSheet.setColumnWidth(1, 150);
  reportSheet.setColumnWidth(2, 100);
  reportSheet.setColumnWidth(3, 25);
  reportSheet.setColumnWidth(4, 100);
  reportSheet.setColumnWidth(5, 75);
  reportSheet.setColumnWidth(6, 50);
  reportSheet.setColumnWidth(7, 50);
  reportSheet.setColumnWidth(8, 50);
  reportSheet.setColumnWidth(9, 50);
  reportSheet.setColumnWidth(10, 100);
  reportSheet.setColumnWidth(11, 100);
  reportSheet.setColumnWidth(12, 100);
  reportSheet.setColumnWidth(13, 100);

  return "Grade and Report sheet successfully created";
}
