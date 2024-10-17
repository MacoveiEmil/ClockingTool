function createNextMonthSheet() {
  // Get the current spreadsheet
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  // Get today's date and determine the next month
  var today = new Date();
  var nextMonth = new Date(today.getFullYear(), today.getMonth() + 3);
  Logger.log(today);
  Logger.log(nextMonth);
  // Format the next month's name
  var nextMonthName = monthNames[nextMonth.getMonth()];

  var currentYear = today.getFullYear();

  // Create the new sheet with the next month's name
  var sheet = spreadsheet.insertSheet(
    nextMonthName + " - " + String(currentYear)
  );

  // Define the static columns
  var staticColumnsSecondRow = [
    "Nr crt",
    "Punct de\nlucru",
    "Norma de\nlucru",
    "Nume si\nPrenume",
    "Data\nangajare",
  ];
  var staticColumnsFirstRow = ["", "", "", "", ""];
  // Get the number of days in the next month
  var daysInMonth = new Date(
    nextMonth.getFullYear(),
    nextMonth.getMonth() + 3,
    0
  ).getDate();
  Logger.log(daysInMonth);
  // Create an array for the header row
  var headerFirstRow = staticColumnsFirstRow.slice();
  var headerSecondRow = staticColumnsSecondRow.slice();

  // Add the day columns to the header row
  for (var i = 1; i <= daysInMonth; i++) {
    var day = new Date(nextMonth.getFullYear(), nextMonth.getMonth(), i);
    var dayName = day.toLocaleDateString("ro-RO", { weekday: "short" }); // e.g., "Su" for Sunday
    headerFirstRow.push(dayName);
    headerSecondRow.push(i);
  }
  headerFirstRow.push(
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    ""
  );
  headerSecondRow.push(
    "Total ore\nlucrate",
    "ON",
    "OW",
    "OSN",
    "SUPL",
    "Total ore\nnelucrate",
    "Zile de\nconcediu\nDE PLATIT",
    "Zile de\nconcediu\nLUATE IN AVANS",
    "",
    "CO",
    "CM",
    "CIC",
    "AB",
    "I",
    "CFP",
    "D",
    "CS",
    "CP"
  );
  // Set the header row in the new sheet
  sheet.getRange(1, 1, 1, headerFirstRow.length).setValues([headerFirstRow]);
  sheet.getRange(2, 1, 1, headerSecondRow.length).setValues([headerSecondRow]);

  // Optionally, format the header row to make it stand out
  sheet
    .getRange(1, 1, 2, headerFirstRow.length)
    .setFontWeight("bold")
    .setBackground("#333f4f")
    .setFontColor("#FFFFFF"); // Header background color (green in this case);
  // sheet.getRange(2, 1, 1, headerSecondRow.length).setFontWeight("bold");

  // Get data from "Sheet2", excluding the header row
  var sourceSheet = spreadsheet.getSheetByName("Angajati - 2024");
  var sourceData = sourceSheet
    .getRange(2, 1, sourceSheet.getLastRow() - 1, sourceSheet.getLastColumn())
    .getValues();

  // Add rows to the new sheet
  for (var j = 0; j < sourceData.length; j++) {
    var row = [];
    row.push(j + 1); // Incremental number
    row.push(sourceData[j][2]); // Sheet2.ColumnC
    row.push(sourceData[j][3]); // Sheet2.ColumnD
    row.push(sourceData[j][0]); // Sheet2.ColumnA
    row.push(sourceData[j][4]); // Sheet2.ColumnE

    // Add 0 for working days and "" for weekends
    for (var i = 1; i <= daysInMonth; i++) {
      var day = new Date(nextMonth.getFullYear(), nextMonth.getMonth(), i);
      var isWeekend = day.getDay() === 0 || day.getDay() === 6; // Sunday or Saturday
      var isHoliday = holidays.some(function (holiday) {
        return holiday.getTime() === day.getTime();
      });
      row.push(isWeekend || isHoliday ? "" : 0);
    }

    // Append the row to the new sheet
    sheet.appendRow(row);

    var lastRow = sheet.getLastRow();
    var startColumn = 6; // Assuming day columns start at column F
    var endColumn = startColumn + daysInMonth - 1;
    var startColumnLetter = columnToLetter(startColumn);
    var endColumnLetter = columnToLetter(endColumn);

    var startColumnLetterAbsente = columnToLetter(endColumn + 10);
    var endColumnLetterAbsente = columnToLetter(endColumn + 18);

    var formula = `=SUM(${startColumnLetter}${lastRow}:${endColumnLetter}${lastRow})`;
    var formulaNelucrate = `=${sourceData[j][3]}*SUM(${startColumnLetterAbsente}${lastRow}:${endColumnLetterAbsente}${lastRow})`;
    var formulaCO = `=COUNTIF(${startColumnLetter}${lastRow}:${endColumnLetter}${lastRow}, "CO")`;
    var formulaCM = `=COUNTIF(${startColumnLetter}${lastRow}:${endColumnLetter}${lastRow}, "CM")`;
    var formulaCIC = `=COUNTIF(${startColumnLetter}${lastRow}:${endColumnLetter}${lastRow}, "CIC")`;
    var formulaAB = `=COUNTIF(${startColumnLetter}${lastRow}:${endColumnLetter}${lastRow}, "AB")`;
    var formulaI = `=COUNTIF(${startColumnLetter}${lastRow}:${endColumnLetter}${lastRow}, "I")`;
    var formulaCFP = `=COUNTIF(${startColumnLetter}${lastRow}:${endColumnLetter}${lastRow}, "CFP")`;
    var formulaD = `=COUNTIF(${startColumnLetter}${lastRow}:${endColumnLetter}${lastRow}, "D")`;
    var formulaCS = `=COUNTIF(${startColumnLetter}${lastRow}:${endColumnLetter}${lastRow}, "CS")`;
    var formulaCP = `=COUNTIF(${startColumnLetter}${lastRow}:${endColumnLetter}${lastRow}, "CP")`;

    sheet
      .getRange(lastRow, endColumn + 1)
      .setFormula(formula)
      .setFontWeight("bold")
      .setBackground("#a8d08d");
    sheet.getRange(lastRow, endColumn + 2).setBackground("#e2efd9");
    sheet.getRange(lastRow, endColumn + 3).setBackground("#e2efd9");
    sheet.getRange(lastRow, endColumn + 4).setBackground("#e2efd9");
    sheet.getRange(lastRow, endColumn + 5).setBackground("#e2efd9");
    sheet
      .getRange(lastRow, endColumn + 6)
      .setFormula(formulaNelucrate)
      .setBackground("#a8d08d");
    sheet.getRange(lastRow, endColumn + 7).setBackground("#a8d08d");
    sheet.getRange(lastRow, endColumn + 8).setBackground("#a8d08d");
    sheet.getRange(lastRow, endColumn + 10).setFormula(formulaCO);
    sheet.getRange(lastRow, endColumn + 11).setFormula(formulaCM);
    sheet.getRange(lastRow, endColumn + 12).setFormula(formulaCIC);
    sheet.getRange(lastRow, endColumn + 13).setFormula(formulaAB);
    sheet.getRange(lastRow, endColumn + 14).setFormula(formulaI);
    sheet.getRange(lastRow, endColumn + 15).setFormula(formulaCFP);
    sheet.getRange(lastRow, endColumn + 16).setFormula(formulaD);
    sheet.getRange(lastRow, endColumn + 17).setFormula(formulaCS);
    sheet.getRange(lastRow, endColumn + 18).setFormula(formulaCP);
  }

  // Auto-resize the columns to fit the content

  sheet
    .getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns())
    .setFontSize(9)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");

  mergeFirstTwoRows();

  sheet.autoResizeColumns(1, headerFirstRow.length);
  // Set border for all data range
  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();
  sheet.getRange(3, 1, lastRow - 2, 5).setBackground("#f3f9fa");
  sheet
    .getRange(3, 1, lastRow - 2, lastColumn)
    .setBorder(true, true, true, true, true, true);

  sheet.getRange(3, lastColumn - 9, lastRow - 2, 10).setBackground("#e2efd9"); // Background color for last column (amber in this case)

  Logger.log(daysInMonth);
}

// Helper function to convert column index to column letter
function columnToLetter(column) {
  var temp,
    letter = "";
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

function mergeFirstTwoRows() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // Define the columns to merge
  var columnsToMerge = [];

  // Add the first 5 columns (A-E)
  for (var i = 1; i <= 5; i++) {
    columnsToMerge.push(i);
  }

  // Add the last 10 columns
  var lastColumn = sheet.getLastColumn();
  for (var j = lastColumn - 17; j <= lastColumn; j++) {
    columnsToMerge.push(j);
  }

  // Remove duplicate columns (if any)
  columnsToMerge = [...new Set(columnsToMerge)];

  // Merge the first two rows for each specified column
  for (var k = 0; k < columnsToMerge.length; k++) {
    var column = columnsToMerge[k];
    var range = sheet.getRange(1, column, 2, 1); // Range from row 1 to row 2 in the specified column
    range.merge(); // Merge the cells
  }
}
function copyRowsWithFormatting() {
  // Get the active spreadsheet
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  // Get the "Main" sheet and the "Second" sheet
  var mainSheet = spreadsheet.getSheetByName("Model");
  var secondSheet = spreadsheet.getSheetByName("September - 2024");

  // Get the first 3 rows from the "Main" sheet
  var rangeToCopy = mainSheet.getRange(1, 1, 3, mainSheet.getLastColumn());

  // Get values and formatting (formulas, background, font, borders, etc.)
  var dataToCopy = rangeToCopy.getValues();
  var backgroundToCopy = rangeToCopy.getBackgrounds();
  var fontStylesToCopy = rangeToCopy.getFontStyles();
  var fontSizesToCopy = rangeToCopy.getFontSizes();
  var fontColorsToCopy = rangeToCopy.getFontColorObjects();
  // var bordersToCopy = rangeToCopy.getBorders();
  var numberFormatsToCopy = rangeToCopy.getNumberFormats();
  var formulasToCopy = rangeToCopy.getFormulas();

  // Find the last row with data in the "Second" sheet
  var lastRow = secondSheet.getLastRow();

  // Calculate the target row (3 rows after the last row)
  var targetRow = lastRow + 4; // +4 because we need 3 blank rows

  // Define the range in the "Second" sheet where data will be pasted
  var targetRange = secondSheet.getRange(
    targetRow,
    1,
    3,
    mainSheet.getLastColumn()
  );

  // Set the values and formatting

  targetRange.setBackgrounds(backgroundToCopy);
  targetRange.setFontStyles(fontStylesToCopy);
  targetRange.setFontSizes(fontSizesToCopy);
  targetRange.setFontColors(fontColorsToCopy);
  // targetRange.setBorders(bordersToCopy);
  targetRange.setNumberFormats(numberFormatsToCopy);
  targetRange.setFormulas(formulasToCopy);
  targetRange.setValues(dataToCopy);
}
