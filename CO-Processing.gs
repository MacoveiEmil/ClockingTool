function processFormResponses(coRequest, row) {
  getEmployersData();

  // Process Form Responses
  var colA = coRequest[0][0]; // Column A
  var colC = Utilities.formatDate(coRequest[0][2], "GMT", "yyyy-MM-dd"); // Column C (Date)
  var colD = Utilities.formatDate(coRequest[0][3], "GMT", "yyyy-MM-dd"); // Column D (Date)
  var colE = coRequest[0][4]; // Column G
  var colF = coRequest[0][5]; // Column G
  var colH = coRequest[0][7]; // Column G

  colE = coType[coRequest[0][4]][0]; // If the term is an abbreviation, return the full term
  coColor = coType[coRequest[0][4]][1];

  if (!colH && colF) {
    // Check if Column G is null
    Logger.log("Here " + sheet2Map[colA][0]);
    var employName = sheet2Map[colA][0];
    Number(sheet2Map[colA][2] + sheet2Map[colA][3]);

    if (employName) {
      concediiSheet.getRange(row, 8).setValue("DA");
      updateDates(employName, colC, colD, colE, coColor);
      let coDaysLeft;
      colE == "CO"
        ? (coDaysLeft =
            sheet2Map[colA][4] - (daysForFirstMonth + daysForSecondMonth))
        : (coDaysLeft = sheet2Map[colA][4]);
      emailForCO(
        employName,
        colA,
        Utilities.formatDate(coRequest[0][2], "GMT", "dd-MM-yyyy"),
        Utilities.formatDate(coRequest[0][3], "GMT", "dd-MM-yyyy"),
        false,
        "Aprobare cerere de concediu",
        coDaysLeft
      );
    }
  }
}

function updateDates(employName, startDate, endDate, coType, coColor) {
  let weHaveTwoDifferentMonths = false;

  if (!startDate || !endDate) return;

  // Initialize date objects
  let start = new Date(startDate);
  let end = new Date(endDate);

  let startMonth = start.getMonth(); // Month of the startDate
  let endMonth = end.getMonth(); // Month of the endDate
  let startYear = start.getFullYear(); // Year of the startDate
  let endYear = end.getFullYear(); // Year of the endDate

  // Determine if we have two different months or the same month
  weHaveTwoDifferentMonths = startMonth !== endMonth || startYear !== endYear;

  // Get the row from the sheet for the employee
  var rowFromMontlyTable = findRowByColumnDValue(
    employName,
    "Angajati - " + startYear
  );

  // Loop through the dates
  while (start <= end) {
    let currentYear = start.getFullYear();
    let currentMonth = start.getMonth(); // zero-based month
    let currentDay = start.getDate();

    let montlySheet = ss.getSheetByName(
      monthNames[currentMonth] + " - " + currentYear
    );

    // Get the correct row for the employee
    let rowFromMontlyTable = findRowByColumnDValue(
      employName,
      monthNames[currentMonth] + " - " + currentYear
    );

    // Check if the current day is not a weekend
    if (!isWeekend(currentYear + "-" + (currentMonth + 1) + "-" + currentDay)) {
      // Set the value and background for the day in the monthly sheet
      montlySheet
        .getRange(rowFromMontlyTable, currentDay + 5, 1, 1)
        .setValue(coType)
        .setBackground(coColor);
      if (currentMonth === startMonth && currentYear === startYear) {
        // Count days for the first month
        daysForFirstMonth++;
      } else if (currentMonth === endMonth && currentYear === endYear) {
        // Count days for the second month
        daysForSecondMonth++;
      }
    }

    // Move to the next day
    start.setDate(start.getDate() + 1);
  }
  if (coType == "CO") {
    if (weHaveTwoDifferentMonths) {
      // Update for the first month
      var newDaysForFirstMonthNumber =
        sheetAngajati
          .getRange(rowFromMontlyTable, startMonth + 9, 1, 1)
          .getValue() + daysForFirstMonth;
      sheetAngajati
        .getRange(rowFromMontlyTable, startMonth + 9, 1, 1)
        .setValue(newDaysForFirstMonthNumber)
        .setBackground(coColor);

      // Update for the second month
      var newDaysForSecondMonthNumber =
        sheetAngajati
          .getRange(rowFromMontlyTable, endMonth + 9, 1, 1)
          .getValue() + daysForSecondMonth;
      sheetAngajati
        .getRange(rowFromMontlyTable, endMonth + 9, 1, 1)
        .setValue(newDaysForSecondMonthNumber)
        .setBackground(coColor);
    } else {
      // If both dates fall in the same month, update for that month
      var newDaysNumber =
        sheetAngajati
          .getRange(rowFromMontlyTable, startMonth + 9, 1, 1)
          .getValue() +
        (daysForFirstMonth + daysForSecondMonth);
      sheetAngajati
        .getRange(rowFromMontlyTable, startMonth + 9, 1, 1)
        .setValue(newDaysNumber)
        .setBackground(coColor);
    }
  }
}
