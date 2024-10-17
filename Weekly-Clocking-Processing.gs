var days = 0;
var currentDay;

function processSheetData() {
  // Step 1: Automatically detect the active sheet
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
    "Week 30-09-2024 / 30-09-2024"
  );

  // Get all data from the sheet
  var data = sheet.getDataRange().getValues();

  // Step 2: Process the data
  var result = [];

  // Loop through each row
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    days = 0;
    // The first column is not important, so we start with the second one
    var identifier = row[1];

    // Loop through the columns in groups of 6, starting from the 3rd column (index 2)
    for (var j = 2; j < row.length; j += 1) {
      // Extract the group of 6 columns
      var group = row.slice(j, j + 1);

      // Save the identifier and the group of data
      result.push([identifier].concat(group));
      days = days + 1;
    }
  }

  getEmployersData();

  Logger.log(sheet2Map);

  // Step 3: Save the result (e.g., log it, or save it to another sheet)
  Logger.log(days);
  Logger.log(result);
  for (var j = 0; j < days; j++) {
    currentDay = null;
    for (var i = 0; i < result.length / days; i++) {
      let index = i * days + j;
      if (currentDay == null) {
        var dateStr = String(result[index][1])
          .replace("How did you spend your time this week? [", "")
          .replace("]", "")
          .replace(/ - /g, "-");
        var parts = dateStr.split("-");
        currentDay = new Date(parts[2], parts[1] - 1, parts[0]);
      }
      Logger.log("index=" + index);
      Logger.log("Current Row: " + result[index]);
      Logger.log("CurrentDay= " + currentDay);
      if (
        result[index][0] != "Email Address" &&
        result[index][1] == "Working Day"
      ) {
        updateWoringDates(
          sheet2Map[result[index][0]][0],
          currentDay,
          sheet2Map[result[index][0]][1]
        );
      }
    }
  }
}

function updateWoringDates(employName, date, hours) {
  if (!date) return;
  Logger.log(date);
  var day = date.getDate();
  var month = date.getMonth();
  var year = date.getFullYear();
  var sheet = ss.getSheetByName(monthNames[month] + " - " + year);
  var rowFromTable = findRowByColumnDValue(
    employName,
    monthNames[month] + " - " + year
  );

  Logger.log(employName);
  Logger.log("day: " + day);

  sheet.getRange(rowFromTable, day + 5, 1, 1).setValue(hours);
}
