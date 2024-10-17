function findRowByColumnDValue(valueToFind, sheetName) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName); // Replace with your sheet name
  var dataRange = sheet.getRange("A:D"); // Get all data in column D
  var data = dataRange.getValues(); // Get all the values in column D

  for (var i = 0; i < data.length; i++) {
    if (data[i][0] == valueToFind || data[i][3] == valueToFind) {
      var rowNumber = i + 1; // Rows are 1-indexed in Google Sheets
      return rowNumber; // Return the row number where the match is found
    }
  }

  return -1; // Return -1 if no match is found
}

function getEmployersData() {
  for (var i = 1; i < sheet2Data.length; i++) {
    var sheet2Value = sheet2Data[i][1]; // Column B value
    var sheet2Key = sheet2Data[i][0];
    var sheet2SecondKey = sheet2Data[i][3]; // Column A value
    var sheet2CurrentYearCODays = sheet2Data[i][7];
    var sheet2LastYearCODays = sheet2Data[i][6];
    var sheet2coDaysLast = sheet2Data[i][21];
    sheet2Map[sheet2Value] = [
      sheet2Key,
      sheet2SecondKey,
      sheet2CurrentYearCODays,
      sheet2LastYearCODays,
      sheet2coDaysLast,
    ];
  }
}

function isWeekend(dateStr) {
  // Parse the date string into a Date object
  var dateParts = dateStr.split("-");
  var year = parseInt(dateParts[0], 10);
  var month = parseInt(dateParts[1], 10) - 1; // Month is 0-based in JavaScript
  var day = parseInt(dateParts[2], 10);

  var date = new Date(year, month, day);
  Logger.log("date: " + date);
  Logger.log("date.getDay(): " + date.getDay());
  // Check if the day is Saturday (6) or Sunday (0)
  var isWeekend = date.getDay() === 0 || date.getDay() === 6;

  return isWeekend;
}

function emailForCO(
  emplyerName,
  emailAddresses,
  startDate,
  endDate,
  approved,
  useCase,
  coDaysLeft
) {
  var subject = "";
  var message = "<HTML><Body>";
  message += `Salut, ${emplyerName}<br><br>`;

  if (useCase == "Aprobare cerere de concediu") {
    subject = "Cerere de concediu aprobata";
    message += `Cererea de concediu pentru perioada ${startDate} - ${endDate} a fost aprobata.<br>`;
    message += `Mai aveti disponibile ${coDaysLeft} zile de concediu pentru acest an.<br><br>`;
  } else {
    if (approved) {
      subject = "Cerere de concediu inregistrata";
      message += `Cererea de concediu pentru perioada ${startDate} - ${endDate} a fost inregistrata si urmeaza a fi aprobata.<br><br>`;
    } else {
      subject = "Cerere de concediu respinsa";
      message += `Cererea de concediu pentru perioada ${startDate} - ${endDate} nu a putut fi inregistrata pe motiv: "Zile insuficiente de concediu ramase."<br>`;
      message += `Va rugam sa contactati echipa HR pentru clarificarea situatiei.<br><br>`;
    }
  }
  message += `Zi faina,<br>HR Team`;
  message += "</Body></HTML>";
  Logger.log(MailApp.getRemainingDailyQuota());
  MailApp.sendEmail(emailAddresses, subject, "", {
    cc: "rose_itu@windowsreport.com",
    htmlBody: message,
  });
}
