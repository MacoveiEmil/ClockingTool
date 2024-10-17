function coProcessing(e) {
  getEmployersData();
  // Get the active sheet where the edit was made
  const editedSheet = e.source.getActiveSheet();

  // Check if the edited sheet's name matches the target sheet name
  if (editedSheet.getName() === "Cereri Concedii") {
    // Get the new value that was entered
    const value = e.value;

    // Run your custom functions, passing the range and new value
    Logger.log(value);
    Logger.log("e.range.getColumn() " + e.range.getColumn());
    Logger.log("e.range.getRow() " + e.range.getRow());
    Logger.log(
      ("e.range.getColumn() == 6 && value == true : " + e.range.getColumn() ==
        6 + " - " + value) ==
        true
    );
    if (e.range.getColumn() == 6 && value == "TRUE") {
      var data = ss
        .getSheetByName(editedSheet.getName())
        .getRange(e.range.getRow(), 1, 1, 10)
        .getValues();
      Logger.log(data);
      processFormResponses(data, e.range.getRow());
    }
  }
}

function forumSubmit(e) {
  getEmployersData();
  // Get the active sheet where the edit was made
  const editedSheet = e.source.getActiveSheet();

  // Check if the edited sheet's name matches the target sheet name
  if (editedSheet.getName() === "Cereri Concedii") {
    // Get the new value that was entered
    const values = ss
      .getSheetByName(editedSheet.getName())
      .getRange(
        e.range.getRow(),
        e.range.getColumn(),
        1,
        e.range.getLastColumn()
      )
      .getValues();

    // Run your custom functions, passing the range and new value
    Logger.log(values);
    if (
      values[0][4] == "Concediu Odihna" &&
      values[0][5] == false &&
      values[0][6].length > 0 &&
      values[0][7] == null
    ) {
      var scheduledVacationDays = getscheduledVacationDays(values[0][0]);
      Logger.log("scheduledVacationDays: " + scheduledVacationDays);
      Logger.log(
        "coDays: " +
          Number(
            Number(sheet2Map[values[0][0]][2]) +
              Number(sheet2Map[values[0][0]][3])
          )
      );
      var numberOfCODays = checkCODays(
        Utilities.formatDate(values[0][2], "GMT", "yyyy-MM-dd"),
        Utilities.formatDate(values[0][3], "GMT", "yyyy-MM-dd"),
        Number(
          Number(sheet2Map[values[0][0]][2]) +
            Number(sheet2Map[values[0][0]][3])
        ),
        scheduledVacationDays
      );
      renameDriveFile(
        values[0][6].replace("https://drive.google.com/open?id=", ""),
        coType[values[0][4]][0] +
          " - " +
          sheet2Map[values[0][0]][0] +
          " - " +
          Utilities.formatDate(values[0][2], "GMT", "MMMM-YYYY")
      );
      Logger.log("numberOfCODays: " + numberOfCODays);
      if (numberOfCODays > 0) {
        Logger.log("Concediu Aprobat");
        ss.getSheetByName(editedSheet.getName())
          .getRange(e.range.getRow(), e.range.getLastColumn() + 2, 1, 1)
          .setValue("Verificata si aprobata");
        ss.getSheetByName(editedSheet.getName())
          .getRange(e.range.getRow(), e.range.getLastColumn() + 3, 1, 1)
          .setValue(numberOfCODays);
        emailForCO(
          sheet2Map[values[0][0]][0],
          values[0][0],
          Utilities.formatDate(values[0][2], "GMT", "dd-MM-yyyy"),
          Utilities.formatDate(values[0][3], "GMT", "dd-MM-yyyy"),
          true,
          false
        );
      } else {
        Logger.log("Concediu Neaprobat");
        ss.getSheetByName(editedSheet.getName())
          .getRange(e.range.getRow(), e.range.getLastColumn() + 2, 1, 1)
          .setValue("Verificata si neaprobata");
        emailForCO(
          sheet2Map[values[0][0]][0],
          values[0][0],
          Utilities.formatDate(values[0][2], "GMT", "dd-MM-yyyy"),
          Utilities.formatDate(values[0][3], "GMT", "dd-MM-yyyy"),
          false,
          false
        );
      }
    } else {
      Logger.log("conditiile nu sunt indelinite");
      if (values[0][4] != "Concediu Odihna") {
        renameDriveFile(
          values[0][6].replace("https://drive.google.com/open?id=", ""),
          coType[values[0][4]][0] +
            " - " +
            sheet2Map[values[0][0]][0] +
            " - " +
            Utilities.formatDate(values[0][2], "GMT", "MMMM-YYYY")
        );
      }
    }
  }
}

function renameDriveFile(fileId, newName) {
  try {
    // Get the file by ID
    var file = DriveApp.getFileById(fileId);

    // Rename the file
    file.setName(newName);

    Logger.log("File renamed to: " + newName);
  } catch (e) {
    Logger.log("Error: " + e.toString());
  }
}

function checkCODays(startDate, endDate, coDays, scheduledVacationDays) {
  let sumOfCODays = 0;

  if (!startDate || !endDate) return;

  // Convert startDate and endDate to Date objects
  let start = new Date(startDate);
  let end = new Date(endDate);

  // Loop from start date to end date
  while (start <= end) {
    let currentYear = start.getFullYear();
    let currentMonth = start.getMonth() + 1; // getMonth() is zero-based
    let currentDay = start.getDate();

    // Check if the current day is not a weekend
    if (!isWeekend(currentYear + "-" + currentMonth + "-" + currentDay)) {
      sumOfCODays += 1;
    }

    // Move to the next day
    start.setDate(start.getDate() + 1);
  }

  Logger.log("sumOfCODays: " + sumOfCODays);

  if (sumOfCODays + scheduledVacationDays > coDays) {
    Logger.log("Nu sunt suficiente zile de concediu");
    return 0;
  } else {
    Logger.log("Sunt suficiente zile de concediu");
    return sumOfCODays;
  }
}

function createTrigger() {
  var sheet = SpreadsheetApp.getActive();
  ScriptApp.newTrigger("coProcessing").forSpreadsheet(sheet).onEdit().create();
}

function getscheduledVacationDays(employerEmail) {
  var vacantionDays = 0;
  for (var j = 1; j < concediiSheetData.length; j++) {
    if (
      concediiSheetData[j][0] == employerEmail &&
      concediiSheetData[j][8] == "Verificata si aprobata"
    ) {
      vacantionDays += concediiSheetData[j][9];
    }
  }
  return vacantionDays;
}
