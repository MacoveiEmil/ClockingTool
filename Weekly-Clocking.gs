function createWeeklyForm() {
  // If testDate is not provided, use today's date
  // var today = new Date();
  var today = new Date(2024, 09, 11);
  Logger.log(today);
  var dayOfWeek = today.getDay(); // 0 = Sunday, 1 = Monday, ..., 6 = Saturday
  var daysToSubtract = dayOfWeek === 0 ? 6 : dayOfWeek - 1; // Adjust for Monday as the first day
  var monday = new Date(today);
  monday.setDate(today.getDate() - daysToSubtract);

  // Adjust the date if today is the last day of the month and it's a working day
  var lastDayOfMonth = new Date(today.getFullYear(), today.getMonth() + 1, 0);
  var sectionsToAdd = 5; // Default number of sections

  if (
    today.getDate() === lastDayOfMonth.getDate() &&
    dayOfWeek >= 1 &&
    dayOfWeek <= 5
  ) {
    sectionsToAdd = dayOfWeek; // Only add sections for the days of the current week
  }

  if (dayOfWeek == 5 || today.getDate() === lastDayOfMonth.getDate()) {
    var form = FormApp.create("Weekly Work Report");

    // Automatically collect the respondent's email address
    form.setCollectEmail(true);
    form.setAllowResponseEdits(false);

    var firstSectionDate = null;
    var lastSectionDate = null;

    // Create the multiple-choice grid question
    var gridQuestion = form.addGridItem();
    gridQuestion.setTitle("How did you spend your time this week?");

    var questionDate = null;
    var choices = [
      "Working Day",
      "Free Day (Any type of leave: CO, paid day off)",
    ];
    var rows = [];

    // Add the sections and fill the rows with dates
    for (var i = 0; i < sectionsToAdd; i++) {
      questionDate = new Date(monday);
      questionDate.setDate(monday.getDate() + i);

      if (questionDate.getMonth() === today.getMonth()) {
        var formattedDate = Utilities.formatDate(
          questionDate,
          Session.getScriptTimeZone(),
          "dd-MM-yyyy"
        );
        if (!firstSectionDate) firstSectionDate = formattedDate;
        // Record the first and last section dates for the tab name
        lastSectionDate = formattedDate;

        rows.push(formattedDate); // Add the formatted date as a row in the grid
      }
    }

    // Set the rows (dates) and columns (choices) for the grid question
    gridQuestion.setRows(rows);
    gridQuestion.setColumns(choices);

    // Link the form to this sheet
    var formDestinationId = ss.getId();
    form.setDestination(FormApp.DestinationType.SPREADSHEET, formDestinationId);

    // Log the URL to the form for easy access
    Logger.log("Form URL: " + form.getPublishedUrl());

    Utilities.sleep(5000);

    ScriptApp.newTrigger("setWeeklyTabName")
      .timeBased()
      .after(60 * 1000) // 1 minute
      .create();
  } else {
    Logger.log("Not Today, Maybe Tomorow!");
  }
}

function setWeeklyTabName() {
  var today = new Date();
  var dayOfWeek = today.getDay(); // 0 = Sunday, 1 = Monday, ..., 6 = Saturday
  var daysToSubtract = dayOfWeek === 0 ? 6 : dayOfWeek - 1; // Adjust for Monday as the first day
  var monday = new Date(today);
  monday.setDate(today.getDate() - daysToSubtract);

  // Adjust the date if today is the last day of the month and it's a working day
  var lastDayOfMonth = new Date(today.getFullYear(), today.getMonth() + 1, 0);
  var sectionsToAdd = 5; // Default number of sections

  if (
    today.getDate() === lastDayOfMonth.getDate() &&
    dayOfWeek >= 1 &&
    dayOfWeek <= 5
  ) {
    sectionsToAdd = dayOfWeek; // Only add sections for the days of the current week
  }

  var firstSectionDate = null;
  var lastSectionDate = null;

  var questionDate = null;
  // Add the sections and fill the rows with dates
  for (var i = 0; i < sectionsToAdd; i++) {
    questionDate = new Date(monday);
    questionDate.setDate(monday.getDate() + i);

    if (questionDate.getMonth() === today.getMonth()) {
      var formattedDate = Utilities.formatDate(
        questionDate,
        Session.getScriptTimeZone(),
        "dd-MM-yyyy"
      );
      if (!firstSectionDate) firstSectionDate = formattedDate;
      lastSectionDate = formattedDate;
    }
  }

  // Create the custom tab name
  var tabName = "Week " + firstSectionDate + " / " + lastSectionDate;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var formSheet;

  // Find the form responses sheet
  for (var i = 0; i < sheets.length; i++) {
    Logger.log(sheets[i].getName());
    if (sheets[i].getName().startsWith("Form Responses")) {
      formSheet = sheets[i];
      formSheet.setName(tabName);
      break;
    }
  }
}
