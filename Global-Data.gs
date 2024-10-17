var ss = SpreadsheetApp.getActiveSpreadsheet();

var sheetAngajati = ss.getSheetByName("Angajati - 2024");
var sheet2Data = sheetAngajati.getDataRange().getValues();

var concediiSheet = ss.getSheetByName("Cereri Concedii");
var concediiSheetData = concediiSheet.getDataRange().getValues();
var sheet2Map = {};

var daysForFirstMonth = 0;
var daysForSecondMonth = 0;

var coColor;

var monthNames = [
  "January",
  "February",
  "March",
  "April",
  "May",
  "June",
  "July",
  "August",
  "September",
  "October",
  "November",
  "December",
];

var holidays = [
  // 2024 Holidays
  new Date(2024, 0, 1), // New Year's Day
  new Date(2024, 0, 2), // Day after New Year's Day
  new Date(2024, 0, 24), // Unification Day
  new Date(2024, 3, 14), // Orthodox Good Friday
  new Date(2024, 3, 16), // Orthodox Easter Sunday
  new Date(2024, 3, 17), // Orthodox Easter Monday
  new Date(2024, 4, 1), // Labour Day
  new Date(2024, 5, 20), // Orthodox Pentecost
  new Date(2024, 5, 21), // Orthodox Pentecost Monday
  new Date(2024, 7, 15), // St. Mary’s Day
  new Date(2024, 11, 1), // National Day
  new Date(2024, 11, 25), // Christmas Day
  new Date(2024, 11, 26), // Day after Christmas Day

  // 2025 Holidays
  new Date(2025, 0, 1), // New Year's Day
  new Date(2025, 0, 2), // Day after New Year's Day
  new Date(2025, 0, 24), // Unification Day
  new Date(2025, 3, 4), // Orthodox Good Friday
  new Date(2025, 3, 6), // Orthodox Easter Sunday
  new Date(2025, 3, 7), // Orthodox Easter Monday
  new Date(2025, 4, 1), // Labour Day
  new Date(2025, 5, 8), // Orthodox Pentecost
  new Date(2025, 5, 9), // Orthodox Pentecost Monday
  new Date(2025, 7, 15), // St. Mary’s Day
  new Date(2025, 11, 1), // National Day
  new Date(2025, 11, 25), // Christmas Day
  new Date(2025, 11, 26), // Day after Christmas Day
];

var coType = {
  "Concediu Odihna": ["CO", "#00b050"],
  "Concediu medical": ["CM", "#ff0000"],
  "Crestere copil": ["CIC", "#f3f3e1"],
  Invoiri: ["I", "#ff99ff"],
  "Concediu fara plata": ["CFP", "#09bfff"],
  Delegatie: ["D", "#f4b083"],
  Absente: ["AB", "#7030a0"],
  "Concediu zile speciale": ["CS", "#ffff00"],
  "Concediu parental": ["CP", "#aa9208"],
};
