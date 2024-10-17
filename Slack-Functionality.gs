function doPost(e) {
  var params = e.parameter; // e.parameter contains the slash command parameters
  getEmployersData();
  // List all the keys in the 'params' object
  var keys = Object.keys(params);

  // Log all keys and their values
  for (var i = 0; i < keys.length; i++) {
    Logger.log("Key: " + keys[i] + ", Value: " + params[keys[i]]);
  }

  // Respond back to Slack with the list of keys and values
  var response = {
    text:
      "Received the following parameters:\n" +
      keys
        .map(function (key) {
          return key + ": " + params[key];
        })
        .join("\n"),
  };

  var coDays = sheet2Map[String(params.user_name + "@windowsreport.com")][4];

  return ContentService.createTextOutput(
    "Salut, mai ai " + coDays + " zile de concediu ramase!"
  ).setMimeType(ContentService.MimeType.JSON);
}
