function translateSheetData(sourceSheetUrl, sourceSheetName, sourceRange, targetSheetUrl, targetSheetName, targetRange, targetLanguage) {
  try {
    // Log start of function
    Logger.log("Starting translateSheetData function");

    // Open the source and target spreadsheets
    Logger.log("Opening source spreadsheet: " + sourceSheetUrl);
    var sourceSpreadsheet = SpreadsheetApp.openByUrl(sourceSheetUrl);

    Logger.log("Opening target spreadsheet: " + targetSheetUrl);
    var targetSpreadsheet = SpreadsheetApp.openByUrl(targetSheetUrl);

    // Get the source and target sheets
    Logger.log("Getting source sheet: " + sourceSheetName);
    var sourceSheet = sourceSpreadsheet.getSheetByName(sourceSheetName);

    Logger.log("Getting target sheet: " + targetSheetName);
    var targetSheet = targetSpreadsheet.getSheetByName(targetSheetName);

    // Get the source data
    Logger.log("Fetching data from source range: " + sourceRange);
    var sourceData = sourceSheet.getRange(sourceRange).getValues();

    // Translate the source data with delay
    Logger.log("Translating data to target language: " + targetLanguage);
    var translatedData = sourceData.map(function(row) {
      return row.map(function(cell) {
        if (cell) {
          try {
            var translatedCell = LanguageApp.translate(cell, '', targetLanguage);
            Utilities.sleep(1000); // Add delay between translations
            return translatedCell;
          } catch (e) {
            Logger.log("Error translating cell: " + cell + " Error: " + e.message);
            return cell;
          }
        }
        return cell;
      });
    });

    // Clear any non-locked cells in the target range
    Logger.log("Clearing non-locked cells in target range: " + targetRange);
    var targetRangeObj = targetSheet.getRange(targetRange);
    var targetBackgrounds = targetRangeObj.getBackgrounds();
    for (var i = 0; i < targetBackgrounds.length; i++) {
      for (var j = 0; j < targetBackgrounds[i].length; j++) {
        if (targetBackgrounds[i][j] !== "#c1ffd1") {
          targetRangeObj.getCell(i + 1, j + 1).clearContent();
        }
      }
    }

    // Paste the translated data into the target range
    Logger.log("Pasting translated data into target range");
    targetRangeObj.setValues(translatedData);

    // Log successful completion
    Logger.log("translateSheetData function completed successfully");
  } catch (e) {
    // Log any unexpected errors
    Logger.log("Error in translateSheetData function: " + e.message);
    MailApp.sendEmail({
      to: "your-email@example.com",
      subject: "Error in translateSheetData function",
      body: "An error occurred: " + e.message
    });
    throw e;
  }
}
