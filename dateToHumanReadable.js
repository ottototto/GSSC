// This function adds a custom menu to the Google Sheets UI.
function onOpen() {
    var ui = SpreadsheetApp.getUi();
    // Adds a custom menu with the name "Custom Formats" to the Google Sheets UI.
    ui.createMenu('Custom Formats')
        .addItem('Format Date-Time Cells', 'formatSelectedCells')
        .addToUi();
  }
  
  // This function formats any selected cells containing ISO 8601 date-time strings.
  function formatSelectedCells() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var range = sheet.getActiveRange();
    var values = range.getValues();
    
    for (var i = 0; i < values.length; i++) {
      for (var j = 0; j < values[i].length; j++) {
        var cellValue = values[i][j];
        if (typeof cellValue === 'string' && cellValue.includes('T')) {
          try {
            var formattedDate = formatDate(cellValue);
            if (formattedDate) {
              values[i][j] = formattedDate;
            }
          } catch(e) {
            // If there's an error, log it and leave the cell as-is.
            console.error('Error formatting date:', e);
          }
        }
      }
    }
    
    range.setValues(values);
  }

  function formatDate(dateTimeStr) {
    try {
      var date = new Date(dateTimeStr);
      return Utilities.formatDate(date, Session.getScriptTimeZone(), "MM/dd/yyyy HH:mm:ss");
    } catch (e) {
      return null; // Return null if the date cannot be parsed.
    }
  }