/**
 * Adds a custom menu to the spreadsheet UI.
 */
function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('Personal Report')
      .addItem('Auto-fill Month and Day', 'autoPopulateMonthAndDay')
      .addToUi();
}

/**
 * Automatically populates the month and day of the week in the report sheet.
 * It also handles months with fewer than 31 days by inserting a hyphen.
 */
function autoPopulateMonthAndDay() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName('sitename'); // Please adjust the sheet name if it is different

  if (!sheet) {
    Browser.msgBox('The sheet "sitename" was not found. Please check the sheet name.');
    return;
  }

  const today = new Date();
  const year = today.getFullYear();
  const month = today.getMonth(); // 0-11
  
  // Enter the month in cell A2
  const monthCell = sheet.getRange('A2');
  monthCell.setValue((month + 1) + 'M');
  
  // Get the last day of the current month
  const lastDayOfMonth = new Date(year, month + 1, 0).getDate();
  
  // Process the range for days and weekdays
  for (let i = 4; i <= 34; i++) { // From row 4 to 34 (Day 1 to 31)
    const day = sheet.getRange(i, 3).getValue(); // Column C (Day)
    const dayCell = sheet.getRange(i, 2); // Column B (Weekday)

    if (day && !isNaN(day)) {
      if (day <= lastDayOfMonth) {
        // If the day is within the month, calculate and enter the weekday
        const date = new Date(year, month, day);
        const weekday = date.getDay(); // 0:Sun, 1:Mon, ..., 6:Sat
        
        let weekdayText;
        switch (weekday) {
          case 0: weekdayText = 'sunday'; break;
          case 1: weekdayText = 'monday'; break;
          case 2: weekdayText = 'tuesday'; break;
          case 3: weekdayText = 'wensday'; break;
          case 4: weekdayText = 'thursday'; break;
          case 5: weekdayText = 'friday'; break;
          case 6: weekdayText = 'saturday'; break;
        }
        
        dayCell.setValue(weekdayText);
      } else {
        // If the day is after the last day of the month, enter a hyphen
        dayCell.setValue('-');
      }
    } else {
        // If there is no value in column C (Day), clear the cell
        dayCell.setValue('');
    }
  }

  Browser.msgBox('Month and day auto-filling is complete.');
