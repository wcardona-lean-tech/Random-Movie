var totalMovies = 0;

function randomMovie() {
  var spreadsheet = SpreadsheetApp.getActive();
  var sheet = spreadsheet.getActiveSheet();
  var range = sheet.getRange('B2:D18'); 
  var data = range.getValues();
  
  if(totalMovies > (data.length * 3)){
    Browser.msgBox('Next Movie', 'There are no pending movies!' , Browser.Buttons.OK);
    return;
  }
  
  var column = Math.floor(Math.random()*(data[0].length));//data[0].length: Total of Columns in the range
  var row = Math.floor(Math.random()*(data.length));
  
  var element = sheet.getRange(row + 2, column + 2);
  
  if(element.getBackgroundColor() === '#ffff00'){
    totalMovies++;
    return random();
  }
  
  var movieName = data[row][column]; 
  sheet.getRange(6, 6).setValue(movieName).setBackground("#c1f1a5"); 
  element.setBackground("#ffff00");
  
  Browser.msgBox('Next Movie', movieName, Browser.Buttons.OK);
}
