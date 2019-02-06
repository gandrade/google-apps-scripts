var sheet = SpreadsheetApp.getActiveSpreadsheet();
var ui = SpreadsheetApp.getUi();

function setValue(cellRange, value) {
  sheet.getRange(cellRange).setValue(value);
}

function getValue(cellRange) {
  return sheet.getRange(cellRange).getValue();
}

function getNextRow() {
  return sheet.getLastRow() + 1;
}

function getCurrentRow() {
  return sheet.getLastRow();
}

function punchIn() {
  var punchOutRange = 'B' + getCurrentRow()
  
  if (hasPunchedOut(punchOutRange)) {
    ui.alert("You must punch-out first before continuing.");
    return;
  } 
  setValue('A' + getNextRow(), new Date());    
}

function hasPunchedOut(cellRange) {
  return getValue('B' + getCurrentRow()) == ""; 
}
  

function punchOut() {
  var punchInRange = 'A' + getCurrentRow();
  var punchOutRange = 'B' + getCurrentRow();  
  var timeSpentRange = 'C' + getCurrentRow();
  var timeSpentValue = '=' + punchOutRange + '-' + punchInRange;

  setValue(punchOutRange, new Date());
  setValue(timeSpentRange, timeSpentValue);
}