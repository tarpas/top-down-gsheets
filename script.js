var ui = SpreadsheetApp.getUi();
var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheet = ss.getActiveSheet();
var rangeData = sheet.getDataRange();
var lastColumn = rangeData.getLastColumn();
var lastRow = rangeData.getLastRow();
var searchRange = sheet.getRange(1, 1, lastRow, lastColumn);

function update_all()
{
  ui = SpreadsheetApp.getUi();
  ss = SpreadsheetApp.getActiveSpreadsheet();
  sheet = ss.getActiveSheet();
  rangeData = sheet.getDataRange();
  lastColumn = rangeData.getLastColumn();
  lastRow = rangeData.getLastRow();
  searchRange = sheet.getRange(1, 1, lastRow, lastColumn);
}

function is_row_empty(row)
{
  for(var i = 0; i < row.length; i++) {
    if (!(row[i] instanceof Date) && row[i] != '')
      return false;
  }
  return true;
}

function find_date_occurence(range_values)
{
  for(var i = 0; i < lastRow; i++) {
    for (var j = 0 ; j < lastColumn; j++) {    
      if (range_values[i][j] instanceof Date) {
        return [i, j];
      }
    }
  }
  return [-1, -1];
}

function reverse_order()
{
  var range_values = searchRange.getValues();
  
  var date_occurence = find_date_occurence(range_values);
  var i = date_occurence[0];
  var j = date_occurence[1];
  
  if (i == -1)
    return;
  
  var dates = range_values.slice(i);
  dates.reverse();
  
  sheet.getRange(i+1, 1, dates.length, lastColumn).setValues(dates);
}

function insert_top_row(i, j, rangeValues)
{
  if (i != -1 && is_row_empty(rangeValues[i])) {
    sheet.getRange(i+1, j+1).setValue(new Date());
  } else {
    var insert_row = !(i > 0 && is_row_empty(rangeValues[i-1]));

    var number_format = sheet.getRange(i+1, j+1).getNumberFormat();
    if (insert_row)
      sheet.insertRowBefore(i+1);

    sheet.getRange(i+1, j+1).setNumberFormat(number_format);
    sheet.getRange(i+1, j+1).setValue(new Date());
  }
}

function onEdit()
{
  update_all();
  
  // if the top row with date is filled, add new emtpy row
  var rangeValues = searchRange.getValues();
  var date_occurence = find_date_occurence(rangeValues);
  var i = date_occurence[0];
  var j = date_occurence[1];
  
  if (i == -1)
    return;  
  
  for (var k = j; k < lastColumn; k++) {
    if (rangeValues[i][k] == '')
      return;
  }
  insert_top_row(i, j, rangeValues);
}

function onOpen()
{
  ui.createMenu('Timesheet tools')
      .addItem('Reverse order', 'reverse_order')
      .addToUi();
  
  var rangeValues = searchRange.getValues();

  var date_occurence = find_date_occurence(rangeValues);
  var i = date_occurence[0];
  var j = date_occurence[1];

  // if the row with the date is empty, just update the date
  // otherwise create a new row with current date
  insert_top_row(i, j, rangeValues);
  
  return;
}

