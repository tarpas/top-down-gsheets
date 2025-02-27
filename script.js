var ui;
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

function insert_date(i, j, range_values, number_format)
{
  if (range_values[i][j] == '') {
    sheet.getRange(i+1, j+1).setNumberFormat(number_format);
    sheet.getRange(i+1, j+1).setValue(new Date());
  }
}

function copy_formulas_from_row3_to_row2() {
  update_all();
  
  // Check if we have at least 3 rows
  if (lastRow < 3) {
    return;
  }
  
  // Get formulas from row 3
  var row3Range = sheet.getRange(3, 1, 1, lastColumn);
  var row3Formulas = row3Range.getFormulas()[0];
  
  // Check each cell in row 3 for formulas
  for (var col = 0; col < row3Formulas.length; col++) {
    if (row3Formulas[col] !== '') {
      // Instead of directly copying the formula, we'll use the built-in
      // Google Sheets functionality to handle reference adjustments
      var sourceCell = sheet.getRange(3, col + 1);
      var targetCell = sheet.getRange(2, col + 1);
      
      // This will copy the formula with proper reference adjustments
      sourceCell.copyTo(targetCell, SpreadsheetApp.CopyPasteType.PASTE_FORMULA, false);
    }
  }
}

function insert_top_empty_row(i, j, rangeValues)
{                                           
  var number_format = sheet.getRange(i+1, j+1).getNumberFormat();
  if (!(i > 0 && is_row_empty(rangeValues[i-1]))) {
    sheet.insertRowBefore(i+1);
    // After inserting a row, copy formulas from row 3 to row 2
    copy_formulas_from_row3_to_row2();
  }
}

function onEdit(e)
{
  update_all();
  
  var rangeValues = searchRange.getValues();
  var date_occurence = find_date_occurence(rangeValues);
  var i = date_occurence[0];
  var j = date_occurence[1];

  if (i == -1)
    return;
  
  var range = e.range;
  var row = range.getRow();
  var number_format = sheet.getRange(i+1, j+1).getNumberFormat();
    
  // quick workaround because if user manually inserts
  // an empty row, the last row triggers onEdit which
  // adds date to the row with sum(money)
  if (row != lastRow)
    insert_date(row-1, j, rangeValues, number_format);
}

function onOpen()
{
  update_all();

  ui.createMenu('Timesheet tools')
      .addItem('Reverse order', 'reverse_order')
      .addToUi();
  
  var rangeValues = searchRange.getValues();

  var date_occurence = find_date_occurence(rangeValues);
  var i = date_occurence[0];
  var j = date_occurence[1];

  insert_top_empty_row(i, j, rangeValues);
}

function every_hour_trigger()
{
  var rangeValues = searchRange.getValues();
  var date_occurence = find_date_occurence(rangeValues);
  var i = date_occurence[0];
  var j = date_occurence[1];
  
  insert_top_empty_row(i, j, rangeValues);
}

