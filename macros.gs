import 'google-apps-script'

var ui = SpreadsheetApp.getUi();
var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();  
var lastRow = sheet.getLastRow();
var lastColumn = sheet.getLastColumn();
var range = sheet.getDataRange();
var values = range.getValues();
var spreadsheet = SpreadsheetApp.getActive();
var lookupRange = sheet.getRange(3,3, lastRow, lastColumn);
var lookupRangeValues = lookupRange.getValues();

function onOpen(){
  goToEnd();
  //docNavigation();
  analyticsByLetters();
}

function doGet(e) {
  console.log(e);
  //var html = HtmlService.createHtmlOutputFromFile("page");
  var html = HtmlService.createTemplateFromFile("page").evaluate();
  return html;
}

//function docNavigation(){
  
  //ui.createMenu('Навигация')
      //.addItem('Перейти к заявке...','goToRow')
      //.addItem('Перейти к последней заявке','goToEnd')
      //.addItem('Перейти к первой незакрытой заявке', 'goToFirstFalseOrder')
    //  .addToUi();
  //}

function analyticsByLetters(){
  
  ui.createMenu('Проверка')
      .addItem('Наличие писем', 'showFormSidebarAnalytics')
      //.addItem('Переходы по заявкам', 'showFormSidebarNavigation')
      .addToUi();
  }

function goToRow(rowNumber)
{ 
  //var rowNumber = Browser.inputBox('Order number');
  //var columnNumber = 1;
  var selectedCell = sheet.getRange(rowNumber, 1);
  sheet.setCurrentCell(selectedCell);
  return true;
}

  function goToEnd() { 
  var lastCell = sheet.getRange(getFirstEmptyRowWholeRow(), 1);
  sheet.setCurrentCell(lastCell);
}

function goToFirstFalseOrder() {
  //Browser.msgBox(goToFirstFalse());
  var lastCell = sheet.getRange(goToFirstFalse(), 1);
  sheet.setCurrentCell(lastCell);
}

function getFirstEmptyRowWholeRow() {
  var row = 0;
  for (var row=3; row<values.length; row++) {
    if(values[row][0] == '' && values[row][1] == '' && values[row][2] == '' && values[row][3] == '') break;
  }
  return (row);
}

function goToFirstFalse() {
  var row = 0;
  for (var row=3; row<values.length; row++) {
    if(values[row][0] == '') break;
  }
  return (row+1);
}

function showFormSidebarAnalytics() {
  var html = HtmlService.createHtmlOutputFromFile('analytics_page')
    .setTitle('Check letters')
    .setWidth(300);
  SpreadsheetApp.getUi()
    .showSidebar(html);
}

function showFormSidebarNavigation() {
  var html = HtmlService.createHtmlOutputFromFile('navigation_page')
    .setTitle('Navigation')
    .setWidth(300);
  SpreadsheetApp.getUi()
    .showSidebar(html);
}

function SearchNotLetters(letters)
{
  var tf, all; 
  var missingLetters = letters;
  for(n = 0; n<letters.length; n++){
  tf = spreadsheet.createTextFinder(letters[n]);
  all = tf.findAll();
  for (var i = 0; i < all.length; i++) {    
    missingLetters[n]="";
    }
  }
  missingLetters = missingLetters.filter(Boolean);
return missingLetters;
}