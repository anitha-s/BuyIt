function onOpen() {
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
      .createMenu('BuyIt')
      .addItem('Settings', 'openDialog')
      .addToUi();
}

function openDialog() {
  var html = HtmlService.createHtmlOutputFromFile('dialog')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
      .showModalDialog(html, 'This is to link your Request & Approver Forms with your Expense Sheet.');
}

function onSettingsSubmit(formObj){
  var scriptProperties = PropertiesService.getScriptProperties();
  var approverUrl = formObj.approver_url;
  scriptProperties.setProperty('APPROVER_FORM_URL', approverUrl);
  var requestUrl = formObj.request_url;
  scriptProperties.setProperty('REQUEST_FORM_URL', requestUrl);
  var expenseUrl = formObj.expense_url;
  scriptProperties.setProperty('EXPENSE_SHEET_URL', expenseUrl);
  return;
}

function updateForms(e){
  Utilities.sleep(3000);
  calFunction(e);
  updateApproverForm(e);
  updateRequestForm(e);
}


function updateApproverForm(e) {
  Utilities.sleep(3000);
  SpreadsheetApp.flush();
 var scriptProperties = PropertiesService.getScriptProperties();
  var approverFormUrl = scriptProperties.getProperty('APPROVER_FORM_URL');
  var expenseSheetUrl = scriptProperties.getProperty('EXPENSE_SHEET_URL');
  var form = FormApp.openByUrl(approverFormUrl); //Open the approver form 
  
  //var form = FormApp.openByUrl('https://docs.google.com/a/multunus.com/forms/d/1kjTnGzvbfDWIVYon6D_pr5fQiUWKLtORn8zDU7BAbVI/edit'); //Open the Open Source approver form in which the edit has to be shown.
  var items = form.getItems(FormApp.ItemType.LIST); 
  var listItem = items[1].asListItem();
  Logger.log(listItem.getTitle());
  var ss = SpreadsheetApp.openById(expenseSheetUrl);
  var sheet = ss.getSheetByName("Summary");
  var data = sheet.getSheetValues(4, 1, 28, 1);
  var filteredData = [];
  var doNotUseRegex = /.*do\snot\suse.*/i;
  for(var index = 0; index < data.length; index++) {
    if(doNotUseRegex.exec(data[index]) == null){
     filteredData.push(data[index]);
    }
  }
  listItem.setChoiceValues(filteredData);
  
}

function updateRequestForm(e) {
  Utilities.sleep(3000);
  SpreadsheetApp.flush();
  var scriptProperties = PropertiesService.getScriptProperties();
  var requestFormUrl = scriptProperties.getProperty('REQUEST_FORM_URL');
  var expenseSheetUrl = scriptProperties.getProperty('EXPENSE_SHEET_URL');
  var form = FormApp.openByUrl(requestFormUrl); //Open the approver form 
 
  //var form = FormApp.openByUrl('https://docs.google.com/a/multunus.com/forms/d/1Ut5fbFQnKVgdoS9mH-l91J8BHnO6wcdLyhFg9QnS96g/edit'); //Open the Open SOurce Request form in which the edit has to be shown.
  var items = form.getItems(FormApp.ItemType.LIST); 
  var listItem = items[0].asListItem();
  Logger.log(listItem.getTitle());
  var ss = SpreadsheetApp.openByUrl(expenseSheetUrl);
  var sheet = ss.getSheetByName("Summary");
  var data = sheet.getSheetValues(4, 1, 28, 1);
  var filteredData = [];
  var doNotUseRegex = /.*do\snot\suse.*/i;
  for(var index = 0; index < data.length; index++) {
    if(doNotUseRegex.exec(data[index]) == null){
     filteredData.push(data[index]);
    }
  }
  listItem.setChoiceValues(filteredData);
  
}

function calFunction(e) {
  var scriptProperties = PropertiesService.getScriptProperties();
  var expenseSheetUrl = scriptProperties.getProperty('EXPENSE_SHEET_URL');
  var ss = SpreadsheetApp.openByUrl(expenseSheetUrl);
  var summarySheet = ss.getSheetByName("Summary");
  var summaryCostCategoryCell = summarySheet.getRange("A4:A31");
  summaryCostCategoryCell.setFormula("=CONCATENATE(J4, \" [\",B4,\"] [Budget Left: Rs.\",TEXT(F4,\"##,###\"),\"] [Limit:\",\"Rs.\",TEXT(K4,\"##,###\"),\"]\")");
  var range = summarySheet.getRange("B4:B31");
  var data = range.getValues();
  for (var i=0; i < data.length; i++){
      var costCategoryCell = data[i];
      Logger.log(costCategoryCell);
      var summarySheetApportionedExpensedCell = summarySheet.getRange("H4:H31").getCell(i+1, 1);
    summarySheetApportionedExpensedCell.setFormula("=SUMIFS('BuyIt Form Entries'!F:F,'BuyIt Form Entries'!H:H,\"*"+costCategoryCell+"*\",'BuyIt Form Entries'!M:M,\"Yes\",'BuyIt Form Entries'!B:B,B1)");
      
  
  var summarySheetExpensedCell = summarySheet.getRange("E4:E31");
  summarySheetExpensedCell.setFormula("=SUMIFS('BuyIt Form Entries'!F:F,'BuyIt Form Entries'!B:B,$B$1,'BuyIt Form Entries'!H:H,CONCATENATE(\"*\", B4, \"*\"))-H4");
}  
}
