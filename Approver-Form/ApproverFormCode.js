

function onOpen() {
  FormApp.getUi() // Or DocumentApp or FormApp.
      .createMenu('BuyIt')
      .addItem('Settings', 'openDialog')
      .addToUi();
}

function openDialog() {
  var html = HtmlService.createHtmlOutputFromFile('dialog')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  FormApp.getUi() // Or DocumentApp or FormApp.
      .showModalDialog(html, 'Link your Approver Form with Expense Sheet & your Trello Board');
}

function onSettingsSubmit(formObj){
  var scriptProperties = PropertiesService.getScriptProperties();
  var approverUrl = formObj.approver_url;
  scriptProperties.setProperty('APPROVER_FORM_URL', approverUrl);
  var expenseUrl = formObj.expense_url;
  scriptProperties.setProperty('EXPENSE_SHEET_URL', expenseUrl);
  var AppKey = formObj.app_key;
  scriptProperties.setProperty('TRELLO_APP_KEY', AppKey);
  var UserToken = formObj.user_token;
  scriptProperties.setProperty('TRELLO_USER_TOKEN', UserToken);
  var list_Id = formObj.list_Id;
  scriptProperties.setProperty('TRELLO_LIST_ID', list_Id);
  return;
}

function updateApproverForm(e) {
  Utilities.sleep(3000);
  SpreadsheetApp.flush();
 // var range = e.range;
  //Logger.log(range.getRow());
  var scriptProperties = PropertiesService.getScriptProperties();
  var approverFormUrl = scriptProperties.getProperty('APPROVER_FORM_URL');
  var expenseSheetUrl = scriptProperties.getProperty('EXPENSE_SHEET_URL');
  var form = FormApp.openByUrl(approverFormUrl); //Open the approver form 
  var items = form.getItems(FormApp.ItemType.LIST); 
  var listItem = items[1].asListItem();
  Logger.log(listItem.getTitle());
  Logger.log(expenseSheetUrl);
 var ss = SpreadsheetApp.openByUrl(expenseSheetUrl); //URL of the Expense Sheet
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

function onFormSubmit(e) {
  var scriptProperties = PropertiesService.getScriptProperties();
  var approverFormUrl = scriptProperties.getProperty('APPROVER_FORM_URL');
  var form = FormApp.openByUrl(approverFormUrl);
  var itemResponse = e.response.getItemResponses();
  var trelloCard = itemResponse[2].getResponse();
  var selectOption = itemResponse[13].getResponse();  
  var recurring = itemResponse[11].getResponse();
  var approver = itemResponse[0].getResponse();
  var commentResponse = itemResponse[14].getResponse();
  var comment = "**" + approver + "**" + ": " + commentResponse;
  var getCard = trelloCard.split("/");
  var getCardId = getCard[4];
  Logger.log(getCardId);
  Logger.log(recurring);
createComment(comment, getCardId);
  
moveTrelloCard(getCardId, recurring);
  }
    
  function moveTrelloCard(getCardId, recurring) {
    var scriptProperties = PropertiesService.getScriptProperties();
    var KEY = scriptProperties.getProperty('TRELLO_APP_KEY'); //app key of the account which was used to create the Trello board
    var scriptProperties = PropertiesService.getScriptProperties();
    var USER_TOKEN = scriptProperties.getProperty('TRELLO_USER_TOKEN'); //user token of the Trello board 
    Logger.log(getCardId);
    var scriptProperties = PropertiesService.getScriptProperties();
    var listId1 = scriptProperties.getProperty('TRELLO_LIST_ID');  //ID of the list to which you want the card to be moved
    Logger.log(listId1);
    if (recurring == "No") {
       var url = "https://api.trello.com/1/cards/" + getCardId + "/idList?key=" + KEY + "&token=" + USER_TOKEN + "&value=" + listId1; //http://stackoverflow.com/questions/20644927/trello-api-how-do-i-move-a-card-to-a-different-list
    
var options = {
       "method": "PUT",
       "oAuthServiceName": "trello",
       "oAuthUseToken": "always",
       "payload1": listId1,
  
        };
var response = UrlFetchApp.fetch(url, options);
    

  Logger.log(response);
    }
}
function createComment(comment, getCardId) {
  Logger.log(getCardId);
  Logger.log(comment);
  var url = "https://api.trello.com/1/cards/" + getCardId + "/actions/comments";
  var commentInfo = {
    "text": comment,
    "key": "fd287be53aea157e4bc05c991f06c005", //BuyIt's app key
    "token": "ed1082da5a54f4b19c9f9f8e717241ffdba2091c3b30fd98c19ef87699461011" //BuyIt's token
  };     
 Logger.log(commentInfo);
var options = {
       "method": "POST",
       "oAuthServiceName": "trello",
       "oAuthUseToken": "always",
       "payload": commentInfo
 
        };
  Logger.log(options);
  Logger.log(url);
var response = UrlFetchApp.fetch(url,options);
Logger.log(response);   
  
}   
