//ReadMe
//Make sure to create triggers for the script
##Triggers to be created
![Request Form Triggers](http://screencast.com/t/jFwBdCnon8q)

//The members on the Trello board and the members assigned for each category should be the same.

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
      .showModalDialog(html, 'Link your Request Form with Expense Sheet & your Trello Board');
}


function onSettingsSubmit(formObj){
  var scriptProperties = PropertiesService.getScriptProperties();
  var requestUrl = formObj.request_url;
  scriptProperties.setProperty('REQUEST_FORM_URL', requestUrl);
  var expenseUrl = formObj.expense_url;
  scriptProperties.setProperty('EXPENSE_SHEET_URL', expenseUrl);

  var memberName1 = formObj.member_name1;
  var memberId1 = formObj.member_Id1;
  var memberName2 = formObj.member_name2;
  var memberId2 = formObj.member_Id2;
  var memberName3 = formObj.member_name3;
  var memberId3 = formObj.member_Id3;
  
  
  var trelloMembersMapping = {};
  trelloMembersMapping[memberName1] = memberId1;
  trelloMembersMapping[memberName2] = memberId2;
  trelloMembersMapping[memberName3] = memberId3;
  
  scriptProperties.setProperty('TRELLO_MEMBERS_MAP', JSON.stringify(trelloMembersMapping));
    
  var AppKey = formObj.app_key;
  scriptProperties.setProperty('TRELLO_APP_KEY', AppKey);
  var UserToken = formObj.user_token;
  scriptProperties.setProperty('TRELLO_USER_TOKEN', UserToken);
  var list_Id = formObj.list_Id;
  scriptProperties.setProperty('TRELLO_LIST_ID', list_Id);
  return;
}


function updateRequestForm(e) {
  Utilities.sleep(3000);
  SpreadsheetApp.flush();
  var scriptProperties = PropertiesService.getScriptProperties();
  var requestFormUrl = scriptProperties.getProperty('REQUEST_FORM_URL');
  var expenseSheetUrl = scriptProperties.getProperty('EXPENSE_SHEET_URL');
  var form = FormApp.openByUrl(requestFormUrl); //Open the request form 
  var items = form.getItems(FormApp.ItemType.LIST); 
  var listItem = items[0].asListItem();
  Logger.log(listItem.getTitle());
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

function OnFormSubmit(e) {
 var scriptProperties = PropertiesService.getScriptProperties();
 
 var trelloMemberIdMapping =  JSON.parse(scriptProperties.getProperty('TRELLO_MEMBERS_MAP'));
 Logger.log("on form submit");
 Logger.log(trelloMemberIdMapping); 
  var approverName = "";
  var stringName = "";
  var resultName = "";
  var itemResponses = e.response.getItemResponses();

  var title = itemResponses[1].getResponse();
  var dueDate = itemResponses[4].getResponse();
  itemResponses.splice(1,1);
  var description = "";
  for(var i=0; i<itemResponses.length; i++){
    Logger.log("--" + itemResponses[i].getResponse());
      description = description + "**" + itemResponses[i].getItem().getTitle() + "**: " + itemResponses[i].getResponse() + "\n";
    if(i==2){
     stringName = itemResponses[i].getResponse();
      Logger.log(stringName);
     resultName = stringName.split(" ");
     approverName = resultName[0];
     Logger.log(approverName);
      
    }
      
  }
  createTrelloCard(title, description, trelloMemberIdMapping[approverName], dueDate)
}


//Steps
  //Get the api key from the link [https://trello.com/app-key]
  //Use the api key in this link to give trello access to your script [https://trello.com/1/authorize?key=<app key>&name=MyApp&scope=read,write&expiration=never&response_type=token]
  //This link will return a token. Copy the token and use it in the script below as user token.
  //Get all the lists of a particular org. visible board [board id to be taken from the board url] from this link 
  //https://api.trello.com/1/boards/<substitutewithboardid>?lists=open&list_fields=name&fields=name,desc&key=<apikey>&token=<usertoken>
  

function createTrelloCard(title, description, trelloMemberId, date) {
    var scriptProperties = PropertiesService.getScriptProperties();
    var KEY = scriptProperties.getProperty('TRELLO_APP_KEY'); // app key of the account which was used to create the Trello board
    var scriptProperties = PropertiesService.getScriptProperties();
    var USER_TOKEN = scriptProperties.getProperty('TRELLO_USER_TOKEN'); //user token of the Trello board 
    var scriptProperties = PropertiesService.getScriptProperties();
    var listID = scriptProperties.getProperty('TRELLO_LIST_ID');  //ID of the list on which you want to create the request card
 
  var url = "https://api.trello.com/1/cards?key=" + KEY + "&token=" + USER_TOKEN;
  var cardInfo = {"name": title,               
                 "desc": description,  
                 "pos":"top",          
                 "due": date,          
                 "idList":listID,  
                 "labels": "",  
                 "idMembers": trelloMemberId,       
                 "idCardSource": "",     
                 "keepFromSource": "all",
                };
  
    
  var options = {"method" : "POST",
                 "payload": cardInfo
                };
      var response = UrlFetchApp.fetch(url, options);
  
      
 
  
  //Logger.log(response);
}



