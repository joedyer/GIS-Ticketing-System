function refreshTeamOptions(){
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Teams');
  var lastrow = sheet.getLastRow();
  var GISrange = sheet.getRange(2, 1, lastrow-1, 2).getValues();
  var ENGrange = sheet.getRange(2, 5, lastrow-1, 2).getValues();
  
  GISrange = cutBlanks(GISrange);
  ENGrange = cutBlanks(ENGrange);
  
  updateOptions('1skE9ECS29_XWNdWyGZy01ko-LiCTyIaQ9CjMpNe6iNQ', '705630236', GISrange);
  updateOptions('1eK7ptBkVD2WgPM664e89V-uEGA6P5xb88HIGkZJObsA', '307946213', ENGrange);
}

function cutBlanks(arr){
  var tempArr = [];
  for(var i=0; i<arr.length; i++){
    if(arr[i][0] != ""){
      tempArr.push(arr[i])
    }
  }
  return tempArr;
}

function updateOptions(formId, itemId, arr){
  
  var form = FormApp.openById(formId);
  var item = form.getItemById(itemId).asMultipleChoiceItem();
  
  var choices = [];
  
  for(var i = 0; i < arr.length; i++){
    choices.push(item.createChoice(arr[i][0]+" "+arr[i][1]));
  }
  
  item.setChoices(choices)
}

function test(){
  var form = FormApp.openById('1skE9ECS29_XWNdWyGZy01ko-LiCTyIaQ9CjMpNe6iNQ');
}