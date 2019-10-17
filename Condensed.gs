function onNTPFormSubmit(e){
  var sheet = e.source.getActiveSheet();
  var sheetName = sheet.getName();
  
  if(sheetName == "NTP Form Responses"){
    
    //set overview sheet
    var arr = e.range.getValues();
    var NtpSheet = e.source.getSheetByName("NED NTP Receipt");
    var por = arr[0][5];
    arr[0][5] = 1;
    
    //add entries for as many POR as NTP has
    for(var i = 2; i <= por; i++){
      arr.push(arr[0].slice());
      arr[i-1][5] = i;
    }
    
    //set range
    var range = NtpSheet.getRange(NtpSheet.getLastRow()+1, 1, arr.length, 15);
    range.setValues(arr);
    
    //***************************//
    
    //send out notification email
    sendEmail(arr[0][4], sheetName);
    
  } else {
    
    var info = e.range.getValues()[0];
    info.shift(); //removes timestamp
    var ntp = info.shift(); //shifts out the Ntp number
    
    updateReceiptByNTP(ntp,getUpdateCol(sheetName),info);
    
    sendEmail(ntp,sheetName);
    
  } 
}

function getUpdateCol(sheetName){
  switch(sheetName){
    case 'NTP Assignment':
      return 16;
      break;
    case 'GIS Response':
      return 17;
      break;
    case 'Engineering Response':
      return 21;
      break;
    case 'GIS GDB Response':
      return 23;
      break;
    case 'VZ/LCS':
      return 25;
      break;
    default:
      return 1;
      break;
  }
}

function updateReceiptByNTP(ntp, startCol, insert){
 
  var numCol = insert.length;
  var sheet = SpreadsheetApp.getActive().getSheetByName('NED NTP Receipt');
  //get range from start to end of insert
  var range = sheet.getRange(4, 1, (sheet.getLastRow()-3), (startCol+numCol-1));
  var values = range.getValues();
  var flag = false;
  
  //iterate through sheet 
  for(var i = values.length-1; i >= 0; i--){
    
    //match NTPs
    if(values[i][4] == ntp){
      
      //inserting insert array
      for(var j = 0; j < numCol; j++){
        values[i][startCol-1+j] = insert[j];
      }
      
      //NTPs are grouped, once the last is entered, break the loop.
      flag = true;
    }else if(flag){
      break;
    }
  }
  
  range.setValues(values);
}

function sendEmail(ntp, sheetName){
  
  var eI = emailInfo(ntp, sheetName);
  var ntpInfo = getNTPInfo(ntp);
  var body = getInfoHeader(ntpInfo);

  var tmpl = HtmlService.createTemplateFromFile(eI[2]);
  tmpl.ntp = ntp;
  body += tmpl.evaluate().getContent();

  Logger.log(body);
  
  MailApp.sendEmail({
    to: eI[0],
    subject: eI[1],
    htmlBody: body,
  });
}

function emailInfo(ntp, sheetName){
  
//  this.recipient;
//  this.subject;
//  this.template;
  
  switch(sheetName){
      
    case "NTP Form Responses":
      //email going to Manager to assign an NTP to a tech
      return ['joseph.dyer@engineeringassociates.com',"To Assign: "+ntp,'ntp Assignment Template.html'];
      break;
      
    case 'NTP Assignment':
      //email going to tech to notify assignment
      return ['joseph.dyer@engineeringassociates.com',"You've been assigned "+ntp,'ntp assigned.html'];
      break;
      
    case 'GIS Response':
      //email going to engineering to notify completed GIS assignment
      return ['joseph.dyer@engineeringassociates.com',"GIS has finished NTP "+ntp, 'engineering form.html'];
      break;
      
    case 'Engineering Response':
      //email going to tech for GDB
      //find to whom NTP was assigned
      var AssignForm = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('NTP Assignment');
      var AssignValues = AssignForm.getSheetValues(1, 1, AssignForm.getLastRow(), 3);
      var assignedTech;
      for(var i = AssignValues.length-1; i >=0; i--){
        if(AssignValues[i][1] == ntp){
          assignedTech = AssignValues[i][2].split(' ')[2];
          break;
        }
      }
      Logger.log(assignedTech);
      return [assignedTech, "Engineering has finished with "+ntp, 'GIS GDB.html'];
      break;
      
    case 'GIS GDB Response':
      //email to administration to notify GDB upload by GIS team
      return ['joseph.dyer@engineeringassociates.com',"GIS GDB has been submitted for "+ntp, 'VZ/LCS.html'];
      break;
      
    default:
      return;
      break;
  }
}

function getNTPInfo(ntp){
  
  var NTPForm = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('NTP Form Responses');
  var NTPValues = NTPForm.getSheetValues(1, 1, NTPForm.getLastRow(), NTPForm.getLastColumn());
  
  for(var i = NTPValues.length-1; i >= 0; i--){
    if(NTPValues[i][4] == ntp){
      return NTPValues[i];
      break;
    }
  } 
}

function getInfoHeader(NTPinfo){  
  return "<ul><li><b>Recept Date:</b> "+NTPinfo[1]+"</li><li><b>Market:</b> "+NTPinfo[2]+"</li><li><b>Shareable Link:</b> "+NTPinfo[3]+"</li><li><b>NTP Number:</b> "+NTPinfo[4]+"</li><li><b>VZW POR Number:</b> "+NTPinfo[5]+"</li><li><b>Include Customer Drop:</b> "+NTPinfo[6]+"</li><li><b>Wireless Location:</b> "+NTPinfo[7]+"</li><li><b>A LOC Access Point Description:</b> "+NTPinfo[8]+"</li><li><b>A LOC Access Point Lat:</b> "+NTPinfo[9]+"</li><li><b>A LOC Access Point Long:</b> "+NTPinfo[10]+"</li><li><b>Additional Fiber Length for Splice Point:</b> "+NTPinfo[11]+"</li><li><b>Z LOC Lat:</b> "+NTPinfo[12]+"</li><li><b>Z LOC Long:</b> "+NTPinfo[13]+"</li><li><b>NTP Work Description:</b> "+NTPinfo[14]+"</li></ul><br>";
}