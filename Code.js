//Function to get last populated row at the time of call of the function.
function getLastPopulatedRow(sheet) {
  var spr = sheet;
  var column = spr.getRange('A:A');
  var values = column.getValues(); // get all data in one call
  var ct = 0;
  while ( values[ct][0] != "" ) {
    ct++;
  }
  Logger.log(ct);
  return (ct);
}

function getEnteredData()
{
  var dataEntered = new Array(17);
  dataEntered[0] = UserProperties.getProperty('timestamp');
  dataEntered[1] = UserProperties.getProperty('username');
  dataEntered[2] = UserProperties.getProperty('campus');
  dataEntered[3] = UserProperties.getProperty('teamType');
  dataEntered[4] = UserProperties.getProperty('fullGroupName');
  dataEntered[5] = UserProperties.getProperty('emailAddress')+'@alaska.edu';
  dataEntered[6] = UserProperties.getProperty('description');
  dataEntered[7] = UserProperties.getProperty('natureOfUse');
  dataEntered[8] = UserProperties.getProperty('addNotes');
  dataEntered[9] = UserProperties.getProperty('memEstimate');
  dataEntered[10] = UserProperties.getProperty('initialOwner');  
 /* if(dataEnetered[10]=='Yes')
  {
    dataEntered[10] = 'Yes, '+dataEntered[1]+' is the initial owner.';
  }
  else
  {
    dataEntered[10] = 'The initial owner is: '+dataEntered[10];
  }*/
  if (UserProperties.getProperty('decision')=='Yes')
  {
    dataEntered[11] = 'checked';
  }
  else if (UserProperties.getProperty('decision')=='No')
  {
    dataEntered[12] = 'checked';
  }
  else if (UserProperties.getProperty('decision')=='-' || UserProperties.getProperty('decision')=='Pending')
  {
    dataEntered[13] = 'checked';
  }
  
  if (UserProperties.getProperty('gltype')=='Group')
  {
    dataEntered[14] = 'checked';
  }
  else if (UserProperties.getProperty('gltype')=='List')
  {
    dataEntered[15] = 'checked';
  }
  else if (UserProperties.getProperty('gltype')=='' || UserProperties.getProperty('decision')=='?')
  {
    dataEntered[16] = 'checked';
  }
  return dataEntered;
}


//This is the main function which gets executed when the form is submitted.
function onFormSubmit(event) { 
  //getting the active spreadsheet
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  //get the approval sheet by name from the active spreadsheet
  var sheets = spreadsheet.getSheetByName('Approval Worksheet');
  //Get values that have to be put on the email.
  var timestamp = event.values[0];  //Time of submisssion
  var username = event.values[1];  //The username of the submitter
  var campus = event.values[2];  //The campus of the submitter
  var gltype = event.values[3]; //The group or list type
  var fullname = event.values[4];  //The full name of the Group or List
  var description = event.values[6]; //The description provided for the group or list
  var replacement = event.values[7];  //The group or list that this new group or list is replacing, if any
  var usage = event.values[8];  //The nature of use of the group or list
  var members = event.values[9];   //The number of members estimated for the group or list
  var notes = event.values[11];   //Any notes attached with the request
  var initialowner = event.values[10];  //Reply to the question if the submitter is the initial owner.
  //Setting up all the event info from form as ScriptProperties
  ScriptProperties.setProperty('timestamp', timestamp);
  ScriptProperties.setProperty('username', username);
  ScriptProperties.setProperty('campus', campus);
  ScriptProperties.setProperty('gltype', gltype);
  ScriptProperties.setProperty('fullname', fullname);
  ScriptProperties.setProperty('description', description);  
  ScriptProperties.setProperty('replacement', replacement);  
  ScriptProperties.setProperty('usage', usage);  
  ScriptProperties.setProperty('members', members);  
  ScriptProperties.setProperty('notes', notes);  
  ScriptProperties.setProperty('initialowner', initialowner);  
  //The service url that is required for approval on the email.
  var serviceurl = 'https://script.google.com/a/macros/alaska.edu/s/AKfycbydAnjZUtOkwwqYlL-qIcFXXh9QosAZxNUVGwokTcPjGnMiU52V/exec';
  //Adding spreadsheet ID to the service URL
  //serviceurl+='?spreadsheetId='+spreadsheet.getId();
  var rowNum = getLastPopulatedRow(sheets);  //Getting the row number of the last row that just got populated
  ScriptProperties.setProperty('rowNum', rowNum); 
  serviceurl+='?row='+rowNum; //adding the last populated row number value to the service URL
 var eMailAdd = sheets.getRange(rowNum, 15).getValue();//The username portion of the address suggested for the new group or list
  eMailAdd = eMailAdd+'@alaska.edu';
 ScriptProperties.setProperty('eMailAdd', eMailAdd); 
  //Setting the message that goes on the email sent to the approver.
  var message = '<html><body>There is a new Submission to the UAF List or Group approval Workflow.'+
      '<br /><br /><b>Time of Submission:</b> '+timestamp+' <b>from</b> '+username
      +'<br /><br /><b>Group Full Name:</b> '+fullname+' <b>on</b> '+campus+' <b>campus.</b><br /><br /><b>Group Email Address Requested:</b> '+eMailAdd
      +'<br /><br /><b>Description of the List/Group:</b> '+description + '<br /><br /><b>Is this replacing an existing eMail List or Group:</b> '
      +replacement+'<br /><br /><b>Usage:</b> '+usage+'<br /><br /><b>Anticipated number of members:</b> '+members+'<br /><br /><b>Brief notes about the request :</b>'
      +notes+'<br /><br /><b>Are you the initial owner :</b>'+initialowner+'<br /><br /><a href=\"'+serviceurl+'\"><button>Process Request</button></a></body></html>';
  //Title for the mail sent.
  var title = 'New Group request for '+fullname;
  MailApp.sendEmail('rssimon@alaska.edu, ua-oit-groups-approvers@alaska.edu', title ,"", {htmlBody: message});
}

function processForm(form){
var spreadsheet = SpreadsheetApp.openById('0AtfUVGyOf_82dERDeGdVZHR2bzc3WXlTcDlQaDdUalE');
var asheet = spreadsheet.getSheetByName("Approval Worksheet");
var rowN = UserProperties.getProperty('row');   
var approvalDecisionRange = asheet.getRange(rowN, 18);
var approvalNoteRange = asheet.getRange(rowN, 21);  
var approvalGLRange = asheet.getRange(rowN, 17);  
var approvalUserRange = asheet.getRange(rowN, 19); 
var approvalNameRange = asheet.getRange(rowN, 5);
var approvalEmailRange = asheet.getRange(rowN, 15);  
var approveName = form.nameInput;
var approveEmail = form.emailInput.split('@');
if(form.approvalNotes != "Enter notes on this approval...")  
{
  var approveNote = form.approvalNotes;  
}  
else
{
  var approveNote = '';  
}  
var approveDecision = form.decision;
var approveGLType = form.gltype;
approvalNameRange.setValue(approveName);  
approvalEmailRange.setValue(approveEmail[0]);
approvalNoteRange.setValue(approveNote);
approvalGLRange.setValue(approveGLType);
approvalDecisionRange.setValue(approveDecision);
var activeUserValue = Session.getActiveUser().getUserLoginId();
var activeUserArray = activeUserValue.split('@');  
approvalUserRange.setValue(activeUserArray[0]);
var approvalDateRange = asheet.getRange(rowN, 20);  
var date = new Date();
approvalDateRange.setValue(date);
return 1;
}


function doGet(e)
{
  var spreadsheet = SpreadsheetApp.openById('0AtfUVGyOf_82dERDeGdVZHR2bzc3WXlTcDlQaDdUalE');
  var asheet = spreadsheet.getSheetByName('Approval Worksheet');
  var row = e.parameter['row'];
  UserProperties.setProperty('row', row);
  UserProperties.setProperty('timestamp', asheet.getRange(row, 1).getValue());
  UserProperties.setProperty('username', asheet.getRange(row, 2).getValue());
  Logger.log(asheet.getRange(row, 2).getValue());
  UserProperties.setProperty('campus', asheet.getRange(row, 3).getValue());
  UserProperties.setProperty('teamType', asheet.getRange(row, 4).getValue());
  UserProperties.setProperty('fullGroupName', asheet.getRange(row, 5).getValue());  
  UserProperties.setProperty('emailAddress', asheet.getRange(row, 15).getValue());
  UserProperties.setProperty('description', asheet.getRange(row, 7).getValue()); 
  UserProperties.setProperty('natureOfUse', asheet.getRange(row, 9).getValue());  
  UserProperties.setProperty('addNotes', asheet.getRange(row, 16).getValue());
  UserProperties.setProperty('memEstimate', asheet.getRange(row, 10).getValue());     
  UserProperties.setProperty('initialOwner', asheet.getRange(row, 13).getValue());      
  UserProperties.setProperty('decision', asheet.getRange(row, 18).getValue());      
  UserProperties.setProperty('gltype', asheet.getRange(row, 17).getValue());      
  //MailApp.sendEmail('rssimon@alaska.edu', 'Test Benchmark', 'Test Benchmark for row: '+e.parameter['row'] +'. Checking user property value ' + asheet.getRange(2,10).getValue());//asheet.getRange('C'+row+':C'+row).getValue());  
  return HtmlService.createTemplateFromFile('approvalForm.html').evaluate().setTitle('UAF Google Groups Approval');
 }
