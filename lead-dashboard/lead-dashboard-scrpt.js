var addPatchEntryInputCells = ["D9","D7","H7","F11","D11","D13","H13","F13","H11","G9"];

// patchid,queueDate,issue,purpose,type,product,priority,owner,component,status,update,accepted,started,signing,completed,notes,released,proactive,pmt,staging - 20
var newPatchLeadDashDataCells =["F6","I11","C6","I9","M9","K9","F9","C9","K15","C11","I17","K11","M11","K13","I15","C15","M13","C13","C17","I13"];

// ***DEPRECATED*** // patchid,queueDate,issue,purpose,type,product,priority,owner,component,status,update,accepted,started,signing,completed,notes,released,proactive
var patchLeadDashDataCells = ["D11","I18","D13","I13","F15","I11","F18","D18","D15","I15","G11","F21","H21","J21","I23","D23","G23","D21"];

// ***DEPRECATED***// patchid,queueDate,issue,purpose,type,product,priority,owner,component,status,update,accepted,started,signing,released,notes
var memberLeadDashDataCells = ["D9","I16","D11","I11","F13","I9","F16","D16","D13","I13","G9","D19","F19","H19","J19","D21"];

var addEntryStatusCell = addPatchEntryInputCells[8];
var addEntryTypeCell = addPatchEntryInputCells[2];
var addEntryAcceptedCell = addPatchEntryInputCells[9];
var addEntryQueueCell = addPatchEntryInputCells[0];
var leadDashSearchPatchIDCell = "C3";
var memberDashSearchPatchIDCell = "D5";

var notStarted = "Not Started";
var unassigned = "Unassigned";
var customer = "Customer";
var done = "Done";
var inProgress = "In Progress"
var sentSigning = "Sent for Signing";
var onHold = "On Hold";
var backlog = "Backlog";
var duplicate = "Duplicate";
var rejected = "Rejected";
var prSent = "PR Sent";
var staging = "Staging Tests";
var merged = "PR Merged";
var released = "Released";

var addPatchEntrySheetName = "Add Patch Entry";
var patchDBSheetName = "Patch DataBase"

//================Send EMAIL==========================Send EMAIL============================Send EMAIL=============================
//=================================================================================================================================

/**
 * Send Daily Email.
 *
 * @param {*} objectArray Placeholder map as key value pairs.
 */
function SEND_DAILY_EMAIL(objectArray) {

  var emailTemplate = HtmlService.createTemplateFromFile("dailyEmailTemplate");

  // Get the sheet where the data is in.
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Daily Email Settings");

  var subjectOfTheMail = sheet.getRange("B6").getValue();
  var sendTo = sheet.getRange("B8").getValue();
  var ccList = sheet.getRange("B9").getValue();

  var htmlMsg = emailTemplate.evaluate().getContent();
  var emailBody = replacePlaceHolders(htmlMsg,objectArray);

  //Logger.log(sendTo);
  //Logger.log(ccList);

  // References: https://developers.google.com/apps-script/reference/gmail/gmail-app#sendemailrecipient,-subject,-body,-options
  GmailApp.sendEmail(
    sendTo,
    subjectOfTheMail,
    "Your mailing app does not support HTML. Contact sominda@wso2.com or try with a different app.",
    {
      cc: ccList,
      htmlBody: emailBody
    }
  );
}

/**
 * Send Daily Email.
 *
 * @param {*} objectArray Placeholder map as key value pairs.
 */
function SEND_WELCOME_EMAIL(objectArray) {

  var emailTemplate = HtmlService.createTemplateFromFile("welcomeEmailTemplate");

  // Get the sheet where the data is in.
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Welcome Email Settings");

  var subjectOfTheMail = sheet.getRange("B4").getValue();
  var sendTo = sheet.getRange("B6").getValue();

  var emailBody = emailTemplate.evaluate().getContent();

  //Logger.log(sendTo);

  // References: https://developers.google.com/apps-script/reference/gmail/gmail-app#sendemailrecipient,-subject,-body,-options
  GmailApp.sendEmail(
    sendTo,
    subjectOfTheMail,
    "Your mailing app does not support HTML. Contact sominda@wso2.com or try with a different app.",
    {
      htmlBody: emailBody
    }
  );

  Logger.log("Welcome Email Sent Successfully");
}


//================Daily Record=======================Daily Record==========================Daily Record============================
//=================================================================================================================================

/**
 * DEPRECATED----Updates the daily status of the active patch status.
 *
 * NOTE: This method will be DEPRECATED. Use > DETAILED_DAILY_RECORD.
 */
function DAILY_RECORD() {

    // Get daily recordings sheet.
    var dailyRecordingsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Daily Recordings");

    // Get the daily recordings sheet.
    var lastRow = dailyRecordingsSheet.getLastRow() + 1;

     // Get daily mailing sheet.
    var dailyEmailSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Daily Email");

    //Date,In Progress,Staging Tests,Sent For Signing,Not Started(Customer),Not Started(Internal),Unassigned(Customer),Unassigned(Internal),Team Members
    var dataCells = ["G2","L8","L9","L10","L11","L12","L13","L14","H4"];
    for(var counter = 0; counter < dataCells.length ; counter = counter + 1){
        dailyRecordingsSheet.getRange(lastRow, counter + 1).setValue(dailyEmailSheet.getRange(dataCells[counter]).getValue());
    }
    Logger.log("Daily recording successfully updated");
}

/**
 * Updates the daily status of the active patch status.
 */
function DETAILED_DAILY_RECORD() {

  // Get daily recordings sheet.
  var detailedDailyRecordingsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Daily Detailed Records");

  // Get the daily recordings sheet.
  var lastRow = detailedDailyRecordingsSheet.getLastRow() + 1;

  // --------FOLLOWING ARE THE DATA RANGES------These are the cells in the Daily Email Sheet.
  // Date.
  var dateDataCells = ["G2"];

  // Live fixes: Bugs,Features,Sec,Connector,Total - "Bugs","Features","Sec","Connector","Total"
  var inProgressDataCells = ["H8","I8","J8","K8","L8"];
  var stagingDataCells = ["H9","I9","J9","K9","L9"];
  var signingDataCells = ["H10","I10","J10","K10","L10"];
  var notStartedCustomerDataCells = ["H11","I11","J11","K11","L11"];
  var notStartedInternalDataCells = ["H12","I12","J12","K12","L12"];
  var unassignedCustomerDataCells = ["H13","I13","J13","K13","L13"];
  var unassignedInternalDataCells = ["H14","I14","J14","K14","L14"];
  var teamDataCells = ["H4"];

  // Pending WUM. "Bugs","Features","Sec","Connector"
  var issuesPendingWUMDataCells = ["H19","H20","H21","H22"];
  var updatesPendingWUMDataCells = ["I19","I20","I21","I22"];

  // TOTAL CASES "Bugs","Features","Sec","Connector","Total"
  var issuesTotalDataCells = ["H26","H27","H28","H29","H30"];
  var updatesTotalWUMDataCells = ["I26","I27","I28","I29","I30"];

  // Released WUM. "Bugs","Features","Sec","Connector","Total"
  var issuesReleasedDataCells = ["L26","L27","L28","L29","L30"];
  var updatesReleasedWUMDataCells = ["M26","M27","M28","M29","M30"];

  // Total Total cells
  var totalTotalDataCells = ["H15","I15","J15","K15","L15"];

  // Total Unassigned and not started data cells.
  var totUnassignedAndNotStarted = ["H35","H36","I35","I36","H37","I37"];

  //----------------------------------------------
  var dataCells = [dateDataCells, inProgressDataCells, stagingDataCells, signingDataCells, notStartedCustomerDataCells, notStartedInternalDataCells, unassignedCustomerDataCells,
    unassignedInternalDataCells, teamDataCells, issuesPendingWUMDataCells, updatesPendingWUMDataCells, issuesTotalDataCells, updatesTotalWUMDataCells, issuesReleasedDataCells,
    updatesReleasedWUMDataCells, totalTotalDataCells, totUnassignedAndNotStarted];

  // Data keys.  These are the keys in the email template.
  var objectKeys = [
    ["date"],
    ["inProgressBugs","inProgressFeatures","inProgressSec","inProgressConnector","inProgressTotal"],
    ["stagingBugs","stagingFeatures","stagingSec","stagingConnector","stagingTotal"],
    ["signingBugs","signingFeatures","signingSec","signingConnector","signingTotal"],
    ["notStartedCustBugs","notStartedCustFeatures","notStartedCustSec","notStartedCustConnector","notStartedCustTotal"],
    ["notStartedIntBugs","notStartedIntFeatures","notStartedIntSec","notStartedIntConnector","notStartedIntTotal"],
    ["unassignedCustBugs","unassignedCustFeatures","unassignedCustSec","unassignedCustConnector","unassignedCustTotal"],
    ["unassignedIntBugs","unassignedIntFeatures","unassignedIntSec","unassignedIntConnector","unassignedIntTotal"],
    ["team"],
    ["issuesPendingWUMBugs","issuesPendingWUMFeatures","issuesPendingWUMSec","issuesPendingWUMConnector"],
    ["updatesTotalWUMBugs","updatesTotalWUMFeatures","updatesTotalWUMSec","updatesTotalWUMConnector"],
    ["issuesTotalBugs","issuesTotalFeatures","issuesTotalSec","issuesTotalConnector","issuesTotalTotal"],
    ["updatesTotalBugs","updatesTotalFeatures","updatesTotalSec","updatesTotalConnector","updatesTotalTotal"],
    ["issuesReleasedBugs","issuesReleasedFeatures","issuesReleasedSec","issuesReleasedConnector","issuesReleasedTotal"],
    ["updatesReleasedWUMDBugs","updatesReleasedWUMDFeatures","updatesReleasedWUMDSec","updatesReleasedWUMDConnector","updatesReleasedWUMDTotal"],
    ["totalBugs","totalFeatures","totalSec","totalConnector","totalTotal"],
    ["unassignedCustaTotal","notStartedCustaTotal","unassignedInternalTotal","notStartedInternalTotal","custaTotalUnassignedNotStarted","internalTotalUnassignedNotStarted"],
  ];

  // Count the number of input fields in the sheet.
  var totalInputFields = 0;
  for (var count = 0; count < dataCells.length ; count = count + 1) {
    var dataCellArray = dataCells[count];
    for(var count2 = 0; count2 < dataCellArray.length; count2 = count2 + 1){
      totalInputFields++;
    }
  }

  // Create list for placeholder map.
  var objectArray = new Array(totalInputFields);

  // Get daily mailing sheet.
  var dailyEmailSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Daily Email");
  var columnCounter = 1;
  for (var count = 0; count < dataCells.length ; count = count + 1) {
    var dataCellArray = dataCells[count];
    var keyArray = objectKeys[count];
    for(var count2 = 0; count2 < dataCellArray.length; count2 = count2 + 1){

      // Build the key value pair for email template placeholders.
      var value = dailyEmailSheet.getRange(dataCellArray[count2]).getValue();
      objectArray[columnCounter - 1] = {[keyArray[count2]]:value};

      // Add to sheet.
      detailedDailyRecordingsSheet.getRange(lastRow, columnCounter).setValue(value);
      columnCounter = columnCounter + 1;
    }
  }
  Logger.log("Daily data recorded successfully");
  //var testString = "{{inProgressBugs}},{{inProgressFeatures}},{{inProgressSec}},{{inProgressConnector}},{{inProgressTotal}}";
  //replacePlaceHolders(testString,objectArray);
  SEND_DAILY_EMAIL(objectArray);
  Logger.log("Daily patch summary email sent successfully");
}

//======================ON OPEN=============================ON OPEN=======================================ON OPEN==================
//=================================================================================================================================
/**
* What happen when someone perform a reload
*/
function onOpen() {

  // Get sheet SpreadsheetApp object.
  var spreadSheetApp = SpreadsheetApp;

  // Get the current active sheet.
  var targetSheet = spreadSheetApp.getActiveSpreadsheet().getActiveSheet();
  var activeSheetName = targetSheet.getName();

  // Set the default values ONLY if the sheet is Add patch Entry.
  if(addPatchEntrySheetName == activeSheetName){
    var targetSheet = spreadSheetApp.getActiveSpreadsheet().getSheetByName("Add Patch Entry");
    setDefaultValuesInAddPatchEntry(targetSheet);
  }
}

/**
* Set default values in "Add Patch Entry" sheet.
*/
function setDefaultValuesInAddPatchEntry(targetSheet){

  for(var c = 0 ; c < 10 ; c = c + 1){
      targetSheet.getRange(addPatchEntryInputCells[c]).setValue("");
    }
    targetSheet.getRange(addEntryStatusCell).setValue(unassigned);
    targetSheet.getRange(addEntryTypeCell).setValue(customer);
    var now = new Date();
    targetSheet.getRange(addEntryAcceptedCell).setValue(now);
    targetSheet.getRange(addEntryQueueCell).setValue(now);
}

/**
* Daily set the default values in "Add Patch Entry" sheet.
*/
function setDailyDefaultValues(){

  // Get sheet SpreadsheetApp object.
  var spreadSheetApp = SpreadsheetApp;
  var targetSheet = spreadSheetApp.getActiveSpreadsheet().getSheetByName("Add Patch Entry");
  setDefaultValuesInAddPatchEntry(targetSheet);
}
//=================ADD ENTRY==========================ADD ENTRY============================ADD ENTRY===============================
//=================================================================================================================================

/**
* Add the issue entry to the given sheet.
* @customFunction
*/
function ADDENTRY() {

  // Get sheet SpreadsheetApp object.
  var spreadSheetApp = SpreadsheetApp;

  // Get the current working sheet where the user input the value.
  var activeSheet = spreadSheetApp.getActiveSpreadsheet().getActiveSheet();
  var activeSheetName = activeSheet.getName();

  if(addPatchEntrySheetName == activeSheetName){
    var ui = spreadSheetApp.getUi();

    // Input column cells.
    var cells = addPatchEntryInputCells;

    // Get input values.
    var queueDate = activeSheet.getRange(cells[0]).getValue();
    var issue = activeSheet.getRange(cells[1]).getValue();
    var purpose = activeSheet.getRange(cells[2]).getValue();
    var type = activeSheet.getRange(cells[3]).getValue();
    var product = activeSheet.getRange(cells[4]).getValue();
    var priority = activeSheet.getRange(cells[5]).getValue();
    var owner = activeSheet.getRange(cells[6]).getValue();
    var component = activeSheet.getRange(cells[7]).getValue();
    var status = activeSheet.getRange(cells[8]).getValue();
    var acceptedDate = activeSheet.getRange(cells[9]).getValue();
    if(validateAddEntryInputs(spreadSheetApp,queueDate,issue,purpose,type,product,status,acceptedDate,priority)){

      if(notStarted == status && isEmpty(owner)){
        ui.alert("Please assign a OWNER since the patch status is : " + status);
        return false;
      }

      // Get the target sheet where the values should be stored.
      var targetSheet = spreadSheetApp.getActiveSpreadsheet().getSheetByName(patchDBSheetName);

      // Get the last row of the issue entries.
      var lastRow = targetSheet.getLastRow();
      lastRow = lastRow + 1;
      //Logger.log("ADDENTRY: Patch Entry row > " + lastRow);

      var patchID = lastRow-1;
      //Logger.log("ADDENTRY: Patch ID > " + patchID);

      // Create the input data in order to match the columns order.// ADD an EMPTY STRING FOR PATCH NUMBER COLUMN.
      var valuesList = [patchID,queueDate,issue,purpose,type,product,priority,owner,component,status,"",acceptedDate];

      // Get user conset to add the issue.
      var getConsentFromUI = ui.alert("Do you want to add the entry to the issues list?", ui.ButtonSet.YES_NO);

      // Update the DB.
      if(getConsentFromUI == ui.Button.YES){
        for(var column = 1; column < 13;column = column + 1){
          targetSheet.getRange(lastRow,column ).setValue(valuesList[column-1]);
        }
        ui.alert("Entry successfully added to the database with ID : "+patchID);
        setDefaultValuesInAddPatchEntry(activeSheet);
      }
    } else {
      Logger.log("ADDENTRY: Invalid entry");
    }
  }
}

/**
* Validate Issue, Product version, Status and the Purpose.
*/
function validateAddEntryInputs(spreadSheetApp,queueDate,issue,purpose,type,product,status,acceptedDate,priority){

  var ui = spreadSheetApp.getUi();
  if(isEmpty(queueDate)){
    var getConsentFromUI = ui.alert("Add the date that the patch added to the patch queue");
    return false;
  }
  if(isEmpty(acceptedDate)){
    var getConsentFromUI = ui.alert("Add the date that the patch is accepted");
    return false;
  }
  return validateCommons(ui,issue,purpose,type,product,status,priority);
}

//============MEMBER DASHBOARD====================MEMBER DASHBOARD==================MEMBER DASHBOARD===============================
//=================================================================================================================================
/**
* Get patch entry button in the member dashboard.
*/
function GET_PATCH_ENTRY_MEMBER(){

  var spreadSheetApp = SpreadsheetApp;

  // Get the current active sheet.
  var currentSheet = spreadSheetApp.getActiveSpreadsheet().getActiveSheet();
  var patchID = currentSheet.getRange(memberDashSearchPatchIDCell).getValue();

  if(!isEmpty(patchID)){
    // Logger.log("GET_PATCH_ENTRY: Requested patch ID : " + patchID);
    var rowId = patchID + 1;
    Logger.log("GET_PATCH_ENTRY: Requested patch entry row : " + rowId);

    // DB connection.
    var patchDBSheet = spreadSheetApp.getActiveSpreadsheet().getSheetByName(patchDBSheetName);
    var dataArray = new Array(16);

    // patchid,queue date,issue,purpose,type,product,priority,owner,component,status,update,accepted,started,signing,completed,notes,relased,proactive.
    for(var column = 1; column < 18 ; column = column + 1){
    var value = patchDBSheet.getRange(rowId, column).getValue();
    dataArray[column - 1] = value;
    }
    // patchid,queue date,issue,purpose,type,product,priority,owner,component,status,update,accepted,started,signing,completed,notes
    var cells = memberLeadDashDataCells;

    // Empty the cells.
    for(var counter = 0; counter < 16 ; counter = counter + 1){
    currentSheet.getRange(cells[counter]).setValue("");
    }
    // Set values.
    // Order: patchid,queue date,issue,purpose,type,product,priority,owner,component,status,update,accepted,started,signing,released,notes
    currentSheet.getRange(cells[0]).setValue(dataArray[0]);
    currentSheet.getRange(cells[1]).setValue(dataArray[1]);
    currentSheet.getRange(cells[2]).setValue(dataArray[2]);
    currentSheet.getRange(cells[3]).setValue(dataArray[3]);
    currentSheet.getRange(cells[4]).setValue(dataArray[4]);
    currentSheet.getRange(cells[5]).setValue(dataArray[5]);
    currentSheet.getRange(cells[6]).setValue(dataArray[6]);
    currentSheet.getRange(cells[7]).setValue(dataArray[7]);
    currentSheet.getRange(cells[8]).setValue(dataArray[8]);
    currentSheet.getRange(cells[9]).setValue(dataArray[9]);
    currentSheet.getRange(cells[10]).setValue(dataArray[10]);
    currentSheet.getRange(cells[11]).setValue(dataArray[11]);
    currentSheet.getRange(cells[12]).setValue(dataArray[12]);
    currentSheet.getRange(cells[13]).setValue(dataArray[13]);
    currentSheet.getRange(cells[14]).setValue(dataArray[16]);
    currentSheet.getRange(cells[15]).setValue(dataArray[15]);
  } else {
    Logger.log("GET_PATCH_ENTRY: Empty patch ID");
    var ui = spreadSheetApp.getUi();
    ui.alert("Enter a patch ID");
  }
}

/**
* Update a patch entry from the patch Member dashboard.
*/
function UPDATE_ENTRY_MEMBER(){

    // patchid,queueDate,issue,purpose,type,product,priority,owner,component,status,update,accepted,started,signing,released,notes
  var cells = memberLeadDashDataCells;
  var spreadSheetApp = SpreadsheetApp;
  var currentSheet = spreadSheetApp.getActiveSpreadsheet().getActiveSheet();
  var dataArray = new Array(16);
  for(var counter = 0 ; counter < 16 ; counter = counter + 1){
    dataArray[counter] = currentSheet.getRange(cells[counter]).getValue();
  }
  var rowID = dataArray[0]+1;
  var patchID = dataArray[0];
  var updateNumber = dataArray[10];

  var ui = spreadSheetApp.getUi();
  var patchStatus = dataArray[9];
  var today = new Date();

  // Set in progress date.
  if(inProgress == patchStatus && isEmpty(dataArray[12])){
    dataArray[12] = today;
  }
  // Set sent for signing date.
  if(sentSigning == patchStatus && isEmpty(dataArray[13])){
    dataArray[13] = today;
  }
  Logger.log("Update Number > "+ updateNumber);
  Logger.log("Patch Status > "+ patchStatus);
  if(isEmpty(updateNumber)){
      if(!(onHold == patchStatus || duplicate == patchStatus || rejected == patchStatus || backlog == patchStatus || inProgress == patchStatus)){
          ui.alert("Enter Update Number")
          return false;
      }
  }
  var component = dataArray[8];
  if(isEmpty(component)){
    if( prSent == patchStatus || merged == patchStatus || staging == patchStatus || sentSigning == patchStatus){
      ui.alert("Select a component from the list");
      return false;
    }
  }
  var completeEntry = true;
  if(released == patchStatus){
      completeEntry = validateCompletedMemberEntryEntry(ui,dataArray);
      if(isEmpty(dataArray[14])){
        dataArray[14] = new Date();
    }
  }
  if(completeEntry && validateCommons(ui,dataArray[2],dataArray[3],dataArray[4],dataArray[5],patchStatus,dataArray[6])){

    // Get consent.
    var getConsentFromUI = ui.alert("Do you want update the Entry: " + patchID + "?", ui.ButtonSet.YES_NO);
    if(getConsentFromUI == ui.Button.YES){
        // Component, Status, Update Number, Started Date, Sent for signing, Notes, Released.
        var columns = [9,10,11,13,14,16,17];
        var patchDBSheet = spreadSheetApp.getActiveSpreadsheet().getSheetByName(patchDBSheetName);
        patchDBSheet.getRange(rowID,columns[0]).setValue(dataArray[8]);
        patchDBSheet.getRange(rowID,columns[1]).setValue(patchStatus);
        patchDBSheet.getRange(rowID,columns[2]).setValue(updateNumber);
        patchDBSheet.getRange(rowID,columns[3]).setValue(dataArray[12]);
        patchDBSheet.getRange(rowID,columns[4]).setValue(dataArray[13]);
        patchDBSheet.getRange(rowID,columns[5]).setValue(dataArray[15]);
        patchDBSheet.getRange(rowID,columns[6]).setValue(dataArray[14]);
        for(var counter = 0; counter < 16 ; counter = counter + 1){
            currentSheet.getRange(cells[counter]).setValue("");
        }
        ui.alert("Patch : "+ patchID +" successfully updated to the DB");
    }
  }
}

/**
 * Validate a complete patch entry.
 *
 * @param {*} ui UI object
 * @param {*} dataArray Array with data from the UI.
 */
function validateCompletedMemberEntryEntry(ui,dataArray){

    var scenarios = ["Patch Id", "Patch Queue Date","Issue","Purpose","Patch Type","Product Version", "Patch Priority","Owner","Component","Status","Update Number","Patch Entry Accepted Date","Development Started Date","Patch Sent for Signing Date"];
    for(var counter  = 0 ; counter < scenarios.length ; counter = counter + 1){
        if(isEmpty(dataArray[counter])){
            ui.alert("Incomplete patch entry. Missing : " + scenarios[counter]);
            return false;
        }
    }
    return true;
}

//==================LEAD DASHBOARD======================LEAD DASHBOARD===========================LEAD DASHBOARD====================
//=================================================================================================================================

/**
* **DEPRECATED***
* Get patch entry button in the lead dashboard.
*/
function GET_PATCH_ENTRY_LEAD(){

  // Get sheet SpreadsheetApp object.
  var spreadSheetApp = SpreadsheetApp;

  // Get the current active sheet.
  var currentSheet = spreadSheetApp.getActiveSpreadsheet().getActiveSheet();
  var patchID = currentSheet.getRange("D6").getValue();

  if(!isEmpty(patchID)){
    // Logger.log("GET_PATCH_ENTRY: Requested patch ID : " + patchID);
    var rowId = patchID + 1;
    Logger.log("GET_PATCH_ENTRY: Requested patch entry row : " + rowId);

    // DB connection.
    var patchDBSheet = spreadSheetApp.getActiveSpreadsheet().getSheetByName(patchDBSheetName);
    var dataArray = new Array(16);

    // order ::: patchid,queueDate,issue,purpose,type,product,priority,owner,component,status,update,accepted,started,signing,completed,notes,released,proactive.
    for(var column = 1; column < 19 ; column = column + 1){
    var value = patchDBSheet.getRange(rowId, column).getValue();
    dataArray[column - 1] = value;
    }
    // patchid,queue date,issue,purpose,type,product,priority,owner,component,status,update,accepted,started,signing,completed,notes,released,proactive.
    var cells = patchLeadDashDataCells;
    for(var counter = 0; counter < 16 ; counter = counter + 1){
        currentSheet.getRange(cells[counter]).setValue("");
    }
    for(var counter = 0; counter < 18 ; counter = counter + 1){
        currentSheet.getRange(cells[counter]).setValue(dataArray[counter]);
    }
  } else {
    Logger.log("GET_PATCH_ENTRY: Empty patch ID");
    var ui = spreadSheetApp.getUi();
    ui.alert("Enter a patch ID");
  }
}

/**
* Get patch entry button in the lead dashboard.
*/
function GET_PATCH_ENTRY_BY_LEAD1(){

  // Get sheet SpreadsheetApp object.
  var spreadSheetApp = SpreadsheetApp;

  // Get the current active sheet.
  var currentSheet = spreadSheetApp.getActiveSpreadsheet().getActiveSheet();
  var patchID = currentSheet.getRange(leadDashSearchPatchIDCell).getValue();

  if(!isEmpty(patchID)){
    var rowId = patchID + 1;
    Logger.log("GET_PATCH_ENTRY_BY_LEAD1: Requested patch entry row : " + rowId);

    // DB connection.
    var patchDBSheet = spreadSheetApp.getActiveSpreadsheet().getSheetByName(patchDBSheetName);
    var dataArray = new Array(newPatchLeadDashDataCells.length);

    // order ::: patchid,queueDate,issue,purpose,type,product,priority,owner,component,status,update,accepted,started,signing,completed,notes,released,proactive.
    for(var column = 1; column < newPatchLeadDashDataCells.length + 1 ; column = column + 1){
      var value = patchDBSheet.getRange(rowId, column).getValue();
      dataArray[column - 1] = value;
    }
    // patchid,queue date,issue,purpose,type,product,priority,owner,component,status,update,accepted,started,signing,completed,notes,released,proactive.
    var cells = newPatchLeadDashDataCells;
    for(var counter = 0; counter < newPatchLeadDashDataCells.length ; counter = counter + 1){
        currentSheet.getRange(cells[counter]).setValue("");
    }
    for(var counter = 0; counter < newPatchLeadDashDataCells.length ; counter = counter + 1){
        currentSheet.getRange(cells[counter]).setValue(dataArray[counter]);
    }
  } else {
    Logger.log("GET_PATCH_ENTRY: Empty patch ID");
    var ui = spreadSheetApp.getUi();
    ui.alert("Enter a patch ID");
  }
}


/**
* Update a patch entry from the patch lead dashboard.
*/
function UPDATE_ENTRY_LEAD_NEW(){

  var spreadSheetApp = SpreadsheetApp;
  var ui = spreadSheetApp.getUi();
  var currentSheet = spreadSheetApp.getActiveSpreadsheet().getActiveSheet();
  var patchDBSheet = spreadSheetApp.getActiveSpreadsheet().getSheetByName(patchDBSheetName);

  // patchid,queueDate,issue,purpose,type,product,priority,owner,component,status,update,accepted,started,signing,completed,notes,released,proactive,pmt,staging
  var dataArray = new Array(newPatchLeadDashDataCells.length);
  for(var counter = 0 ; counter < newPatchLeadDashDataCells.length ; counter++){
    dataArray[counter] = currentSheet.getRange(newPatchLeadDashDataCells[counter]).getValue();
  }

  var patchID = dataArray[0];
  var patchStatus = dataArray[9];
  var owner = dataArray[7];

  // Validate patch owner with the status.
  if((notStarted == patchStatus || done == patchStatus || released == patchStatus) && isEmpty(owner)){
    ui.alert("Assign a OWNER since the patch status is : " + patchStatus);
    return false;
  }

  if(unassigned == patchStatus && !isEmpty(owner)){
    ui.alert("Remove the OWNER since the patch status is : " + patchStatus);
    return false;
  }

  // Validate complete entry.
  var completeEntry = true;
  if(done == patchStatus){
    completeEntry = validateCompletedLeadPatchEntry(ui,dataArray);
    if(isEmpty(dataArray[14])){
        dataArray[14] = new Date();
    }
  }
  if(validateCommons(ui,dataArray[2],dataArray[3],dataArray[4],dataArray[5],patchStatus,dataArray[6]) && completeEntry){

    // Row of the DB to be updated.
    var rowID = patchID + 1;
    // Get consent.
    var getConsentFromUI = ui.alert("Do you want update the Entry: " + patchID + "?", ui.ButtonSet.YES_NO);
    if(getConsentFromUI == ui.Button.YES){
      // Priority,Owner,Status,Proactive,Notes,PMT.
      // Columns: G,H,J,R,P,S.
      var columns = [7,8,10,15,18,16,19];
      patchDBSheet.getRange(rowID, columns[0]).setValue(dataArray[6]);
      patchDBSheet.getRange(rowID, columns[1]).setValue(dataArray[7]);
      patchDBSheet.getRange(rowID, columns[2]).setValue(dataArray[9]);
      patchDBSheet.getRange(rowID, columns[3]).setValue(dataArray[14]);
      patchDBSheet.getRange(rowID, columns[4]).setValue(dataArray[17]);
      patchDBSheet.getRange(rowID, columns[5]).setValue(dataArray[15]);
      patchDBSheet.getRange(rowID, columns[6]).setValue(dataArray[18]);

      for(var counter = 0; counter < newPatchLeadDashDataCells.length ; counter++){
        currentSheet.getRange(newPatchLeadDashDataCells[counter]).setValue("");
      }
      ui.alert("Patch : "+ patchID +" successfully updated to the DB");
    }
  }
}

/**
* Update a patch entry from the patch lead dashboard.
*/
function UPDATE_ENTRY_LEAD(){

  // patchid,queueDate,issue,purpose,type,product,priority,owner,component,status,update,accepted,started,signing,completed,notes
  var cells = patchLeadDashDataCells;
  var spreadSheetApp = SpreadsheetApp;
  var currentSheet = spreadSheetApp.getActiveSpreadsheet().getActiveSheet();
  var dataArray = new Array(16);
  for(var counter = 0 ; counter < 18 ; counter = counter + 1){
    dataArray[counter] = currentSheet.getRange(cells[counter]).getValue();
  }
  var patchDBSheet = spreadSheetApp.getActiveSpreadsheet().getSheetByName(patchDBSheetName);
  var rowID = dataArray[0]+1;
  var patchID = dataArray[0];

  var ui = spreadSheetApp.getUi();

  // Validate patch owner with the status
  var patchStatus = dataArray[9];
  var owner = dataArray[7];
  if((notStarted == patchStatus || done == patchStatus) && isEmpty(owner)){
    ui.alert("Assign a OWNER since the patch status is : " + patchStatus);
    return false;
  }
  if(unassigned == patchStatus && !isEmpty(owner)){
    ui.alert("Remove the OWNER since the patch status is : " + patchStatus);
    return false;
  }
  // Validate complete entry
  var completeEntry = true;
  if(done == patchStatus){
    completeEntry = validateCompletedLeadEntry(ui,dataArray);
    if(isEmpty(dataArray[14])){
        dataArray[14] = new Date();
    }
  }
  if(validateCommons(ui,dataArray[2],dataArray[3],dataArray[4],dataArray[5],patchStatus,dataArray[6]) && completeEntry){

    // Get consent.
    var getConsentFromUI = ui.alert("Do you want update the Entry: " + patchID + "?", ui.ButtonSet.YES_NO);
    if(getConsentFromUI == ui.Button.YES){
      // Purpose,Type,Product,Priority,owner,component,status,completed,proactive
      var columns = [4,5,6,7,8,9,10,15,16,18];
      patchDBSheet.getRange(rowID, columns[0]).setValue(dataArray[3]);
      patchDBSheet.getRange(rowID, columns[1]).setValue(dataArray[4]);
      patchDBSheet.getRange(rowID, columns[2]).setValue(dataArray[5]);
      patchDBSheet.getRange(rowID, columns[3]).setValue(dataArray[6]);
      patchDBSheet.getRange(rowID, columns[4]).setValue(dataArray[7]);
      patchDBSheet.getRange(rowID, columns[5]).setValue(dataArray[8]);
      patchDBSheet.getRange(rowID, columns[6]).setValue(dataArray[9]);
      patchDBSheet.getRange(rowID, columns[7]).setValue(dataArray[14]);
      patchDBSheet.getRange(rowID, columns[8]).setValue(dataArray[15]);
      patchDBSheet.getRange(rowID, columns[9]).setValue(dataArray[17]);

      for(var counter = 0; counter < 18 ; counter = counter + 1){
        currentSheet.getRange(cells[counter]).setValue("");
      }
      ui.alert("Patch : "+ patchID +" successfully updated to the DB");
    }
  }
}

/**
 * Validate a complete patch entry.
 *
 * @param {*} ui UI object
 * @param {*} dataArray Array with data from the UI.
 */
function validateCompletedLeadPatchEntry(ui,dataArray){

  // patchid,queueDate,issue,purpose,type,product,priority,owner,component,status,update,accepted,started,signing,completed,notes,released,proactive,pmt,staging.
  var scenarios = ["Patch Id", "Patch Queue Date","Issue","Purpose","Patch Type","Product Version", "Patch Priority","Owner","Component","Status","Update Number",
  "Patch Entry Accepted Date","Development Started Date","Patch Sent for Signing Date","Patch Completed Date","Empty Notes","Patch Released Date","Proactive Status",
  "PMT link","Staging date"];
  for(var counter  = 0 ; counter < scenarios.length ; counter = counter + 1){

    // 14 - completed,15 - notes
    if(!(counter == 14 || counter == 15)){
        if(isEmpty(dataArray[counter])){
            ui.alert("Incomplete patch entry. Missing : " + scenarios[counter]);
            return false;
        }
    }
  }
  return true;
}

/**
 * DEPRECATED.
 * Validate a complete patch entry.
 *
 * @param {*} ui UI object
 * @param {*} dataArray Array with data from the UI.
 */
function validateCompletedLeadEntry(ui,dataArray){

    // Order:: patchid,queueDate,issue,purpose,type,product,priority,owner,component,status,update,accepted,started,signing,completed,notes,released,proactive.
    var scenarios = ["Patch Id", "Patch Queue Date","Issue","Purpose","Patch Type","Product Version", "Patch Priority","Owner","Component","Status","Update Number","Patch Entry Accepted Date","Development Started Date","Patch Sent for Signing Date","Update Completed Date","Empty Notes","Update Released Date","Update Proactive Status"];
    for(var counter  = 0 ; counter < scenarios.length ; counter = counter + 1){
        if(!(counter == 14 || counter == 15)){
            if(isEmpty(dataArray[counter])){
                ui.alert("Incomplete patch entry. Missing : " + scenarios[counter]);
                return false;
            }
        }
    }
    return true;
}

//===============Commons================================Commons=======================Commons======================================
//=================================================================================================================================

/**
 * Check for an empty string value.
 *
 * @param {*} str Any string value.
 */
function isEmpty(str) {
  return (!str || 0 == str.length);
}

/**
 * Validate common attributes of a patch entry.
 *
 * @param {*} ui Ui object
 * @param {*} issue Issue
 * @param {*} purpose Purpose
 * @param {*} type Type
 * @param {*} product Product version
 * @param {*} status Status
 */
function validateCommons(ui,issue,purpose,type,product,status,priority){

  //Logger.log("validateCommons: issue > " + issue);
  //Logger.log("validateCommons: purpose > " + purpose);
  //Logger.log("validateCommons: type > " + type);
  //Logger.log("validateCommons: product > " + product);
  //Logger.log("validateCommons: status > " + status);

  if(isEmpty(issue)){
    ui.alert("Add an Issue");
    return false;
  }
  if(isEmpty(purpose)){
    ui.alert("Select the Purpose of the issue");
    return false;
  }
  if(isEmpty(type)){
    ui.alert("Select the type of the issue");
    return false;
  }
  if(isEmpty(product)){
    ui.alert("Select a product from the list");
    return false;
  }
  if(isEmpty(status)){
    ui.alert("Select the status of the issue");
    return false;
  }
  if(isEmpty(priority)){
    ui.alert("Select the priority level for the issue");
    return false;
  }
  return true;
}

/**
 * Replace the placehoders in a string in the format of {{placeHolder}}.
 *
 * @param {String} template Email template in the String format.
 * @param {Array} objectArray PlaceHolders list. Eg: [{placeHolder1:value1},{placeHolder2:value2}]
 */
function replacePlaceHolders(template,objectArray){

  for (let index = 0; index < objectArray.length; index++) {
    var placeHolderObject = objectArray[index];
    for (let [key, value] of Object.entries(placeHolderObject)) {
      var placeholderkey = "{{"+ key +"}}";
      var placeholderValue = ""+ value +"";
      // Logger.log('placeholdes:'+placeholderkey+' : '+placeholderValue);
      template = template.replace(placeholderkey, placeholderValue);
    }
  }
  //Logger.log(template);
  return template;
}
