var addPatchEntryInputCells = ["D9","D7","H7","F11","D11","D13","H13","F13","H11","G9"];

// patchid,queueDate,issue,purpose,type,product,priority,owner,component,status,update,accepted,started,signing,completed,notes,released,proactive
var patchLeadDashDataCells = ["D11","I18","D13","I13","F15","I11","F18","D18","D15","I15","G11","F21","H21","J21","I23","D23","G23","D21"];

// patchid,queueDate,issue,purpose,type,product,priority,owner,component,status,update,accepted,started,signing,released,notes
var memberLeadDashDataCells = ["D9","I16","D11","I11","F13","I9","F16","D16","D13","I13","G9","D19","F19","H19","J19","D21"];

var addEntryStatusCell = addPatchEntryInputCells[8];
var addEntryTypeCell = addPatchEntryInputCells[2];
var addEntryAcceptedCell = addPatchEntryInputCells[9];
var addEntryQueueCell = addPatchEntryInputCells[0];
var leadDashSearchPatchIDCell = "D6";
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

//=================ADD ENTRY==========================ADD ENTRY============================ADD ENTRY===============================
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
* Get patch entry button in the lead dashboard.
*/
function GET_PATCH_ENTRY_LEAD(){

  // Get sheet SpreadsheetApp object.
  var spreadSheetApp = SpreadsheetApp;
  
  // Get the current active sheet.
  var currentSheet = spreadSheetApp.getActiveSpreadsheet().getActiveSheet();
  var patchID = currentSheet.getRange(leadDashSearchPatchIDCell).getValue();
  
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