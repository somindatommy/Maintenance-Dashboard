// ====================== Member dashboard Cells ======================

// patchid,queueDate,issue,purpose,type,product,priority,owner,component,status,update,accepted,started,signing,completed,notes.released,proactive,PMT,Staging Date// length =20
var memberDashDataCells = ["C10","C29","E10","C27","G25","E25","E27","C25","C13","G13","C15","E29","G29","C31","G31","C19","E31","G27","C17","C33"];

var numberOfInputFeilds = memberDashDataCells.length;

// Cell location of the search patch ID.
var memberDashSearchPatchIDCell = "C5";
var memberDashSearchByIssueCell = "H5";

// ====================== Patch Entry Status ======================
var notStarted = "Not Started";
var unassigned = "Unassigned";
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

// ====================== Maintenance Dashboard Variables ======================
var patchDBSheetName = "Patch DataBase"
var maintenanceDbURL = "https://docs.google.com/spreadsheets/d/1TBlwwmkstigSHZRe0y65WnvIYvaq4jted8j-BOld7FM/edit"

// ====================== COMPONENTS ======================
var CONNECTOR_COMPONENT = "Connector";

// ====================== OTHER CONSTANTS ======================
var UPDATE_NUMBER_PREFIX = "WSO2-CARBON-PATCH-";
var COONECTER_UPDATE_NUMBER_PREFIX = UPDATE_NUMBER_PREFIX + "0.0.0-";

//=================================================================================================================================
//=================================================================================================================================
//=================================================================================================================================

function GET_ENRTY_BY_ISSUE() {

  var spreadSheetApp = SpreadsheetApp;

  // Get the current active sheet.
  var currentSheet = spreadSheetApp.getActiveSpreadsheet().getActiveSheet();
  var issue = currentSheet.getRange(memberDashSearchByIssueCell).getValue();

  if (!isEmpty(issue)) {
    issue = issue.trim();

    // Get maintenance Dashboard connection.
    var maintenanceDashboard = spreadSheetApp.openByUrl(maintenanceDbURL);
    var patchDBSheet = maintenanceDashboard.getSheetByName(patchDBSheetName);

    // Get the row with last entry.
    var lastRow = patchDBSheet.getLastRow();
    var range = "Patch DataBase!C1:C" + lastRow;
    var issuesList = patchDBSheet.getRange(range).getValues();

    if (issuesList.length < 1) {
      var ui = spreadSheetApp.getUi();
      ui.alert("Empty patch database.");
    }

    // Get the row Id of the matching issue;
    var rowId = 0;
    for (counter = 0; counter < issuesList.length; counter++) {
      if (containsText(issuesList[counter].toString(), issue.toString())) {
        rowId = counter + 1;
        break;
      }
    }
    if (rowId < 1) {
      var ui = spreadSheetApp.getUi();
      ui.alert("No entries with the given issue.");
      return false;
    }
    // patchId,queueDate,issue,purpose,type,product,priority,owner,component,status,update,accepted,started,signing,
    completed,notes.released,proactive,PMT.
    var dataArray = new getEntryByRowID(rowId, patchDBSheet);
    populateDashBoard(currentSheet, dataArray);
  } else {
    Logger.log("GET_PATCH_ENTRY: Empty value as issue");
    var ui = spreadSheetApp.getUi();
    ui.alert("Enter an Issue ");
  }
}

/**
* Get patch entry button in the member dashboard.
*/
function GET_PATCH_ENTRY_MEMBER() {

  var spreadSheetApp = SpreadsheetApp;

  // Get the current active sheet.
  var currentSheet = spreadSheetApp.getActiveSpreadsheet().getActiveSheet();
  var patchID = currentSheet.getRange(memberDashSearchPatchIDCell).getValue();

  if (!isEmpty(patchID)) {
    var rowId = patchID + 1;
    Logger.log("GET_PATCH_ENTRY: Requested patch entry row : " + rowId);

    // Get maintenance Dashboard connection.
    var maintenanceDashboard = spreadSheetApp.openByUrl(maintenanceDbURL);
    var patchDBSheet = maintenanceDashboard.getSheetByName(patchDBSheetName);

   // patchId,queueDate,issue,purpose,type,product,priority,owner,component,status,update,accepted,started,signing,
    completed,notes.released,proactive,PMT.
    var dataArray = new getEntryByRowID(rowId, patchDBSheet);
    populateDashBoard(currentSheet, dataArray);
  } else {
    Logger.log("GET_PATCH_ENTRY: Empty patch ID");
    var ui = spreadSheetApp.getUi();
    ui.alert("Enter a patch ID");
  }
}

/**
 * Get a patch entry given by the row Id.
 *
 * @param {int} rowId Row Id
 */
function getEntryByRowID(rowId, patchDBSheet) {

  var dataArray = new Array(numberOfInputFeilds);

  // patchd,queueDate,issue,purpose,type,product,priority,owner,component,status,update,accepted,started,signing,completed,notes.released,proactive,PMT,staging.
  for (var column = 1; column < numberOfInputFeilds + 1; column = column + 1) {
    var value = patchDBSheet.getRange(rowId, column).getValue();
    dataArray[column - 1] = value;
  }
  return dataArray;
}

/**
 * Populate the dashboard with the values from the DB.
 * @param {SpeadSheet} currentSheet
 * @param {string[]} dataArray
 */
function populateDashBoard(currentSheet, dataArray) {

  var cells = memberDashDataCells;

  for (var counter = 0; counter < numberOfInputFeilds; counter = counter + 1) {
    currentSheet.getRange(cells[counter]).setValue("");
  }
  // Set values.
  // Order: patchid,queueDate,issue,purpose,type,product,priority,owner,component,status,update,accepted,started,signing,completed,notes.released,proactive,PMT,staging.
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
  currentSheet.getRange(cells[14]).setValue(dataArray[14]);
  currentSheet.getRange(cells[15]).setValue(dataArray[15]);
  currentSheet.getRange(cells[16]).setValue(dataArray[16]);
  currentSheet.getRange(cells[17]).setValue(dataArray[17]);
  currentSheet.getRange(cells[18]).setValue(dataArray[18]);
  currentSheet.getRange(cells[19]).setValue(dataArray[19]);
}

/**
* Update a patch entry from the patch Member dashboard.
*/
function UPDATE_ENTRY_MEMBER() {

  // patchid,queueDate,issue,purpose,type,product,priority,owner,component,status,update,accepted,started,signing,completed,notes.released,proactive,PMT.
  var cells = memberDashDataCells;
  var spreadSheetApp = SpreadsheetApp;
  var currentSheet = spreadSheetApp.getActiveSpreadsheet().getActiveSheet();
  var dataArray = new Array(numberOfInputFeilds);
  var analysisArray = new Array(analysisDetails.length);
  var lastUpdateIDNoBeforeAnalysis = 823;


  // Read Values from the cells.
  for (var counter = 0; counter < numberOfInputFeilds; counter = counter + 1) {
    dataArray[counter] = currentSheet.getRange(cells[counter]).getValue();
  }
  var patchID = dataArray[0];
  var updateNumber = dataArray[10];
  var patchStatus = dataArray[9];
  var component = dataArray[8];
  var notes = dataArray[15];

  var ui = spreadSheetApp.getUi();
  var today = new Date();
  var rowID = patchID + 1;
  var analyticRowID = rowID - lastUpdateIDNoBeforeAnalysis;

  // Set in progress date.
  if (inProgress == patchStatus && isEmpty(dataArray[12])) {
    if (isEmpty(dataArray[12])) {
      dataArray[12] = today;
    }
  }
  // Set sent for signing date.
  if (sentSigning == patchStatus && isEmpty(dataArray[13])) {
    if (isEmpty(dataArray[13])) {
      dataArray[13] = today;
    }
  }
  // Set the staging date.
  if (staging == patchStatus) {
    if (isEmpty(dataArray[19])) {
      dataArray[19] = today;
    }
  }
  if (isEmpty(component)) {
    if (prSent == patchStatus || merged == patchStatus || staging == patchStatus || sentSigning == patchStatus || released == patchStatus) {
      ui.alert("Select a component from the list");
      return false;
    }
  } else {
    // Auto generate a UPDATE NUMBER when the fixed component is "Connector".
    if (component == CONNECTOR_COMPONENT) {
      if (isEmpty(updateNumber)) {
        updateNumber = COONECTER_UPDATE_NUMBER_PREFIX + patchID;
        dataArray[10] = updateNumber;
        Logger.log("Update Number : " + updateNumber + " was automatically generated for PatchID : " + patchID);
      }
    }
  }
  if (isEmpty(updateNumber)) {
    if (!(onHold == patchStatus || duplicate == patchStatus || rejected == patchStatus || backlog == patchStatus || inProgress == patchStatus)) {
      ui.alert("Enter Update Number")
      return false;
    }
  }
  var completeEntry = true;
  if (released == patchStatus) {
    completeEntry = validateCompletedMemberEntryEntry(ui, dataArray);
    if (isEmpty(dataArray[16])) {
      dataArray[16] = today;
    }

    if (patchID > lastUpdateIDNoBeforeAnalysis) {
      // Read Values from the analytics cells.
      for (var counter = 0; counter < analysisDetails.length; counter = counter + 1) {
        analysisArray[counter] = currentSheet.getRange(analysisDetails[counter]).getValue();
      }
      completeEntry = validateAnaysisEntry(ui, analysisArray);
    }

  }

  if (completeEntry && validateCommons(ui, dataArray[2], dataArray[3], dataArray[4], dataArray[5], patchStatus, dataArray[6])) {

    // Get consent.
    var getConsentFromUI = ui.alert("Do you want update the Entry: " + patchID + "?", ui.ButtonSet.YES_NO);
    if (getConsentFromUI == ui.Button.YES) {
      // Component, Status, Update Number, Started Date, Sent for signing, Notes, Released, PMT link, Staging date.
      var columns = [9, 10, 11, 13, 14, 16, 17, 19, 20];

      var maintenanceDashboard = spreadSheetApp.openByUrl(maintenanceDbURL);
      var patchDBSheet = maintenanceDashboard.getSheetByName(patchDBSheetName);
      patchDBSheet.getRange(rowID, columns[0]).setValue(component);
      patchDBSheet.getRange(rowID, columns[1]).setValue(patchStatus);
      patchDBSheet.getRange(rowID, columns[2]).setValue(updateNumber);
      patchDBSheet.getRange(rowID, columns[3]).setValue(dataArray[12]);
      patchDBSheet.getRange(rowID, columns[4]).setValue(dataArray[13]);
      patchDBSheet.getRange(rowID, columns[5]).setValue(notes);
      patchDBSheet.getRange(rowID, columns[6]).setValue(dataArray[16]);
      patchDBSheet.getRange(rowID, columns[7]).setValue(dataArray[18]);
      patchDBSheet.getRange(rowID, columns[8]).setValue(dataArray[19]);

      // Post analysis update
      if (released == patchStatus && patchID > lastUpdateIDNoBeforeAnalysis) {
        var postDBSheet = maintenanceDashboard.getSheetByName(postAnalysisDBSheetName);
        var analysisColumns = [1, 2, 3, 4, 5, 6, 7, 8];

        completeEntry = validateAnaysisEntry(ui, analysisArray);
        postDBSheet.getRange(analyticRowID, analysisColumns[0]).setValue(patchID);
        postDBSheet.getRange(analyticRowID, analysisColumns[1]).setValue(today);
        postDBSheet.getRange(analyticRowID, analysisColumns[2]).setValue(updateNumber);
        postDBSheet.getRange(analyticRowID, analysisColumns[3]).setValue(analysisArray[0]);
        postDBSheet.getRange(analyticRowID, analysisColumns[4]).setValue(analysisArray[1]);
        postDBSheet.getRange(analyticRowID, analysisColumns[5]).setValue(component);
        postDBSheet.getRange(analyticRowID, analysisColumns[6]).setValue(analysisArray[2]);
        postDBSheet.getRange(analyticRowID, analysisColumns[7]).setValue(dataArray[4]);
        ui.alert("Post analysis for Patch : " + patchID + "  successfully updated to the DB");

      }

      // Clean the cells
      for (var counter = 0; counter < numberOfInputFeilds; counter = counter + 1) {
        currentSheet.getRange(cells[counter]).setValue("");
      }
      for (var counter = 0; counter < analysisDetails.length; counter = counter + 1) {
        currentSheet.getRange(analysisDetails[counter]).setValue("");
      }
      ui.alert("Patch : " + patchID + " successfully updated to the DB");
    }
  }
}

/**
 * Validate a complete patch entry.
 *
 * @param {UI object} ui UI object
 * @param {string[]} dataArray Array with data from the UI.
 */
function validateCompletedMemberEntryEntry(ui, dataArray) {

  var completedDate = "Completed Date";
  var notes = "Notes";
  var releasedDate = "Released Date";
  var proactiveStatus = "Proactive Status";
  var pmtLink = "PMT link";
  var scenarios = ["Patch Id", "Patch Queue Date", "Issue", "Purpose", "Patch Type", "Product Version", "Patch Priority", "Owner", "Component", "Status", "Update Number", "Patch Entry Accepted Date",
    "Development Started Date", "Patch Sent for Signing Date", completedDate, notes, releasedDate, proactiveStatus, pmtLink, "Staging Date"];
  for (var counter = 0; counter < scenarios.length; counter = counter + 1) {
    var scenario = scenarios[counter];
    if (scenario == completedDate || scenario == notes || scenario == releasedDate ||
      scenario == proactiveStatus) {
      continue;
    }
    if (isEmpty(dataArray[counter])) {
      ui.alert("Incomplete patch entry. Missing : " + scenario);
      return false;
    }
  }
  return true;
}

/**
 * Validate a complete analysis entry.
 *
 * @param {UI object} ui UI object
 * @param {string[]} dataArray Array with data from the UI.
 */
function validateAnaysisEntry(ui, dataArray) {

  var scenarios = ["Issue Summary", "Solution Summary", "Related Team"];
  for (var counter = 0; counter < scenarios.length; counter = counter + 1) {
    var scenario = scenarios[counter];
    if (isEmpty(dataArray[counter])) {
      ui.alert("Incomplete patch analysis entry. Missing : " + scenario);
      return false;
    }
  }
  return true;
}

//===============Commons================================Commons=======================Commons======================================
//=================================================================================================================================

/**
 * Checks whether the given string contains the given pattern.
 * Returns TRUE if success.
 *
 * @param {string} text Text to be searched
 * @param {string} pattern Pattern to be searched
 */
function containsText(text, pattern) {

  var lengthText = text.length;
  var lengthPattern = pattern.length;
  if (lengthText < lengthPattern) {
    return false;
  }
  if (lengthPattern == lengthText) {
    if (text == pattern) {
      return true;
    }
    return false;
  }
  for (var i = 0; i < lengthText - lengthPattern + 1; i++) {
    if (text.substring(i, lengthPattern + i) == pattern) {
      return true;
    }
  }
  return false;
}

/**
 * Check for an empty string value.
 *
 * @param {string} str Any string value.
 */
function isEmpty(str) {
  return (!str || 0 == str.length);
}

/**
 * Validate common attributes of a patch entry.
 *
 * @param {string} ui Ui object
 * @param {string} issue Issue
 * @param {string} purpose Purpose
 * @param {string} type Type
 * @param {string} product Product version
 * @param {string} status Status
 */
function validateCommons(ui, issue, purpose, type, product, status, priority) {

  if (isEmpty(issue)) {
    ui.alert("Add an Issue");
    return false;
  }
  if (isEmpty(purpose)) {
    ui.alert("Select the Purpose of the issue");
    return false;
  }
  if (isEmpty(type)) {
    ui.alert("Select the type of the issue");
    return false;
  }
  if (isEmpty(product)) {
    ui.alert("Select a product from the list");
    return false;
  }
  if (isEmpty(status)) {
    ui.alert("Select the status of the issue");
    return false;
  }
  if (isEmpty(priority)) {
    ui.alert("Select the priority level for the issue");
    return false;
  }
  return true;
}