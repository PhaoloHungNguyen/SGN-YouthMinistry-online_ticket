/**
 * Responds to a MESSAGE event in Google Chat.
 *
 * @param {Object} event the event object from Google Chat
 */



function onMessage(event) {
  var sChatText = "ShowGroups"; //event.message.text;
  sChatText = sChatText.trim();
  sChatText = sChatText.toUpperCase();
  sChatText = sChatText.replace(/\s\s+/g, ' ');
  var sCommandLine = ValidateCommand(sChatText);
  
  if (sCommandLine == "Invalid Command.")
    return sCommandLine;
  
  var arrCommand = sCommandLine.split(" ");
  var sCommand = arrCommand[0];
  var sParam = arrCommand[1];

  Logger.log(sCommand);

  switch (sCommand) {
    case "SHOWUSERS":
      var sheetID = "1wdbxkMM7rfbVm0qNr8ssSl1NXmSStmC9-pr4h6c820c";
      var iDoneCol = 5;
      var message = ScanSheetForNotDoneRecord(sheetID, iDoneCol);
      break;
    case "SHOWGROUPS":
      var sheetID = "1N8izRZ5417TPTTZAZY4jrHAsHZ5HbUplw9xUnl8DHbA";
      var iDoneCol = 4;
      var message = ScanSheetForNotDoneRecord(sheetID, iDoneCol);
      break;
    default:
      var message = sCommandLine + " on construction";
      break;
  }
  // for (i = 1; i < lastRow; i++) {
  // //   var cell_CheckBox = sheet1_of_sheet_Active.getRange("E" + i.toString());
  // //   var sCheckBox_Value = cell_CheckBox.getDisplayValue();
  // //   if (sCheckBox_Value.length === 0) {
  // //     var cell_UserName = sheet1_of_sheet_Active.getRange("B" + i.toString());
  // //     var sUserName_Value = cell_UserName.getDisplayValue();
  // //     message = message + sUserName_Value + '\n';

  //   if (table_values[i][4].length === 0)
  //     message = message + table_values[i] + "\n";
  // }
  Logger.log(message); 
 
 
 
  // var name = "";

  // if (event.space.type == "DM") {
  //   name = "You";
  // } else {
  //   name = event.user.displayName;
  // }
  // var message = name + " said \"" + event.message.text + "\"";

  return message;
}

/**
 * Responds to an ADDED_TO_SPACE event in Google Chat.
 *
 * @param {Object} event the event object from Google Chat
 */
function onAddToSpace(event) {
  var message = "";

  if (event.space.singleUserBotDm) {
    message = "Thank you for adding me to a DM, " + event.user.displayName + "!";
  } else {
    message = "Thank you for adding me to " +
        (event.space.displayName ? event.space.displayName : "this chat");
  }

  if (event.message) {
    // Bot added through @mention.
    message = message + " and you said: \"" + event.message.text + "\"";
  }

  return { "text": message };
}

/**
 * Responds to a REMOVED_FROM_SPACE event in Google Chat.
 *
 * @param {Object} event the event object from Google Chat
 */
function onRemoveFromSpace(event) {
  console.info("Bot removed from ",
      (event.space.name ? event.space.name : "this chat"));
}

// CommandLine = Command + Argument
function ValidateCommand(sCommandLine) {
  var arrPart = sCommandLine.split(" ");
  var sCommand = arrPart[0];
  var sArg = arrPart[1];
  var arrCommandWithoutArg = ["HELP", "SHOWUSERS", "SHOWGROUPS", "CREATEGROUPS"];

  if (arrCommandWithoutArg.includes(sCommand))
    return sCommandLine;
  else
    return "Invalid Command.";

}

function ScanSheetForNotDoneRecord(sheetID, iDoneCol) { //iCheckCol: position of Done Column
  var ss = SpreadsheetApp.openById(sheetID);
  SpreadsheetApp.setActiveSpreadsheet(ss);
  var sheet_Active = SpreadsheetApp.getActiveSpreadsheet();
  var sheet1_of_sheet_Active = sheet_Active.getSheets()[0];
  var table_values = sheet1_of_sheet_Active.getDataRange().getDisplayValues()
  var lastRow = sheet1_of_sheet_Active.getLastRow();

  var message = table_values[0] + "\n";
  for (i = 1; i < lastRow; i++) {
    if (table_values[i][iDoneCol-1].length === 0)
      message = message + table_values[i] + "\n";
  }
  return message;
}
