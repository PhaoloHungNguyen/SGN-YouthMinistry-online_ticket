// Interaction with Google Sheet
function onAddToSpace(event) {
  var message = "";

  if (event.space.singleUserBotDm) {
    message = "Thank you for adding me to a DM, " + event.user.displayName + "!";
  } else {
    message = "Thank you for adding me to " +
        (event.space.displayName ? event.space.displayName : "this chat");
  }

  var ss = SpreadsheetApp.openById("1UDpX46mnnzPDGRA4t14f3AAroFCICv31pi4I3Sklz1Q");
  Logger.log(ss.getName());
  SpreadsheetApp.setActiveSpreadsheet(ss);
  var sheet_Active = SpreadsheetApp.getActiveSpreadsheet();
  var sheet1_of_sheet_Active = sheet_Active.getSheets()[0];
  var lastRow = sheet1_of_sheet_Active.getLastRow();

  var cell_QRCode_ActiveSheet = sheet1_of_sheet_Active.getRange("B" + lastRow.toString());
  var sQRCode = cell_QRCode_ActiveSheet.getDisplayValue();
  message = message + sQRCode;
  Logger.log(message);


  if (event.message) {
    // Bot added through @mention.
    message = message + " and you said: \"" + event.message.text + "\"";
  }

  return { "text": message };
}
