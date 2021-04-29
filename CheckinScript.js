function myFunction() {
  // Process on ActiveSheet Line
  var sheet_Active = SpreadsheetApp.getActiveSpreadsheet();
  var sheet1_of_sheet_Active = sheet_Active.getSheets()[0];
  var lastRow = sheet1_of_sheet_Active.getLastRow();

  var cell_QRCode_ActiveSheet = sheet1_of_sheet_Active.getRange("B" + lastRow.toString());
  var sQRCode = cell_QRCode_ActiveSheet.getDisplayValue() + "\n1 \n2 \n3"; //add substring \n1 \n2 \n3 to avoid out of index range array[2] below
  sQRCode = sQRCode.split("\n")[2];

  // if (lastRow > 1)
  //   sheet1_of_sheet_Active.deleteRow(sheet1_of_sheet_Active.getLastRow())

  // work on sheet CheckIn
  var sheet_checkin = SpreadsheetApp.openById('1CSBt8CVmDs1-KUWDg8zXONw7qWUnbshZF3N_pLkcWGY')
  var sheet1_of_sheet_checkin = sheet_checkin.getSheets()[0];
  lastRow = sheet1_of_sheet_checkin.getLastRow();

  var range_col_QRCode = sheet1_of_sheet_checkin.getRange(1,9,lastRow,1);
  var col_QRCode_Values = range_col_QRCode.getValues();
  
  var iIndex =1;

  // Logger.log(col_QRCode_Values[iIndex][0]);
  // Logger.log(sQRCode);
  // Logger.log(col_QRCode_Values[iIndex][0]);
  while (iIndex < lastRow) {
    if (col_QRCode_Values[iIndex][0] == sQRCode )
    {
      cell_Status = sheet1_of_sheet_checkin.getRange('J' + (iIndex+1).toString());
      sStatus = cell_Status.getDisplayValue();

      var cell_FullName_sheetCheckIn = sheet1_of_sheet_checkin.getRange("D"+(iIndex+1).toString())
      var sFullName = cell_FullName_sheetCheckIn.getDisplayValue();
      if (sStatus != "1")
      {
        cell_Status.setValue("1");

        // var cell_ParticipantNumber_ActiveSheet = sheet1_of_sheet_Active.getRange("C1");
        // var iParticipant = parseInt(cell_ParticipantNumber_ActiveSheet.getDisplayValue());
        // iParticipant = iParticipant + 1;
        // cell_ParticipantNumber_ActiveSheet.setValue(iParticipant)

        
        // var sMessage = iParticipant.toString() + ' - ' + sFullName;
        var sMessage = sFullName
        SendMessageToRoom(sMessage, 3);
        // sMessage = sMessage + " - Line 3 - " + new Date().toLocaleString("VN", {timeZone: "Asia/Saigon"});
        // SendMessageToRoom(sMessage, 7);
      } else
      {
        var sMessage = "Vé đã được sử dụng " + " - " + sFullName + " - " + new Date().toLocaleString("VN", {timeZone: "Asia/Saigon"});
        Logger.log("Date here:")
        Logger.log(new Date().toLocaleString("VN", {timeZone: "Asia/Saigon"}));
        SendMessageToRoom(sMessage, 3);
        // sMessage = "Vé đã được sử dụng " + " - Line 3 - " + sFullName + " - " + new Date().toLocaleString("VN", {timeZone: "Asia/Saigon"});
        // SendMessageToRoom(sMessage, 8);
      }
      break;
    }
    iIndex = iIndex + 1;
  }

  if (iIndex >= lastRow) // not found sQRCode in database
  {
    var sMessage = "Vé không hợp lệ. Mã QRCode " + " - " + sQRCode + " - Line 3 - " + new Date().toLocaleString("VN", {timeZone: "Asia/Saigon"});
    SendMessageToRoom(sMessage, 3);
    SendMessageToRoom(sMessage, 8);
  }
}

function SendMessageToRoom(sMessage, iLine)
{
  var url = arrWebhookURL = ["",
    "https://chat.googleapis.com/v1/spaces/AAAAkji1QW0/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=zQxWezrpSfS-NJMd2VwzQvWKk0b-2-oVlN9eKrogD-8%3D",
    "https://chat.googleapis.com/v1/spaces/AAAAQUCviZM/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=LymSfnM5X8h6XfS1F4H2GleaESnkPjvVJzK2W7fciAI%3D",
    "https://chat.googleapis.com/v1/spaces/AAAA1hxkzUc/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=25ePZ1VCovQvzGt4hEyjjYePGT49BHuVK2pdg0kAKx0%3D",
    "https://chat.googleapis.com/v1/spaces/AAAAplkgDkk/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=Ui_7fs6h3Td-UN6rVM55ak79SYpIXOqRTw6KYR9eE1c%3D",
    "https://chat.googleapis.com/v1/spaces/AAAADcG1Lk8/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=xKCad5PB3hnIwJqN9mF07TdBQncSaipwQj0MGGz-EGA%3D",
    "https://chat.googleapis.com/v1/spaces/AAAAi6_eghY/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=9lfKSCcF-TFLaLrRqpLRNs0lbQK-989ttial10VCn9o%3D",
    "https://chat.googleapis.com/v1/spaces/AAAAySRfkss/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=sgDZ5-k6-XNv4uetnmGHxL3FSv0SDsGTwKxyjUq7T34%3D",
    "https://chat.googleapis.com/v1/spaces/AAAA_L_3ErY/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=cQFX0R9_LPMM8zlblGieRJxG0bYQfi1YnRAbYVgqsaE%3D",
    ];

  var msg = sMessage;
  var message_headers = {'Content-Type': 'application/json; charset=UTF-8'};

  var options = {
    method : 'POST',
    contentType: 'application/json',
    headers: message_headers,
    payload : JSON.stringify({ text: msg })     //Add your message
    }

  UrlFetchApp.fetch(url[iLine], options);
}
