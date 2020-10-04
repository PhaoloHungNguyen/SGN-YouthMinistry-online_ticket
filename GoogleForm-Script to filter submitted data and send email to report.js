function myFunction(e) {
  var YearOfBirth = e.values[2];
  var iYearOfBirth = parseInt(YearOfBirth.split('/')[2]);
  var iCurrentYear = new Date().getFullYear();
  var iAge = iCurrentYear - iYearOfBirth;
  if (iAge >=9 && iAge <=55) //in range of age to attend
  {
    //Move submited row to sheet SuccessfulRegister
    //https://developers.google.com/apps-script/reference/spreadsheet/spreadsheet-app
    var sheet_from = SpreadsheetApp.getActiveSpreadsheet();
    var sheet_to = SpreadsheetApp.openById('1a4EfskfCG_4C9ut_yYwo72MvDDBTOGWRUgGpPjLDet1c');
    MoveLastRow(sheet_from, sheet_to);
    
    UpdateSheetMonitoring('G', "Registrant")
    
    var email_subject = 'Bạn ' + e.values[1] + ' đăng ký thành công';
    var htmlBody = HtmlService.createTemplateFromFile('htmlBody_SuccessfulRegister');
    
    // set the values for the placeholders
    htmlBody.FullName = e.values[1];
  }
  else
  {
    //Move submited row to sheet RecycleBin
    var sheet_from = SpreadsheetApp.getActiveSpreadsheet();
    var sheet_to = SpreadsheetApp.openById('1voiunFxT5QKairhKyJsAo_nkGE2DN7hm2nPI3P81gO1s');
    MoveLastRow(sheet_from, sheet_to);
    
    var email_subject = 'Bạn ' + e.values[1] + ' đăng ký không thành công';
    
    var sFullName = e.values[1];
    var sParish = e.values[6];
    var sAddress = e.values[7];
    var sGroup = e.values[8];
    sFullName = sFullName.replace(/ /g, "%20");
    sParish = sParish.replace(/ /g, "%20");
    sAddress = sAddress.replace(/ /g, "%20");
    sGroup = sGroup.replace(/ /g, "%20");
    
    var link_to_edit = 'https://docs.google.com/forms/d/e/1FAIpQLSeVnPwGZx9XeVtYQrpP21E2noJHRL5Z4Gh7r1P377OaATckb1w/viewform?usp=pp_url&entry.551681654=' + sFullName + '&entry.1145740948=' + e.values[2] + '&entry.1946284035=' + e.values[5] + '&entry.1146267669=' + sParish + '&entry.1536259180=' + sAddress + '&entry.333704433=' + sGroup + '&entry.2029134402=' + e.values[3] + '&entry.1952234810=' + e.values[4];
    
    var htmlBody = HtmlService.createTemplateFromFile('htmlBody_FalseRegister');
    // set the values for the placeholders
    htmlBody.link_to_edit = link_to_edit;
    htmlBody.FullName = e.values[1];
  }
  
  var email_html = htmlBody.evaluate().getContent();
  Logger.log(email_html);

  sent_to = e.values[4];
  reply_to = 'saigon@gmail.com';
  email_cc = 'saigon@gmail.com';
    
  MailApp.sendEmail(
    { to: sent_to,
      replyTo: reply_to, 
      cc: email_cc, 
      subject: email_subject,
      htmlBody: email_html,
      name: "Carlo Acutis - MVGTSG"}
    ); 
  
}

function MoveLastRow(sheet_from, sheet_to)
{
  var sheet1_of_sheet_from = sheet_from.getSheets()[0];
  var sheet1_of_sheet_to = sheet_to.getSheets()[0];

  // This logs the value in the very last cell of this sheet
  var lastRow = sheet1_of_sheet_from.getLastRow();
  var lastCol = sheet1_of_sheet_from.getLastColumn()
  var last_row_value = sheet1_of_sheet_from.getSheetValues(lastRow, 1, 1, lastCol); //getSheetValues(startRow, startColumn, numRows, numColumns)
  sheet1_of_sheet_from.deleteRow(lastRow);
  
  last_row_value[0].push('G');
  
  //Strim FullName
  var sFullName = last_row_value[0][1].trim();
  sFullName = sFullName.replace(/  /g, " ");
  sFullName = sFullName.replace(/  /g, " ");
  last_row_value[0][1] = sFullName;
  
  //To avoid phone num 090xxxxx become 90xxxxx
  var sPhoneNo = last_row_value[0][3].toString();
  if (sPhoneNo[0] == '0' && sPhoneNo[1] != '0')
  {
    sPhoneNo = sPhoneNo.replace(/0/, "84");
    //sPhoneNo = '8' + sPhoneNo;
    last_row_value[0][3] = parseInt(sPhoneNo);
  }
  
  sheet1_of_sheet_to.appendRow(last_row_value[0]);
  
  //Logger.log(last_row_value);
}

// Update Sheet 4-TicketMonitoring
// sObject = 'G'-General or 'S'-Special
// QuantityType = "TicketReceiver" or "Registrant"
function UpdateSheetMonitoring(sObject, sQuantityType)
{
  var sheet_monitoring = SpreadsheetApp.openById('1pp54YQViQxZOxMf4QXlVBEEl2c3F3bbp-Ugd19ZS0ak');
  var sheet1_of_sheet_monitoring = sheet_monitoring.getSheets()[0];
  
  var row = '2'; //default row for sObject = 'G'
  if (sObject == 'S')
    row = '3';
  var col = 'C'; //default col for sQuantityType == "TicketReceiver"
  if (sQuantityType == "Registrant")
    col = 'D';
  
  var sPosition = col + row;
  var cell = sheet1_of_sheet_monitoring.getRange(sPosition);
  var iCurrentValue = parseInt(cell.getDisplayValue());
  iCurrentValue = iCurrentValue + 1;
  cell.setValue(iCurrentValue);
  Logger.log(iCurrentValue);
}