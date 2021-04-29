function mvgttgpsaigon(e) {
//  var email_to = "phaolo.hung.nguyen@gmail.com";
//  var subject = "test email"
//  var body =""
//  var sBankCode = "";
//  for (i=1; i<1000; i++){
//    sBankCode = Math.random().toString(36).slice(-8);
//    body = body + " " + sBankCode;
//  }                            
  
  var email_to = e.values[5]; 
  //Lấy giá trị của cột số 5 lưu vào biến email. Cột đầu tiên là cột 0.
  var name = e.values[3]; //Lấy giá trị của cột số 1 lưu vào biến name
  var tel = e.values[4]; //Lấy giá trị của cột số 3 lưu vào biến tel
  
  var subject = 'Chúc mừng ' + name + ' đã đăng ký tham dự Workshop thành công'; //Tiêu đề email
  var sBankCode = Math.random().toString(36).slice(-8);
  
  Logger.log(e.values);
  Logger.log(sBankCode);
  var body = 'Xin chào ' + name + ', <br>Cảm ơn bạn đã đăng ký tham dự  Đại hội Giới trẻ Mùa Chay TGP Sài Gòn ngày 27.03.2021.<p>' ; //Nội dung email
  body = body + 'BTC gửi đến ' + name + ' thông tin về chương trình như sau:.<p>' ;
  
  body = body + '<h2>I. THÔNG TIN CHƯƠNG TRÌNH: </h2><p>';
  body = body + '1. Đại hội giới trẻ mùa chay năm 2021 với chủ đề: “CHO TIỀM NĂNG TRỖI DẬY”<p>';
  body = body + '2. Thời gian: Ngày 27/03/2021. (Từ 13g30 đến 21g00) <p>';
  body = body + '3. Địa điểm: giáo xứ Tân Phước. Địa chỉ: 245 Nguyễn Thị Nhỏ, Phường 9, Tân Bình, TP.HCM. <p>';
  body = body + '4. Phí tham dự: 60,000/1 người (đã bao gồm phần ăn tối) <p>';
  
  body = body + '<h2>II. PHƯƠNG THỨC THANH TOÁN: </h2><p>';
  body = body + 'Để hoàn tất việc đăng ký, vui lòng thanh toán phí tham dự. Thông tin chuyển khoản như sau: <p>';
  body = body + 'Mã thanh toán của bạn là: <strong><span style="color: #ff0000;">DHGT-' + sBankCode + ' </span></strong>(bạn copy mã này vào phần nội dung thanh toán khi chuyển khoản nhé)<p>' ;
  body = body + 'Tên tài khoản: HUỲNH THANH VÂN <p>';
  body = body + '1. STK Ngân hàng Vietcombank:  025 100 250 9697 (Chi nhánh Bình Tây, HCM) <p>';
  body = body + '2. STK Ngân hàng Sacombank: 0500 53 999 538 (Chi nhánh Lái Thiêu, Bình Dương) <p>';
  body = body + 'Lưu ý: Trong nội dung chuyển khoản xin vui lòng ghi rõ <strong><span style="color: #ff0000;">DHGT-' + sBankCode + '</span></strong><p>';
  body = body + '<i>(Nếu bạn đóng phí cho nhiều người, vui lòng ghi rõ: DHGT-mã thanh toán 1 – mã thanh toán 2 – mã thanh toán 3. Ví dụ: bạn đóng phí cho 3 người tham dự, nội dung chuyển khoản là: DHGT-432hc0up – 346uh8uj – 873uth0i)</i> <p>';
  body = body + 'Vé tham dự sẽ được gửi qua email của bạn chậm nhất là 10:00 ngày 27.03.2021. <p>';
  body = body + 'Mã thanh toán có giá trị đến 9:00 ngày 27.03.2021. Quá thời gian đề cập, mã sẽ tự động huỷ. <p>';
  
  body = body + '-----------------------------------<p>';
  body = body + 'Mọi thông tin liên hệ xin vui lòng gửi tin nhắn cho BTC theo link: http://m.me/SG.YouthDay<p>';
  body = body + 'Hotline: 0933 20 24 27<p>';
  body = body + 'Hẹn gặp lại bạn tại Đại hội Giới trẻ Mùa Chay năm 2021 “CHO TIỀM NĂNG TRỖI DẬY” <p>';
  
  var htmlBody = HtmlService.createHtmlOutputFromFile('htmlbody');
  htmlBody.append('<span style="font-size: large; color: #333333;">'+ body + '</span>');
  htmlBody.append('<p><img src="https://gioitresaigon.net/wp-content/uploads/2020/01/MVGT-TGP-Sai-Gon-Logo-web.png" alt="MVGTSG" width="280" height="50">');
  var emailHtml = htmlBody.getContent();
//  Logger.log(emailHtml);

  reply_to = 'gioitresaigon@gmail.com';
  //email_cc = 'gioitresaigon@gmail.com';
  email_cc = 'workshopdhgtsg2021@gmail.com';
  
  MailApp.sendEmail(
    { to: email_to,
      replyTo: reply_to, 
      cc: email_cc, 
      subject: subject,
      htmlBody: emailHtml,
      name: "MVGTSG"}
  );
  
  //Move submited row to sheet SentEmail
  //https://developers.google.com/apps-script/reference/spreadsheet/spreadsheet-app
  var sheet_from = SpreadsheetApp.getActiveSpreadsheet();
  var sheet_to = SpreadsheetApp.openById('1Asb3aM2CTaSRhi0NAuoC1YaNyDLBBVCNWB4Dlo5gwFw');
  var sWorkshopID = MoveLastRow(sBankCode, sheet_from, sheet_to);
  
  UpdateSheetMonitoring(sWorkshopID, "Registered")

}

function MoveLastRow(sBankCode, sheet_from, sheet_to)
{
  var sheet1_of_sheet_from = sheet_from.getSheets()[0];
  var sheet1_of_sheet_to = sheet_to.getSheets()[0];

  // This logs the value in the very last cell of this sheet
  var lastRow = sheet1_of_sheet_from.getLastRow();
  var lastCol = sheet1_of_sheet_from.getLastColumn()
  var last_row_value = sheet1_of_sheet_from.getSheetValues(lastRow, 1, 1, lastCol); //getSheetValues(startRow, startColumn, numRows, numColumns)
  sheet1_of_sheet_from.deleteRow(lastRow);
  
  last_row_value[0].push(sBankCode);
  
  Logger.log(last_row_value);
  
  //Change Workshop Name to WorkShopID
  var sWorkshop = last_row_value[0][1];
  if (sWorkshop == 'Cùng đến với nhau - Nhóm và năng động nhóm/ Mr. Sỹ Bằng')
    sWorkshop = '1';
  else if (sWorkshop == 'Cùng đón nhận nhau - DISC / Ms. Châu - Mr. Tuấn')
    sWorkshop = '2';
  else if (sWorkshop == 'Cùng nhau dấn thân - Ơn gọi nghề nghiệp / Ms. Vân - Mr. Nhật')
    sWorkshop = '3';
  else if (sWorkshop == 'Cùng nhau phát triển toàn diện / Mr . Xuân Bằng - Mr. Thiên Ân')
    sWorkshop = '4';
  else if (sWorkshop == 'Cùng nhau là bạn với Media Social / Mr. Tân - Ms. Trâm')
    sWorkshop = '5';
  else if (sWorkshop == 'Cùng nhau cầu nguyện / Nhóm Taize - Mactynho')
    sWorkshop = '6';
  else
    sWorkshop = '7';
  last_row_value[0][1] = sWorkshop;
  
  //Strim FullName
  var sFullName = last_row_value[0][3].trim();
  sFullName = sFullName.replace(/  /g, " ");
  sFullName = sFullName.replace(/  /g, " ");
  last_row_value[0][3] = sFullName;
  
  //To avoid phone num 090xxxxx become 90xxxxx
  var sPhoneNo = last_row_value[0][6].toString();
  if (sPhoneNo[0] == '0' && sPhoneNo[1] != '0')
  {
    sPhoneNo = sPhoneNo.replace(/0/, "84");
    //sPhoneNo = '8' + sPhoneNo;
    last_row_value[0][6] = parseInt(sPhoneNo);
  }
  
  sheet1_of_sheet_to.appendRow(last_row_value[0]);
  return sWorkshop;
  //Logger.log(last_row_value);
}

// Update Sheet Monitor
// Check sheet 2 of Sheet SentEmail for Availability
function UpdateSheetMonitoring(sWorkshopID, sUserType)
{
  //Process on sheetMonitor
  var sheet_monitoring = SpreadsheetApp.openById('1FGYHCcwKMOQkcGSfayfmXIl9tZ9KyaQginDGmXqIgps');
  var sheet1_of_sheet_monitoring = sheet_monitoring.getSheets()[0];
  
  var iRow = parseInt(sWorkshopID)+1

  if (sUserType == "Registered")
    col = 'B';
  
  var sPosition = col + iRow.toString();
  var WorkshopID_cell_Registered = sheet1_of_sheet_monitoring.getRange(sPosition);
  var iCurrentRegistration = parseInt(WorkshopID_cell_Registered.getDisplayValue());
  iCurrentRegistration = iCurrentRegistration + 1;
  WorkshopID_cell_Registered.setValue(iCurrentRegistration);
  
  //Process on sheet2SentEmail
  var sheet2_of_sheet_SentEmail = SpreadsheetApp.openById('1Asb3aM2CTaSRhi0NAuoC1YaNyDLBBVCNWB4Dlo5gwFw').getSheets()[1];
  sPosition = 'F' + iRow.toString();
  var WorkshopID_cell_Availability = sheet2_of_sheet_SentEmail.getRange(sPosition);
  var iCurrentAvailability = parseInt(WorkshopID_cell_Availability.getDisplayValue());

  if (iCurrentAvailability - 10 < 0)
  {
    var url = 'https://chat.googleapis.com/v1/spaces/AAAA_L_3ErY/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=cQFX0R9_LPMM8zlblGieRJxG0bYQfi1YnRAbYVgqsaE%3D';
    var msg = 'Workshop ID '+sWorkshopID+ ' Availability is only '+ iCurrentAvailability.toString();
    var message_headers = {'Content-Type': 'application/json; charset=UTF-8'};
    
    var options = {
      method : 'POST',
      contentType: 'application/json',
      headers: message_headers,
      payload : JSON.stringify({ text: msg })     //Add your message
    }
    
    UrlFetchApp.fetch(url, options);
  }
  Logger.log(iCurrentAvailability);
}