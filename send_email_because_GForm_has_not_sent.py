import datetime
import secrets
import smtplib
import time
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email import encoders

import gspread
from oauth2client.service_account import ServiceAccountCredentials

from SendChatMessage import SendChatMessage
from conf import conf
# https://docs.python.org/3.4/library/email-examples.html
# https://www.thetopsites.net/article/53947988.shtml

DEFAULT_TEMPLATE_PATH = './template/'


def send_info(): # use list of user in FromGForm Sheet to send email
    scope = [
        'https://www.googleapis.com/auth/drive',
        'https://www.googleapis.com/auth/drive.file'
        ]
    creds = ServiceAccountCredentials.from_json_keyfile_name('template/dhgt.json', scope)
    client = gspread.authorize(creds)

    sheetSentEmail = client.open(conf.sSentEmailSheetName).sheet1
    sheet2SentEmail = client.open(conf.sSentEmailSheetName).worksheet('Sheet2')
    sheetThanhVienDangKy = client.open(conf.sFilledByGFormSheetName).sheet1
    sheetMonitor = client.open(conf.sMonitoringSheetName).sheet1

    # sheetMonitor.update_cell(5, 10, "row 5 col 10")

    iNumberOfEmailToSent = len(sheetThanhVienDangKy.get_all_records())   # ader excluded.
    # print("Sheet ThanhVienDangKy, Number of row = {}".format(iNumberOfEmailToSent))
    # print("Record 1 = {}".format(sheetThanhVienDangKy.get_all_records()[0]))
    for i in range(1,iNumberOfEmailToSent+1,1):
        user_sheet_ThanhVienDangKy = sheetThanhVienDangKy.row_values(2)

        to_addr = user_sheet_ThanhVienDangKy[5]
        cc_addr = conf.cc_addr
        from_mail = conf.username
        from_mail_password = conf.password

        sFullName = user_sheet_ThanhVienDangKy[3]
        print("{} - UserName = {} and user email = {}".format(i, sFullName, user_sheet_ThanhVienDangKy[5]))
        sBankCode = secrets.token_urlsafe(16)
        sBankCode="".join(sBankCode.split("-"))
        sBankCode="".join(sBankCode.split("_"))
        sBankCode = sBankCode.lower()
        sBankCode = sBankCode[0:8]

        sWorkshop = user_sheet_ThanhVienDangKy[1]
        if sWorkshop == 'Cùng đến với nhau - Nhóm và năng động nhóm/ Mr. Sỹ Bằng':
            sWorkshop = '1'
        elif sWorkshop == 'Cùng đón nhận nhau - DISC / Ms. Châu - Mr. Tuấn':
            sWorkshop = '2'
        elif sWorkshop == 'Cùng nhau dấn thân - Ơn gọi nghề nghiệp / Ms. Vân - Mr. Nhật':
            sWorkshop = '3'
        elif sWorkshop == 'Cùng nhau phát triển toàn diện / Mr . Xuân Bằng - Mr. Thiên Ân':
            sWorkshop = '4'
        elif sWorkshop == 'Cùng nhau là bạn với Media Social / Mr. Tân - Ms. Trâm':
            sWorkshop = '5'
        elif sWorkshop == 'Cùng nhau cầu nguyện / Nhóm Taize - Mactynho':
            sWorkshop = '6'
        else: sWorkshop = '7'
        print(user_sheet_ThanhVienDangKy)
        user_sheet_ThanhVienDangKy[1] = sWorkshop
        user_sheet_ThanhVienDangKy.append(sBankCode)
        print(user_sheet_ThanhVienDangKy)

        sBody = '<span style="font-size: large; color: #333333;">'
        sBody = sBody + 'Xin chào ' + sFullName + ', <br>Cảm ơn bạn đã đăng ký tham dự Đại hội Giới trẻ Mùa Chay TGP Sài Gòn ngày 27.03.2021.<p>'
        sBody = sBody + 'BTC gửi đến ' + sFullName + ' thông tin về chương trình như sau:.<p>'
        sBody = sBody + '<h2>I. THÔNG TIN CHƯƠNG TRÌNH: </h2><p>'
        sBody = sBody + '1. Đại hội giới trẻ mùa chay năm 2021 với chủ đề: “CHO TIỀM NĂNG TRỖI DẬY”<p>'
        sBody = sBody + '2. Thời gian: Ngày 27/03/2021. (Từ 13g30 đến 21g00) <p>'
        sBody = sBody + '3. Địa điểm: giáo xứ Tân Phước. Địa chỉ: 245 Nguyễn Thị Nhỏ, Phường 9, Tân Bình, TP.HCM. <p>'
        sBody = sBody + '4. Phí tham dự: 60,000/1 người (đã bao gồm phần ăn tối) <p>'

        sBody = sBody + '<h2>II. PHƯƠNG THỨC THANH TOÁN: </h2><p>'
        sBody = sBody + 'Để hoàn tất việc đăng ký, vui lòng thanh toán phí tham dự. Thông tin chuyển khoản như sau: <p>'
        sBody = sBody + 'Mã thanh toán của bạn là: <strong><span style="color: #ff0000;">DHGT-' + sBankCode + ' </span></strong>(bạn copy mã này vào phần nội dung thanh toán khi chuyển khoản nhé)<p>'
        sBody = sBody + 'Tên tài khoản: HUỲNH THANH VÂN <p>'
        sBody = sBody + '1. STK Ngân hàng Vietcombank:  025 100 250 9697 (Chi nhánh Bình Tây, HCM) <p>'
        sBody = sBody + '2. STK Ngân hàng Sacombank: 0500 53 999 538 (Chi nhánh Lái Thiêu, Bình Dương) <p>'
        sBody = sBody + 'Lưu ý: Trong nội dung chuyển khoản xin vui lòng ghi rõ <strong><span style="color: #ff0000;">DHGT-' + sBankCode + '</span></strong><p>'
        sBody = sBody + '<i>(Nếu bạn đóng phí cho nhiều người, vui lòng ghi rõ: DHGT-mã thanh toán 1 – mã thanh toán 2 – mã thanh toán 3. Ví dụ: bạn đóng phí cho 3 người tham dự, nội dung chuyển khoản là: DHGT-432hc0up – 346uh8uj – 873uth0i)</i> <p>'
        sBody = sBody + 'Vé tham dự sẽ được gửi qua email của bạn trong vòng 24 giờ sau khi BTC đã nhận được phí đóng góp. <p>'
        sBody = sBody + 'Mã thanh toán có giá trị tối đa 24h. Quá 24h, mã sẽ tự động huỷ. <p>'

        sBody = sBody + '-----------------------------------<p>'
        sBody = sBody + 'Mọi thông tin liên hệ xin vui lòng gửi tin nhắn cho BTC theo link: http://m.me/SG.YouthDay<p>'
        sBody = sBody + 'Hotline: 0933 20 24 27<p>'
        sBody = sBody + 'Hẹn gặp lại bạn tại Đại hội Giới trẻ Mùa Chay năm 2021 “CHO TIỀM NĂNG TRỖI DẬY” <p>'
        sBody = sBody + '<p><img src="https://gioitresaigon.net/wp-content/uploads/2020/01/MVGT-TGP-Sai-Gon-Logo-web.png" alt="MVGTSG" width="280" height="50">'
        sBody = sBody + '</span>'

        msg = MIMEMultipart()
        msg['From'] = "MVGTSG <{}>".format(conf.username)
        msg['To'] = to_addr
        msg.add_header('reply-to', cc_addr)
        msg['CC'] = cc_addr
        msg['Subject'] = "Chúc mừng {} đã đăng ký tham dự Workshop {} thành công".format(sFullName, sWorkshop)
        to_adds = [to_addr] + [cc_addr]

        html_part = MIMEMultipart(_subtype='related')

        # with open(DEFAULT_TEMPLATE_PATH + 'notification.html') as html:
            # Create a text/plain message
        html_raw = MIMEText(sBody, 'html')
        html_part.attach(html_raw)

        msg.attach(html_part)

        s = smtplib.SMTP('smtp.gmail.com', 587)
        s.starttls()
        s.login(from_mail, from_mail_password)
        text = msg.as_string()

        s.sendmail(from_mail, to_adds, text)
        s.quit()

        #Process tracking value on both sheetMonitor and sheet2SentTicket
        iWorkshopID = int(sWorkshop)
        WorkshopID_row_values = sheetMonitor.row_values(iWorkshopID+1)
        iCurrentRegistration = int(WorkshopID_row_values[1]) +1
        # iLimitedValue = int(WorkshopID_row_values[4])
        # print("Current value of Workshop {} is {}".format(iWorkshopID, iCurrentValue))
        sheetMonitor.update_cell(iWorkshopID + 1, 2, iCurrentRegistration)

        sheetSentEmail.append_row(user_sheet_ThanhVienDangKy)
        sheetThanhVienDangKy.delete_rows(2, 2)

        WorkshopID_row_values = sheet2SentEmail.row_values(iWorkshopID + 1)
        print(WorkshopID_row_values)
        iCurrentAvailability = int(WorkshopID_row_values[5])
        if iCurrentAvailability - 10 < 0:
            sMessage = "WARNING: WORKSHOP {} availability is {} \n{}".format(iWorkshopID, iCurrentAvailability, WorkshopID_row_values)
            print(sMessage)
            SendChatMessage(sMessage, 8)


def main() :
    while True:
        send_info()
        print(datetime.datetime.now().strftime("%H:%M:%S"))
        print("Email will be sent after 10 minutes")
        time.sleep(600)

if __name__ == '__main__':
    main()