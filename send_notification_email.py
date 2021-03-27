import smtplib
import time
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email import encoders

import gspread
from oauth2client.service_account import ServiceAccountCredentials

from conf import conf
# https://docs.python.org/3.4/library/email-examples.html
# https://www.thetopsites.net/article/53947988.shtml

DEFAULT_TEMPLATE_PATH = './template/'

def send_notification(to_addr, sFullName, sWorkshopID, sSubject):
    cc_addr = conf.cc_addr
    from_mail = conf.username
    from_mail_password = conf.password
    msg = MIMEMultipart()
    msg['From'] = "MVGTSG <{}>".format(conf.username)
    msg['To'] = to_addr
    msg.add_header('reply-to', cc_addr)
    msg['CC'] = cc_addr
    msg['Subject'] = sSubject
    to_adds = [to_addr] + [cc_addr]

    html_part = MIMEMultipart(_subtype='related')
    sFileHtmlName = "ws{}.html".format(sWorkshopID)
    with open(DEFAULT_TEMPLATE_PATH + sFileHtmlName) as html:
        # Create a text/plain message
        # html.read() --> <h1>abcd</h1><p>
        sHtml = html.read()
        sHtml = sHtml.replace("@FULLNAME", sFullName)
        # html_raw = MIMEText(html.read(), 'html')
        html_raw = MIMEText(sHtml, 'html')
    html_part.attach(html_raw)

    msg.attach(html_part)

    s = smtplib.SMTP('smtp.gmail.com', 587)
    s.starttls()
    s.login(from_mail, from_mail_password)
    text = msg.as_string()

    s.sendmail(from_mail, to_adds, text)
    s.quit()

def send_notification_BankCodeCancellation(to_addr, sFullName, sBankCode, sExpireDate):
    cc_addr = conf.cc_addr
    from_mail = conf.username
    from_mail_password = conf.password
    msg = MIMEMultipart()
    msg['From'] = "MVGTSG <{}>".format(conf.username)
    msg['To'] = to_addr
    msg.add_header('reply-to', cc_addr)
    msg['CC'] = cc_addr
    msg['Subject'] = "ĐẠI HỘI GIỚI TRẺ - THÔNG BÁO MÃ THANH TOÁN {} HẾT HẠN".format("DHGT-"+sBankCode)
    to_adds = [to_addr] + [cc_addr]

    sBody = '<span style="font-size: large; color: #333333;">'
    sBody = sBody + 'Chào bạn {},<br>'.format(sFullName)
    sBody = sBody + 'Cảm ơn bạn đã đăng ký tham dự Đại hội Giới trẻ Mùa Chay TGP Sài Gòn ngày 27.03.2021 với chủ đề <strong>Cho tiềm năng trỗi dậy.</strong><p>'
    sBody = sBody + '- Thời gian: Ngày 27/03/2021. (Từ 13g30 đến 21g00)<p>'
    sBody = sBody + '- Địa điểm: giáo xứ Tân Phước. Địa chỉ: 245 Nguyễn Thị Nhỏ, Phường 9, Tân Bình, TP.HCM.<p>'
    sBody = sBody + '- Phí tham dự: 60,000/1 người (đã bao gồm phần ăn tối)<p>'
    sBody = sBody + 'Đại hội rất tiếc phải thông báo rằng <span style="color: #ff0000;">mã thanh toán phí tham dự của bạn DHGT-{} đăng ký lúc {} đã hết hạn, </span>vì thế, mã đã bị hủy. Bạn vui lòng đăng ký lại tại link https://forms.gle/6RNomQn57QCvSDbWA và chọn cho mình chủ đề workshop phù hợp.<p>'.format(sBankCode, sExpireDate)
    sBody = sBody + 'Đừng quên kiểm tra email và hoàn tất thanh toán trong thời gian giữ chỗ, bạn nhé! <p>'
    sBody = sBody + 'Hẹn gặp bạn tại đại hội. <p>'
    sBody = sBody + '____________________ <p>'

    sBody = sBody + 'Mọi thông tin liên hệ xin vui lòng gửi tin nhắn cho BTC theo link: http://m.me/SG.YouthDay<p>'
    sBody = sBody + 'Hotline: 0933 20 24 27<p>'
    sBody = sBody + 'Hẹn gặp lại bạn tại Đại hội Giới trẻ Mùa Chay năm 2021 “CHO TIỀM NĂNG TRỖI DẬY”<p>'
    sBody = sBody + '<p><img src="https://gioitresaigon.net/wp-content/uploads/2020/01/MVGT-TGP-Sai-Gon-Logo-web.png" alt="MVGTSG" width="280" height="50">'
    sBody = sBody + '</span>'

    html_part = MIMEMultipart(_subtype='related')

    html_raw = MIMEText(sBody, 'html')
    html_part.attach(html_raw)

    msg.attach(html_part)

    s = smtplib.SMTP('smtp.gmail.com', 587)
    s.starttls()
    s.login(from_mail, from_mail_password)
    text = msg.as_string()

    s.sendmail(from_mail, to_adds, text)
    s.quit()

def main() :

    scope = [
        'https://www.googleapis.com/auth/drive',
        'https://www.googleapis.com/auth/drive.file'
        ]
    creds = ServiceAccountCredentials.from_json_keyfile_name('template/dhgt.json', scope)
    client = gspread.authorize(creds)

    # # Send General Notification to users in sheet Check-in
    # sheetCheckin = client.open("CheckIn").sheet1
    # users = sheetCheckin.get_all_values()
    # iNumberOfUsers = len(users)
    # for i in range(1, iNumberOfUsers, 1):
    #     sWorkshopID= "General"   # in function send_Notification(), filename = "ws" + sWorkshopID
    #     sGeneralNotification = users[i][11]
    #     sEmail = users[i][5]
    #     sFullName = users[i][3]
    #
    #     if sGeneralNotification != "1":
    #         sSubject = "CHO TIỀM NĂNG TRỖI DẬY - NHẮC THAM DỰ ĐẠI HỘI GIỚI TRẺ MÙA CHAY 2021"
    #         send_notification(sEmail, sFullName, sWorkshopID, sSubject)
    #         sheetCheckin.update_cell(i+1, 12, "1")
    #         print("Đã gửi đến {}, Workshop {}".format(sEmail, sWorkshopID))

    # Send Notification of Trainers to users in sheet Check-in
    sheetCheckin = client.open("CheckIn").sheet1
    users = sheetCheckin.get_all_values()
    iNumberOfUsers = len(users)
    for i in range(1, iNumberOfUsers, 1):
        sWorkshopID= users[i][1]
        sTrainerNotification = users[i][10]
        sEmail = users[i][5]
        sFullName = users[i][3]

        if sWorkshopID == "3" and sTrainerNotification != "1":
            sSubject = "WORKSHOP CÙNG NHAU DẤN THÂN - ƠN GỌI NGHỀ NGHIỆP"
            send_notification(sEmail, sFullName, sWorkshopID, sSubject)
            sheetCheckin.update_cell(i+1, 11, "1")
            print("Đã gửi đến {}, Workshop {}".format(sEmail, sWorkshopID))
        elif sWorkshopID == "6" and sTrainerNotification != "1":
            sSubject = "WORKSHOP CÙNG NHAU CẦU NGUYỆN"
            send_notification(sEmail, sFullName, sWorkshopID, sSubject)
            sheetCheckin.update_cell(i+1, 11, "1")
            print("Đã gửi đến {}, Workshop {}".format(sEmail, sWorkshopID))

    # # Send notification of BankCode Cancellation to users in sheRecycleBin
    # sheetRecycleBin = client.open(conf.sRecycleBinSheetName).sheet1
    #
    # users = sheetRecycleBin.get_all_values()
    # iNumberOfUsers = len(users)
    # for i in range(1, iNumberOfUsers, 1):
    #     if users[i][9] != "1": # not yet notified
    #         sEmail = users[i][5]
    #         sFullName = users[i][3]
    #         sBankCode = users[i][7]
    #         sExpireDate = users[i][0]
    #         sheetRecycleBin.update_cell(i+1, 10, "1")
    #         send_notification_BankCodeCancellation(sEmail, sFullName, sBankCode, sExpireDate)
    #         print("Email to {} to notify DHGT-{} BankCode Cancellation".format(sEmail, sBankCode, sExpireDate))

if __name__ == '__main__':
    main()


