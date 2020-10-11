import smtplib
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

def send_notification(to_addr):
    cc_addr = conf.cc_addr
    from_mail = conf.username
    from_mail_password = conf.password
    msg = MIMEMultipart()
    msg['From'] = "Carlo Acutis - MVGTSG <{}>".format(conf.username)
    msg['To'] = to_addr
    msg.add_header('reply-to', cc_addr)
    msg['CC'] = cc_addr
    msg['Subject'] = "[Hoà mạng cùng Carlo] Chúng ta có hẹn vào hôm nay (10.10.2020)"
    to_adds = [to_addr] + [cc_addr]

    html_part = MIMEMultipart(_subtype='related')

    with open(DEFAULT_TEMPLATE_PATH + 'notification.html') as html:
        # Create a text/plain message
        html_raw = MIMEText(html.read(), 'html')
    html_part.attach(html_raw)

    msg.attach(html_part)

    s = smtplib.SMTP('smtp.gmail.com', 587)
    s.starttls()
    s.login(from_mail, from_mail_password)
    text = msg.as_string()

    s.sendmail(from_mail, to_adds, text)
    s.quit()



# scope = [
#     'https://www.googleapis.com/auth/drive',
#     'https://www.googleapis.com/auth/drive.file'
#     ]
# creds = ServiceAccountCredentials.from_json_keyfile_name('template/client_secret1.json', scope)
# client = gspread.authorize(creds)
#
# sheetCheckin = client.open("1-SuccessfullRegister").sheet1
#
# emails = sheetCheckin.col_values(5)
# print("Start sent emails:")
# for i in range(1, len(emails), 1):
#     send_notification(emails[i])
#     print("Đã gửi đến {}".format(emails[i]))

send_notification("tranquyengiaothuy@gmail.com")
