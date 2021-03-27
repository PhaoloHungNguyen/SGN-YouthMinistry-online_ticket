import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email import encoders
from conf import conf
# https://docs.python.org/3.4/library/email-examples.html
# https://www.thetopsites.net/article/53947988.shtml

DEFAULT_TEMPLATE_PATH = './template/'

def send(to_addr, image_name, full_name, sWorkshopID):
    # image_logo = 'MVGT-TGP-Sai-Gon-Logo-web.png'
    # image_carlo = 'carlo_acutis.jpg'
    # image_ticket = 'e_ticket.png'
    # full_name = 'Nguyễn Kiểm Email'

    cc_addr = conf.cc_addr
    from_mail = conf.username
    from_mail_password = conf.password
    msg = MIMEMultipart()
    msg['From'] = "MVGTSG <{}>".format(conf.username)
    msg['To'] = to_addr
    msg.add_header('reply-to', cc_addr)
    msg['CC'] = cc_addr
    msg['Subject'] = "[ {} ] VÉ THAM DỰ ĐẠI HỘI GIỚI TRẺ 2021 của {}".format("MVGTSG", full_name)
    to_cc_addr = [to_addr] + [cc_addr]

    html_part = MIMEMultipart(_subtype='related')

    with open(DEFAULT_TEMPLATE_PATH + 'ticket.html') as html:
        # Create a text/plain message
        sHead = '<p style="text-align: left;"><span style="font-size: large;">Xin chào {} </span></p>'.format(full_name)
        sHead = sHead + '<p style="text-align: left;"><span style="font-size: large;">Chúc mừng bạn đã đăng ký thành công Đại hội Giới trẻ Mùa Chay Tổng giáo phận Sài Gòn<strong> CHO TIỀM NĂNG TRỖI DẬY</strong> với Workshop {}.</span></p>'.format(sWorkshopID)
        html_raw = MIMEText(sHead + html.read(), 'html')

    # If not read html_raw from file, using html string instead, using bw code
    # Set correct scale of image size so that the QRCode image is a square which is catchable by app
    #
    # body = MIMEText('''<p>Hello <strong> {} </strong> </p>
    # <img src="cid:logo" alt="Logo" style="width:410px;height:72px;"/>
    # <img src="cid:carlo" alt="Carlo" style="width:576px;height:322px;"/>
    # <img src="cid:ticket" alt="Ticket" style="width:576px;height:1024px;"/>'''.format(full_name), _subtype='html')
    html_part.attach(html_raw)

    # img_data = open("./images/{}".format(image_logo), 'rb').read()
    # img = MIMEImage(img_data, 'jpg')
    # img.add_header('Content-Id', '<logo>')  # angle brackets are important
    # img.add_header("Content-Disposition", "inline", filename="logo")
    # html_part.attach(img)

    # img_data = open("./images/{}".format(image_carlo), 'rb').read()
    # img = MIMEImage(img_data, 'jpg')
    # img.add_header('Content-Id', '<carlo>')  # angle brackets are important
    # img.add_header("Content-Disposition", "inline", filename="carlo")
    # html_part.attach(img)

    img_data = open("./images/{}".format(image_name), 'rb').read()
    img = MIMEImage(img_data, 'jpg')
    img.add_header('Content-Id', '<ticket>')  # angle brackets are important
    img.add_header("Content-Disposition", "inline", filename="ticket")
    html_part.attach(img)

    msg.attach(html_part)

    s = smtplib.SMTP('smtp.gmail.com', 587)
    s.starttls()
    s.login(from_mail, from_mail_password)
    text = msg.as_string()

    s.sendmail(from_mail, to_cc_addr, text)
    s.quit()

# send("maryhoaian@gmail.com", "DhsinmZI3vRa3DyZeMWEew.png", "An direct email", "6")
