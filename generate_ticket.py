# Original file name of Venn (Khoa) is check.py
import pyqrcode
import png
# from pyqrcode import QRCode
# import xlrd

from conf import conf

from send import send
from convert_data_fit_to_checkin_sheet import convert

import gspread
from oauth2client.service_account import ServiceAccountCredentials
from PIL import Image, ImageDraw, ImageFont

"""
For insert image to another image:
    https://note.nkmk.me/en/python-pillow-paste/
    https://www.geeksforgeeks.org/working-images-python/

For check installed packages, remove installed packages:
    https://note.nkmk.me/en/python-pillow-basic/
        pip show Pillow

For PIL to work:
    easy_install PILLOW
For png to work:
    pip install pypng

Insert text into image
    https://stackoverflow.com/questions/16373425/add-text-on-image-using-pil

"""

# define constant
DEFAULT_IMAGE_PATH = './images/'
DEFAULT_TEMPLATE_PATH = './template/'
FULL_PATH_TICKET_FILE_G = DEFAULT_TEMPLATE_PATH + 'e_ticket.png'   # template of ticket
FULL_PATH_TICKET_FILE_S = DEFAULT_TEMPLATE_PATH + 'e_ticket-S.png'

def make_qr_code(QRCode_image_name, sFullCode) : # QRCode_image_name w/o extention
    #sShortCode is utilized to create file name
    #sFullCode is utilized to create QRCode

    sQRCodeFile = DEFAULT_IMAGE_PATH + QRCode_image_name + '.png'
    url = pyqrcode.create(sFullCode, encoding = 'utf-8')
    url.png(sQRCodeFile, scale=10)

def add_image(sFullPathTicket, sFullPathQRCode, sFinalImageName) : # insert QRcode image to the ticket
    imgTicket = Image.open(sFullPathTicket)
    imgQR_code = Image.open(sFullPathQRCode).resize((518,518))  # need to resize because size of QR Code image depend on length of Code.

    back_im = imgTicket.copy()
    back_im.paste(imgQR_code, (102, 102))

    sFullPathFinalImageFile = DEFAULT_IMAGE_PATH + sFinalImageName

    back_im.save(sFullPathFinalImageFile, quality=95)

def add_text(sFullPathTicketFile, sText, sFinalImageName) :
    imgTicket = Image.open(sFullPathTicketFile)
    draw = ImageDraw.Draw(imgTicket)
    font = ImageFont.truetype("JetBrainsMono-Bold.ttf", 35)
    sName = sText

    # Get Left Align for sName: supposed that 32 characters/letters fit the width of the ticket.
    if len(sName) > 27:     # supposed that 27 words fit to the box
        arrNames = sName.split(" ")
        sName = arrNames[len(arrNames)-2] + ' ' + arrNames[len(arrNames)-1]
    nLeftAlign = (32 - len(sName))/2

    draw.text((nLeftAlign*22, 675), sName,(255,255,255),font=font)
    imgTicket.save(DEFAULT_IMAGE_PATH + sFinalImageName)

def generate_ticket(sShortCode, sFullCode, sFullName, sObject) :
    #Review Gsheet "List of registered user":
        # ShortCode is utilized to create file name of QRCode_image, name of specific user's ticket
        # FullCode i sutilized to generate QRCode.
    make_qr_code("QRCode", sFullCode) # generate <sShortCode>.png in DEFAULT_IMAGE_PATH

    sTempFileName = 'temp.png'
    sQRCodeFile = DEFAULT_IMAGE_PATH + 'QRCode' + '.png'
    if sObject == 'G':
        add_image(FULL_PATH_TICKET_FILE_G, sQRCodeFile, sTempFileName) # add QRCode image to ticket image -> final image temp.png in DEFAULT_IMAGE_PATH
    else:
        add_image(FULL_PATH_TICKET_FILE_S, sQRCodeFile, sTempFileName)
    sTicket_with_QRCodeFile = DEFAULT_IMAGE_PATH + sTempFileName
    sFinalImageName = sShortCode + '.png'
    add_text(sTicket_with_QRCodeFile, sFullName, sFinalImageName)  # add <sFullName> to image of Ticket_with_QRCode

# How to enable GoogleSheet API: https://developers.google.com/sheets/api/quickstart/python
# interact with GSheet: https://www.twilio.com/blog/2017/02/an-easy-way-to-read-and-write-to-a-google-spreadsheet-in-python.html

def check_and_send():  # check if paid and not_sent_email, then send email
    scope = [
        'https://www.googleapis.com/auth/drive',
        'https://www.googleapis.com/auth/drive.file'
        ]
    creds = ServiceAccountCredentials.from_json_keyfile_name('client_secret.json', scope)
    client = gspread.authorize(creds)

    sheetSuccessfulRegister = client.open("Copy 1-SuccessfullRegister").sheet1
    sheetCheckin = client.open("Copy 2-Check-in").sheet1

    users = sheetSuccessfulRegister.get_all_values()
#    col_Paid = sheetListOfRegisteredUser.col_values(10)

    iLength = len(users)
    if iLength < 2:
        print("No one to send ticket")
    else:
        for n in range(1, iLength, 1):  # i = [1-> iLength)
            user_info = convert(users[n])

            sFullName = user_info["FullName"]
            sEmail = user_info["Email"]
            sCode = user_info["QRCode"]
            sFull_Code = user_info["FullName"] + ' ' + user_info["Object"] + '\n' + user_info["PhoneNo"] + '\n' + user_info["QRCode"] + '\n' + 'Chân Phước Carlo Acutis'

            image_ticket = sCode + '.png'
            generate_ticket(sCode, sFull_Code, sFullName, user_info["Object"])
            send(sEmail, image_ticket, sFullName)
            # print(user_info)
            sheetCheckin.append_row([user_info["timestamp"], user_info["FullName"], user_info["Birthday"], user_info["PhoneNo"],
                                    user_info["Email"], user_info["Gender"], user_info["QRCode"], user_info["Object"],
                                    user_info["CheckInStatus"], user_info["Parish"], user_info["Address"], user_info["Group"]])
            sheetSuccessfulRegister.delete_rows(2, 2)
            print("Sent ticket to {}, {}".format(user_info["FullName"], user_info["Email"]))
        print("Please make sure {} doesnt' receive Alert Email from Google (sent successfully)!".format(conf.username))


def main() :
    check_and_send()



if __name__ == '__main__':
    main()
