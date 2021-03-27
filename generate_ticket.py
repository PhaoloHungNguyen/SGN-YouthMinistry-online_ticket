# Original file name of Venn (Khoa) is check.py
import secrets
import time

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

from send_notification_email import send_notification

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
TICKET_WORKSHOP = ['e_ticket.png', 'e_ticket1.png', 'e_ticket2.png', 'e_ticket3.png', 'e_ticket4.png', 'e_ticket5.png', 'e_ticket6.png' ]

def make_qr_code(QRCode_image_name, sFullCode) : # QRCode_image_name w/o extention
    #sShortCode is utilized to create file name
    #sFullCode is utilized to create QRCode

    sQRCodeFile = DEFAULT_IMAGE_PATH + QRCode_image_name + '.png'
    url = pyqrcode.create(sFullCode, encoding = 'utf-8')
    url.png(sQRCodeFile, scale=7)

def add_image(sFullPathTicket, sFullPathQRCode, sFinalImageName) : # insert QRcode image to the ticket
    imgTicket = Image.open(sFullPathTicket)
    imgQR_code = Image.open(sFullPathQRCode).resize((518,518))  # need to resize because size of QR Code image depend on length of Code.

    back_im = imgTicket.copy()
    back_im.paste(imgQR_code, (102, 222))

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

    draw.text((nLeftAlign*22, 1100), sName,(255,255,255),font=font)
    imgTicket.save(DEFAULT_IMAGE_PATH + sFinalImageName)

def generate_ticket(sShortCode, sFullCode, sFullName, sWorkshopID) :
    #Review Gsheet "List of registered user":
        # ShortCode is utilized to create file name of QRCode_image, name of specific user's ticket
        # FullCode i sutilized to generate QRCode.
    make_qr_code("QRCode", sFullCode) # generate <sShortCode>.png in DEFAULT_IMAGE_PATH

    sTempFileName = 'temp.png'
    sQRCodeFile = DEFAULT_IMAGE_PATH + 'QRCode' + '.png'
    iTicketNo = int(sWorkshopID)

    # Use workshop 3 for the ones who do not choose workshop.
    if iTicketNo > 6:
        iTicketNo = 3

    sFullPathOfTicket = DEFAULT_TEMPLATE_PATH + TICKET_WORKSHOP[iTicketNo]
    add_image(sFullPathOfTicket, sQRCodeFile, sTempFileName) # add QRCode image to ticket image -> final image temp.png in DEFAULT_IMAGE_PATH

    sTicket_with_QRCodeFile = DEFAULT_IMAGE_PATH + sTempFileName
    sFinalImageName = sShortCode + '.png'
    add_text(sTicket_with_QRCodeFile, sFullName, sFinalImageName)  # add <sFullName> to image of Ticket_with_QRCode

# How to enable GoogleSheet API: https://developers.google.com/sheets/api/quickstart/python
# interact with GSheet: https://www.twilio.com/blog/2017/02/an-easy-way-to-read-and-write-to-a-google-spreadsheet-in-python.html

def check_and_send():  # nd email to users in SentEmail sheet, move these users to GSheet Checkn
    scope = [
        'https://www.googleapis.com/auth/drive',
        'https://www.googleapis.com/auth/drive.file'
        ]
    creds = ServiceAccountCredentials.from_json_keyfile_name('template/dhgt.json', scope)
    client = gspread.authorize(creds)

    sheetSentEmail = client.open(conf.sSentEmailSheetName).sheet1
    sheet2SentEmail = client.open(conf.sSentEmailSheetName).worksheet('Sheet2')
    sheetCheckin = client.open(conf.sCheckInSheetName).sheet1
    sheet1PayslipVCB = client.open(conf.sPayslipVCBSheetName).sheet1
    sheet2PayslipVCB = client.open(conf.sPayslipVCBSheetName).worksheet('Sheet2')

    # sheetRecycleBin = client.open(conf.sRecycleBinSheetName).sheet1

    users = sheetSentEmail.get_all_values()          # users index starts from 0
    users_col_BankCode = sheetSentEmail.col_values(8)
    VCB_col_BankCodes = sheet1PayslipVCB.col_values(5) # col number start from 1
    VCB_transactions = sheet1PayslipVCB.get_all_values()

    # print(users[1][7])
    # print(VCB_col_BankCodes)
    # print(SCB_col_BankCodes)
    # print(users_col_BankCode)

    iLength_VCB_BankCodes = len(VCB_col_BankCodes)
    if iLength_VCB_BankCodes < 2:
        print("No VCB_BankCode to send ticket")
    else:
        for n in range(1, iLength_VCB_BankCodes, 1):  # start from 1 to avoid head row, range [start, stop)
            sBankCode = VCB_col_BankCodes[n]
            #initial process to avoid user typo
            sBankCode = sBankCode.lower()
            sBankCode = sBankCode.replace("_", "-")
            sBankCode = sBankCode.replace("-", " ")    # final string: ...dhgt bankcode1 bankcode2.ct tu...


            # sBankCode - split head part
            sBankCodeListNotPair = ""
            arr_temp = sBankCode.split("dhgt")
            iTemp = len(arr_temp)
            if iTemp == 2: # as expected
                sBankCode = arr_temp[1]
            elif iTemp == 1:
                print("Important note: There is no string dhgt to split")
                sBankCodeListNotPair = sBankCode
            else: #iTemp > 2
                print("Important note: sBankCode was splited into {} by string dhgt".format(iTemp))
                del arr_temp[0]
                sBankCode = arr_temp[0]
                del arr_temp[0]
                for x in arr_temp:
                    sBankCode = sBankCode + "dhgt" + x

            if iTemp >=2: #exist string dhgt in sBankCode, then continue to split tail part.
                sBankCode = sBankCode.split(".ct tu")[0]
                sBankCode = ' '.join(sBankCode.split())    #trim & replace "  " by " "
                arrBankCode = sBankCode.split(" ")         #split codes separated by space
                # print(sBankCode)
                # print(arrBankCode)

                # pair arrBankCode with BankCode in EmailSent Sheet
                iLength_arrBankCode = len(arrBankCode)
                sBankCodeListNotPair = ""
                for m in range(0, iLength_arrBankCode, 1):
                    iFlag = -1
                    for i, sBankCode_element in enumerate(users_col_BankCode):
                        if arrBankCode[m] == sBankCode_element:
                            # print(users_col_BankCode)
                            # print("BankCode = {}, and index = {}. Generate ticket".format(arrBankCode[m], i))
                            iFlag = i

                            # generate and senicket.
                            user = users[i]
                            print("sent ticket to {} and move user i={} to sheet Checkin".format(user, i))

                            sFullName = user[3]
                            sEmail = user[5]
                            sWorkshopID = user[1]
                            sPhoneNo = user[6]
                            sQRCode = secrets.token_urlsafe(16)
                            sFull_Code = sFullName + ' ' + sWorkshopID + '\n' + sPhoneNo + '\n' + sQRCode + '\n' + 'ĐHGT 2021'

                            image_ticket = sQRCode + '.png'
                            generate_ticket(sQRCode, sFull_Code, sFullName, sWorkshopID)
                            if sWorkshopID != "7":
                                send(sEmail, image_ticket, sFullName, sWorkshopID)
                                iWorkshopID = int(sWorkshopID)
                                iCurrentValue = int(sheet2SentEmail.row_values(iWorkshopID + 1)[2])
                                iCurrentValue = iCurrentValue + 1
                                print("Workshop {}, number of generated ticket = {}".format(iWorkshopID, iCurrentValue))
                                sheet2SentEmail.update_cell(iWorkshopID + 1, 3, iCurrentValue) #update_cell(FirstRow=1, FirstCol=1, value)
                            else:
                                print("User {}, email {}, sPhone {} has not chosen Workshop. Move him back. Recover paid BankCode {}".format(sFullName, sEmail, sPhoneNo, sBankCode_element))

                            # Move user from sheet SentEmail to CheckIn.
                            user.append(sQRCode)
                            user.append("0")
                            sheetCheckin.append_row(user)
                            time.sleep(5) # to avoid quota of write request exceeded.
                            sheetSentEmail.delete_rows(i+1,i+1)

                            # After delete raw in Gsheet SentEmail, also remove raw in array.
                            del users[i]
                            del users_col_BankCode[i]
                            break
                    if iFlag == -1:
                        sBankCodeListNotPair = sBankCodeListNotPair + " " + arrBankCode[m]
                if sBankCodeListNotPair != "":
                    sBankCodeListNotPair = "DHGT-" + sBankCodeListNotPair
            if sBankCodeListNotPair != "":
                print("BankCode {} is not in sheet EmailSent. User made typo in payment processing".format(sBankCodeListNotPair))

            #append sBankCodeListNotPair to last col and move data between 2 sheet of BankCode.
            arrTransaction = VCB_transactions[n]
            arrTransaction.append(sBankCodeListNotPair)
            # print("Move {} to VCB sheet 2".format(arrTransaction))
            sheet2PayslipVCB.append_row(arrTransaction)
            sheet1PayslipVCB.delete_rows(2,2)





    # users = sheetSentEmail.get_all_values()
#
#     iLength = len(users)
#     if iLength < 2:
#         print("No one to send ticket")
#     else:
#         for n in range(1, iLength, 1):  # i = [1-> iLength)
#             user_info = convert(users[n])
#
#             sFullName = user_info["FullName"]
#             sEmail = user_info["Email"]
#             sCode = user_info["QRCode"]
#             sFull_Code = user_info["FullName"] + ' ' + user_info["Object"] + '\n' + user_info["PhoneNo"] + '\n' + user_info["QRCode"] + '\n' + 'Chân Phước Carlo Acutis'
#
#             image_ticket = sCode + '.png'
#             generate_ticket(sCode, sFull_Code, sFullName, user_info["Object"])
#             send(sEmail, image_ticket, sFullName)
#             # send_notification(sEmail)
#
#             # print(user_info)
#             sheetCheckin.append_row([user_info["timestamp"], user_info["FullName"], user_info["Birthday"], user_info["PhoneNo"],
#                                     user_info["Email"], user_info["Gender"], user_info["QRCode"], user_info["Object"],
#                                     user_info["CheckInStatus"], user_info["Parish"], user_info["Address"], user_info["Group"]])
#             sheetSentEmail.delete_rows(2, 2)
#             print("Sent ticket to {}, {}".format(user_info["FullName"], user_info["Email"]))
#         print("Please make sure {} doesnt' receive Alert Email from Google (sent successfully)!".format(conf.username))
#

def main() :
    # while True:
    #     print("Gửi email")
        check_and_send()
    #     print("1 phút sau sẽ gửi tiếp")
    #     time.sleep(60)



if __name__ == '__main__':
    main()
