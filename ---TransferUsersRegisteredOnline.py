CONTSTANT_WS_LIMITATION = {
    "tytt": 500,    # Tình yêu thánh thể
    "ntvndpt": 300, # Người trẻ và những điều phi thường
    "ntkn": 300,    # Người trẻ kết nối
    "ntvms": 400   # Người trẻ và mưu sinh
}

import secrets

# How to get Google Credential: https://www.twilio.com/blog/2017/02/an-easy-way-to-read-and-write-to-a-google-spreadsheet-in-python.html
#Below package needed to access Google Sheet
import gspread
from oauth2client.service_account import ServiceAccountCredentials

def generate_password():
    password = secrets.token_urlsafe(16)
    return password

def TransferUsersRegisteredOnline():  # check if paid and not_sent_email, then send email
    scope = [
        'https://www.googleapis.com/auth/drive',
        'https://www.googleapis.com/auth/drive.file'
        ]
    creds = ServiceAccountCredentials.from_json_keyfile_name('client_secret.json', scope)
    client = gspread.authorize(creds)

    sheetOnlineRegister = client.open("1-OnlineRegister").sheet1
    # sheetSuccessfulRegister = client.open("2-SuccessfulRegister").sheet1
    # sheetParseSuccessfulRegister = client.open("3-ParseSuccessfulRegister").sheet1
    # sheetTicketMonitoring = client.open("4-TicketMonitoring").sheet1


    sheetOnlineUsers = sheetOnlineRegister.get_all_values()
    iLength = len(sheetOnlineUsers)
    if iLength < 2:
        print("No user to transfer")
    else:
        for n in range(iLength, 1, -1):  # n=1 will not work
            user_info = sheetOnlineUsers[n - 1]

            user_info[7] = user_info[7].upper()
            # add more filed that suitable to sheet "2-List of registered user"
            user_info.append('N')  # field 8 = Check-in status = N
            user_info.append('N')  # field 9 = Paid = N

            iCount = int(user_info[5])   # arr[5] is number of ticket registered
            if iCount > 10:
                print("Double check user {}, {} tickets, , email {}, phone {}.".format(user_info[1], iCount, user_info[4], user_info[3]))
            else:
                if iCount == 1:
                    user_info[5] = generate_password()  # replace field of number of ticket by qrcode
                    # Create Full QR_Code
                    user_info[6] = user_info[1] + ' ' + user_info[2] + '\n' + user_info[7] + '\n' + user_info[5] + '\n' + 'Chân Phước Carlo Acutis'

                    sheetListOfRegisteredUser.append_row(user_info)
                else:
                    sName = user_info[1]
                    while iCount > 0 :
                        user_info[1] = sName + ' ' + str(iCount)
                        user_info[5] = generate_password()  # replace field of number of ticket by qrcode
                        # Create Full QR_Code
                        user_info[6] = user_info[1] + ' ' + user_info[2] + '\n' + user_info[7] + '\n' + user_info[5] + '\n' + 'Chân Phước Carlo Acutis'

                        sheetListOfRegisteredUser.append_row(user_info)
                        iCount -=1
                sheetListOfOnlineRegisteredUser.delete_rows(n, n)


TransferUsersRegisteredOnline()