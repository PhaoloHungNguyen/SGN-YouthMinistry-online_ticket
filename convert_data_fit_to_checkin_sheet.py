import secrets

# How to get Google Credential: https://www.twilio.com/blog/2017/02/an-easy-way-to-read-and-write-to-a-google-spreadsheet-in-python.html
#Below package needed to access Google Sheet
# import gspread
# from oauth2client.service_account import ServiceAccountCredentials

def generate_password():
    password = secrets.token_urlsafe(16)
    return password
def convert(arr_Successfull_Register_User):
    check_in_user = {
        "timestamp": arr_Successfull_Register_User[0],
        "FullName": arr_Successfull_Register_User[1],
        "Birthday": arr_Successfull_Register_User[2],
        "PhoneNo": arr_Successfull_Register_User[3],
        "Email": arr_Successfull_Register_User[4],
        "Gender": arr_Successfull_Register_User[5],
        "QRCode": generate_password(),
        "Object": arr_Successfull_Register_User[9],
        "CheckInStatus": "N",
        "Parish": arr_Successfull_Register_User[6],
        "Address": arr_Successfull_Register_User[7],
        "Group": arr_Successfull_Register_User[8]
    }
    return check_in_user

# scope = [
#     'https://www.googleapis.com/auth/drive',
#     'https://www.googleapis.com/auth/drive.file'
#     ]
# creds = ServiceAccountCredentials.from_json_keyfile_name('client_secret.json', scope)
# client = gspread.authorize(creds)
#
# sheetSuccessfulRegister = client.open("1-SuccessfullRegister").sheet1
# users = sheetSuccessfulRegister.get_all_values()
# user = users[1]
# converted_user = convert(user)
# print(converted_user["FullName"])
