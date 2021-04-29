import datetime
import threading
import time

from conf import conf

import gspread
from oauth2client.service_account import ServiceAccountCredentials

from SendChatMessage import SendChatMessage

scope = [
    'https://www.googleapis.com/auth/drive',
    'https://www.googleapis.com/auth/drive.file'
    ]

arr_sheet_of_line = list()
arr_ParticipantNumber_of_line = list()
iParticipantNumber_of_AllLine = 0
for i in range(1,5,1): #[1,4)  number of lines = 4
    creds = ServiceAccountCredentials.from_json_keyfile_name("./template/client_secret{}.json".format(i), scope)
    client = gspread.authorize(creds)
    arr_sheet_of_line.append(client.open("Line_{}".format(i)).sheet1)
    arr_ParticipantNumber_of_line.append(int(arr_sheet_of_line[i - 1].row_values(1)[2]))
    iParticipantNumber_of_AllLine = iParticipantNumber_of_AllLine + arr_ParticipantNumber_of_line[i-1]

creds = ServiceAccountCredentials.from_json_keyfile_name("./template/dhgt.json", scope)
client = gspread.authorize(creds)
sheet_Checkin = client.open(conf.sCheckInSheetName).sheet1

col_QRcodes = sheet_Checkin.col_values(9)
time.sleep(2)
col_CheckInStatus = sheet_Checkin.col_values(10)
time.sleep(2)
col_FullName = sheet_Checkin.col_values(4)
time.sleep(2)

def dhgt_checkin(iLineNo):
    global threads
    global iParticipantNumber_of_AllLine
    sheet_of_lineNo = arr_sheet_of_line[iLineNo - 1]

    while True:
        try:
            row_to_check = sheet_of_lineNo.row_values(2)
        except Exception as e:
            print("Important Note: Can't read Gsheet Line{},\nMaybe caused by bad internet connection".format(iLineNo))
        else:
           if row_to_check:
                # print(row_to_check)
                try:
                    arrRow = row_to_check[1].split("\n")
                    sQRCode = arrRow[2]
                except Exception as e:
                    # print("Vé không hợp lệ")
                    sQRCode = ""

                try:
                    value_index = col_QRcodes.index(sQRCode)    # index of string sQRCode in string list col_QRCodes
                except Exception as e:
                    value_index = -1 # fake ticket

                if value_index > 0: # Vé Hợp Lệ
                    # print("value_index = {}".format(value_index))
                    status_of_checking_ticket = col_CheckInStatus[value_index]  # get current status of checking_ticket
                    if status_of_checking_ticket != '1':
                        arr_ParticipantNumber_of_line[iLineNo-1] = arr_ParticipantNumber_of_line[iLineNo-1] +1
                        sheet_of_lineNo.update_cell(1, 3, arr_ParticipantNumber_of_line[iLineNo-1])
                        iParticipantNumber_of_AllLine = iParticipantNumber_of_AllLine + 1

                        sheet_Checkin.update_cell(value_index + 1, 10, '1') # update status in sheet, status ='1' as already checked in
                        col_CheckInStatus[value_index] = '1'  # update status in array

                        # Send message to rooms
                        sTime = datetime.datetime.now().strftime("%H:%M:%S")
                        sMessage = "{}-{}".format(arr_ParticipantNumber_of_line[iLineNo-1], col_FullName[value_index])
                        SendChatMessage(sMessage, iLineNo)
                        SendChatMessage(sMessage + " Line {} ".format(iLineNo) + sTime, 7)
                        # print("Tổng số người tham dự:: {}".format(iParticipantNumber_of_AllLine))
                    else:
                        sTime = datetime.datetime.now().strftime("%H:%M:%S")
                        sMessage ="Vé đã được sử dụng - " + col_FullName[value_index]
                        SendChatMessage(sMessage + " - " + sTime, iLineNo)
                        SendChatMessage(sMessage + " - Line {} - ".format(iLineNo) + sTime, 8)
                else:
                    sTime = datetime.datetime.now().strftime("%H:%M:%S")
                    sMessage = "Vé không hợp lệ. Mã vé không tồn tại trong cơ sở dữ liệu"
                    SendChatMessage(sMessage + " - " + sTime, iLineNo)
                    SendChatMessage(sMessage + " - QRCode {} - Line {} - ".format(sQRCode, iLineNo) + sTime, 8)

                sheet_of_lineNo.delete_rows(2, 2)

                # recover dead thread
                for index, thread in enumerate(threads):  # index = 0->3
                    if not thread.is_alive():
                        print("Thread {} is dead".format(index))
                        print("Remove and add new Thread as we can't restart a dead thread")
                        new_thread = threading.Thread(target=dhgt_checkin, name="Thread{}".format(index),
                                                      args=(index,))
                        threads.remove(thread)
                        threads.insert(i, new_thread)
                # time.sleep(1)
           else:
                print("no data to check for line {}".format(iLineNo))
                time.sleep(1.5)

threads = list()
def main() :
    for i in range(1, 5, 1):
        thread = threading.Thread(target=dhgt_checkin, name="Thread{}".format(i), args=(i,))
        thread.start()
        threads.append(thread)


if __name__ == '__main__':
    main()