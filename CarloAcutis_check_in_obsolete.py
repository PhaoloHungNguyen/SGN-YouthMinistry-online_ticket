# https://realpython.com/intro-to-python-threading/
# https://www.w3schools.com/python/gloss_python_global_variables.asp
# https://wiki.videolan.org/Python_bindings
import threading
import time
import vlc
from conf import conf

#Below package needed to access Google Sheet
import gspread
from oauth2client.service_account import ServiceAccountCredentials

from selenium import webdriver

#browser exposes an executable file
#Through Selenium test we need to invoke the executable file which will then invoke actual browser
driver=webdriver.Firefox(executable_path="/home/paulnguyen/PycharmProjects/geckodriver_v0_27_0_ForFireFox_MinV60/geckodriver")
#driver = webdriver.Ie(executable_path="C:\\IEDriverServer.exe")
# driver = webdriver.Chrome(executable_path="/home/paulnguyen/PycharmProjects/chromedriver_ChromeV85/chromedriver")

driver.get("http://carloauticus.com/index.html")
# driver.fullscreen_window()

# use creds to create a client to interact with the Google Drive API
scope = [
    'https://www.googleapis.com/auth/drive',
    'https://www.googleapis.com/auth/drive.file'
    ]
# Different credentials from different users to avoid Google Limitation of Read GSheet.
# creds = ServiceAccountCredentials.from_json_keyfile_name('./template/client_secret1.json', scope)
# client1 = gspread.authorize(creds)
# creds = ServiceAccountCredentials.from_json_keyfile_name('./template/client_secret2.json', scope)
# client2 = gspread.authorize(creds)
# creds = ServiceAccountCredentials.from_json_keyfile_name('./template/client_secret3.json', scope)
# client3 = gspread.authorize(creds)
# creds = ServiceAccountCredentials.from_json_keyfile_name('./template/client_secret4.json', scope)
# client4 = gspread.authorize(creds)

# Find a workbook by name and open the first sheet
# Make sure you use the right name here.
# To increase more lines: create sheet Line_n, share it, add code sheetLine_n = client.open("Line_n").sheet1
sheet_lines = list()
no_participant_of_lines = list()
sheet_checkin_user_for_lines = list()
no_participant_all_line = 0

for i in range(1,5,1): #[1,4)
    creds = ServiceAccountCredentials.from_json_keyfile_name("./template/client_secret{}.json".format(i), scope)
    client = gspread.authorize(creds)
    sheet_checkin_user_for_lines.append(client.open("2-Check-in").sheet1)
    sheet_lines.append(client.open("Line_{}".format(i)).sheet1)
    no_participant_of_lines.append(int(sheet_lines[i-1].row_values(1)[2]))
    no_participant_all_line = no_participant_all_line + no_participant_of_lines[i-1]
# sheet_lines.append(client1.open("Line_1").sheet1)
# sheet_lines.append(client2.open("Line_2").sheet1)
# sheet_lines.append(client3.open("Line_3").sheet1)
# sheet_lines.append(client4.open("Line_4").sheet1)
#
# sheetLine1 = client1.open("Line_1").sheet1
# sheetLine2 = client2.open("Line_2").sheet1
# sheetLine3 = client3.open("Line_3").sheet1
# sheetLine4 = client4.open("Line_4").sheet1
#
# no_participant_of_line1 = int(sheetLine1.row_values(1)[2])
# no_participant_of_line2 = int(sheetLine2.row_values(1)[2])
# no_participant_of_line3 = int(sheetLine3.row_values(1)[2])
# no_participant_of_line4 = int(sheetLine4.row_values(1)[2])
# no_participant_all_line = no_participant_of_line1 + no_participant_of_line2 + no_participant_of_line3 + no_participant_of_line4

# sheetCheckInUser = client1.open("2-Check-in").sheet1
col_QRcodes = sheet_checkin_user_for_lines[0].col_values(7)
time.sleep(1)
col_CheckInStatus = sheet_checkin_user_for_lines[0].col_values(9)
time.sleep(1)
col_FullName = sheet_checkin_user_for_lines[0].col_values(2)
time.sleep(1)
music_playing = "N"
def welcome_vip():
    global music_playing
    if music_playing == "N":
        music_playing = "Y"
        music = vlc.MediaPlayer("file:///home/paulnguyen/PycharmProjects/CarloAcutis/template/dgm.mp3")
        music.play()
        time.sleep(60)
        music.stop()
        music_playing = "N"


def update_participant_and_screen(iPosition, ticket_status, iLineNo):
    global no_participant_all_line

    if ticket_status == 'N':
        sheet_checkin_user_for_lines[iLineNo-1].update_cell(iPosition, 9, 'Y') # "Y" means checked in
        no_participant_all_line = no_participant_all_line + 1
        sTemp = driver.find_element_by_css_selector("input[id = 'line{}_1']".format(iLineNo)).get_attribute("value")
        driver.find_element_by_css_selector("input[id = 'line{}_2']".format(iLineNo)).clear()
        driver.find_element_by_css_selector("input[id = 'line{}_2']".format(iLineNo)).send_keys(sTemp)
        driver.find_element_by_css_selector("input[id = 'line{}_1']".format(iLineNo)).clear()
        driver.find_element_by_css_selector("input[id = 'line{}_1']".format(iLineNo)).send_keys("{}. {}".format(iLineNo, col_FullName[iPosition-1]))
        driver.find_element_by_css_selector("input[id = 'participant']").clear()
        driver.find_element_by_css_selector("input[id = 'participant']").send_keys(no_participant_all_line)

        # if col_QRcodes[iPosition-1] in conf.gv_adkar_qrcodes:
        #     driver.execute_script("window.open('');")
        #     driver.switch_to.window(driver.window_handles[1])
        #     driver.get("http://carloauticus.com/gv_adkar.html")
        #     music = vlc.MediaPlayer("file:///home/paulnguyen/PycharmProjects/CarloAcutis/template/gv_adkar.mp3")
        #     music.play()
        #     time.sleep(20)
        #     music.stop()
        #     driver.close()
        #     driver.switch_to.window(driver.window_handles[0])
        # else:

        # if col_QRcodes[iPosition-1] in conf.dgm_qrcodes:
        #     thread_welcome = threading.Thread(target=welcome_vip, name="Thread_Welcome", args=( ))
        #     thread_welcome.start()
        #     thread_welcome.join()

    else:
        if ticket_status == 'Y':
            # sTemp = driver.find_element_by_css_selector("textarea[id = 'note']").get_attribute("value")
            sTemp = "{}. {} đã scan.\n".format(iLineNo, col_FullName[iPosition-1])   # + sTemp.split('\n')[0]
            # driver.find_element_by_css_selector("textarea[id = 'note']").clear()
            driver.find_element_by_css_selector("textarea[id = 'note']").send_keys(sTemp)
        else:   # ticket_status == 'W' == Fake ticket
            # sTemp = driver.find_element_by_css_selector("textarea[id = 'note']").text
            sTemp = "{}. Vé không hợp lệ.\n".format(iLineNo)    # + sTemp.split('\n')[0]
            # driver.find_element_by_css_selector("textarea[id = 'note']").clear()
            driver.find_element_by_css_selector("textarea[id = 'note']").send_keys(sTemp)

def check_in_line(iLineNo):
    global threads
    # switcher ={
    #     1: sheetLine1,
    #     2: sheetLine2,
    #     3: sheetLine3,
    #     4: sheetLine4
    # }
    sheet_of_lineNo = sheet_lines[iLineNo-1]
        # switcher.get(iLineNo, [])

    while True:
        try:
            row_to_check = sheet_of_lineNo.row_values(2)
        except Exception as e:
            print("Important Note: Can't read Gsheet Line{},\nMaybe caused by internet connection issue".format(iLineNo))
        else:
           if row_to_check:
                print(row_to_check)
                try:
                    arrRow = row_to_check[1].split("\n")
                    sQRCode = arrRow[2]
                except Exception as e:
                    # print("Vé không hợp lệ")
                    sQRCode = ""

                status_of_checking_ticket = "N"
                try:
                    value_index = col_QRcodes.index(sQRCode)    # index of string sQRCode in string list col_QRCodes
                except Exception as e:
                    value_index = -1
                    status_of_checking_ticket = "W"    # fake ticket
                    print("Vé không hợp lệ")

                if value_index > 0:
                    print("Vé Hợp Lệ")
                    # print("value_index = {}".format(value_index))
                    status_of_checking_ticket = col_CheckInStatus[value_index]  # get current status of checking_ticket
                    if status_of_checking_ticket == 'N':
                        no_participant_of_lines[iLineNo-1] = no_participant_of_lines[iLineNo-1] +1
                        sheet_of_lineNo.update_cell(1, 3, no_participant_of_lines[iLineNo-1])
                    col_CheckInStatus[value_index] = 'Y'    # update status in array first, then update status in sheet through function update_participant_and_screen

                iPosiion = value_index + 1
                update_participant_and_screen(iPosiion, status_of_checking_ticket, iLineNo)
                sheet_of_lineNo.delete_rows(2, 2)

                # recover dead thread
                for index, thread in enumerate(threads):  # index = 0->3
                    if not thread.is_alive():
                        print("Thread {} is dead".format(index))
                        print("Remove and add new Thread as we can't restart a dead thread")
                        new_thread = threading.Thread(target=check_in_line, name="Thread{}".format(index),
                                                      args=(index,))
                        threads.remove(thread)
                        threads.insert(i, new_thread)

           else:
                print("no data to check")
                time.sleep(1.5)
            # check_in_line(iLineNo, no_participant_of_line)

threads = list()
for i in range(1, 5, 1):
    # switcher = {
    #     1: no_participant_of_line1,
    #     2: no_participant_of_line2,
    #     3: no_participant_of_line3,
    #     4: no_participant_of_line4
    # }
    # participant = no_participant_of_lines[i-1]
        # switcher.get(i, 0)

    thread = threading.Thread(target=check_in_line, name="Thread{}".format(i), args=(i, ))
    thread.start()
    threads.append(thread)

# for index, thread in enumerate(threads):  # index = 0->3
#     print("index ={}, threadName = {}".format(index, thread.getName()))
#     if not thread.is_alive():
#         print("index ={}, threadName = {} is not alive".format(index, thread.getName()))


