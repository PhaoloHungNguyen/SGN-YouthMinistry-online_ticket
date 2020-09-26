import time

#Below package needed to access Google Sheet
import gspread
from oauth2client.service_account import ServiceAccountCredentials

from selenium import webdriver

def UpdateStatus_and_Attendant(iPosition):
    iNumberAttendantCheckedIn = int(sheetLine1.row_values(1)[2])
    rowSheetRegisteredUser = sheetRegisteredUser.row_values(iPosition)

    if rowSheetRegisteredUser[7] == "N":               #check-in
        sheetRegisteredUser.update_cell(iPosition, 8, "Y")
        iNumberAttendantCheckedIn = iNumberAttendantCheckedIn +1
        # driver.find_element_by_css_selector("span[id = 'welcome']").send_keys("Welcome " + rowSheetRegisteredUser[1] + "to Carlo Festival")
        driver.find_element_by_css_selector("input[id = 'welcome']").clear()
        driver.find_element_by_css_selector("input[id = 'welcome']").send_keys(rowSheetRegisteredUser[1])
    else:     # check-out
        sheetRegisteredUser.update_cell(iPosition, 8, "N")
        iNumberAttendantCheckedIn = iNumberAttendantCheckedIn - 1
        driver.find_element_by_css_selector("input[id = 'checkout']").clear()
        driver.find_element_by_css_selector("input[id = 'checkout']").send_keys(rowSheetRegisteredUser[1])
    sheetLine1.update_cell(1, 3, iNumberAttendantCheckedIn)
    # driver.find_element_by_css_selector("span[id = 'participant']").send_keys("Participant: " + str(iNumberAttendantCheckedIn))
    driver.find_element_by_css_selector("input[id = 'participant']").clear()
    driver.find_element_by_css_selector("input[id = 'participant']").send_keys(str(iNumberAttendantCheckedIn))

#browser exposes an executable file
#Through Selenium test we need to invoke the executable file which will then invoke actual browser
driver=webdriver.Firefox(executable_path="/home/paulnguyen/PycharmProjects/geckodriver_v0_27_0_ForFireFox_MinV60/geckodriver")
#driver = webdriver.Ie(executable_path="C:\\IEDriverServer.exe")
# driver = webdriver.Chrome(executable_path="/home/paulnguyen/PycharmProjects/chromedriver_ChromeV85/chromedriver")

driver.get("http://carloauticus.com/index_line1.html")

# use creds to create a client to interact with the Google Drive API
scope = [
    'https://www.googleapis.com/auth/drive',
    'https://www.googleapis.com/auth/drive.file'
    ]
creds = ServiceAccountCredentials.from_json_keyfile_name('client_secret.json', scope)
client = gspread.authorize(creds)

# Find a workbook by name and open the first sheet
# Make sure you use the right name here.
sheetLine1 = client.open("Line_1").sheet1
sheetRegisteredUser = client.open("3-List of sent_email user").sheet1
colRegisteredUser_Code = sheetRegisteredUser.col_values(7)

while 1>0:
    rowLine1 = sheetLine1.row_values(2)
    try:
        print(rowLine1[1])
        if rowLine1[1] in colRegisteredUser_Code:
            print("Vé Hợp Lệ")
            iRowPosition = 1
            for sCode in colRegisteredUser_Code:
                if sCode == rowLine1[1]:
                    UpdateStatus_and_Attendant(iRowPosition)
                    break
                iRowPosition = iRowPosition + 1

        else:
            print("Lưu ý: Vé không hợp lệ")
            driver.find_element_by_css_selector("input[id = 'welcome']").clear()
            driver.find_element_by_css_selector("input[id = 'welcome']").send_keys("???????")
        sheetLine1.delete_rows(2, 2)
    except Exception as e:
        print("no data to check")

    time.sleep(3)



"""
row = sheet.row_values(4)
print('\nFetched Row')
print(row)

cell = sheet.cell(2,2)
print('\nFetched Cell')
print(cell.value)

col = sheet.col_values(2)
print(col)
for s in  col:
    if "42" in s:
        print(s)
        break

print(col[3])



#sheet.update_cell(3, 5, "Yeah! Great!")

row_number = len(col)
print("Row Number:")
print(row_number)
sheet.delete_rows(2,2)

"""

"""
driver.find_element_by_name("name").send_keys("Paul Nguyen")
driver.find_element_by_id("IdOfElement").click()
driver.find_element_by_css_selector("input[name='name']")
print(driver.find_elements_by_class_name("alert-success").text)
driver.find_element_by_link_text("Forgot your password?").click()
"""

# Generate XPath:
# #//div[@id='usernamegroup']/label     => parent is div, and child is label
# //form[@id='login_form']/div[1]    => In case there are many locator found, [1] to select 1st locator.
# //form[@id='login_form']/div[1]/label

# Generate CSS:
# div[id='usernamegroup'] label
# form[id='login_form'] div:nth-child(1)  => In case there are many locator found, [1] to select 1st locator.
# form[name='login'] label:nth-child(1)

# dropdown = Select(driver.find_element_by_id("id_of_select"))
# dropdown.select_by_visible_text("Female")
# dropdown.select_by_index(0)    ==> first item in dropdown list

# message = driver.find_element_by_classname("alert-success").text
# assert "success" in message

# Handle AutoSuggestive Dynamic Dropdown List
# driver.find_element_by_id("autosuggestive").send_key("RoomName")
# time.sleep(3)
# rooms = driver.find_element_by_css_selector("li[class='ui-menu-item'] a")span[id='participant']
# for room in rooms:
#    if room.text == "Room 8888":
#       room.click()
#       break
# print(driver.find_element_by_id("autosuggestive").text = "Room 8888")
# assert driver.find_element_by_id("autosuggestive").get_attribute('value') == "Room 8888"
