# Author: Nicolas Agudelo.

from  selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from time import sleep
import datetime
from pathlib import Path

downloads_path = str(Path.home() / "Downloads")

print(downloads_path)

today = datetime.date.today()

##################################
############# LOGIN ##############
##################################
username = "nmc"

driver = webdriver.Edge('msedgedriver')

driver.get('http://avjaspsrvr/jasperserver/login.html?showPasswordChange=null/')

driver.find_element('name', 'j_username').send_keys(username)
driver.find_element('name', 'j_password').send_keys(username)
driver.find_element('name','btnsubmit').click()

# wait the ready state to be complete
WebDriverWait(driver=driver, timeout=10).until(
    lambda x: x.execute_script("return document.readyState === 'complete'")
)
error_message = "Incorrect username or password."
# get the errors (if there are)
errors = driver.find_elements("css selector", ".flash-error")
# print the errors optionally
# for e in errors:
#     print(e.text)
# if we find that error message within errors, then login is failed
if any(error_message in e.text for e in errors):
    print("[!] Login failed")
else:
    print("[+] Login successful")

#####################################
######## GENERATE REPORTS ###########
#####################################

week = 0

for index in range (0, 5):
    driver.get('http://avjaspsrvr/jasperserver/flow.html?_flowId=viewReportFlow&reportUnit=/Aldeavision_Reports_Definition/JIRA_Reports/NetworkMaintenance/MonthlyServiceAvailability&standAlone=true&ParentFolderUri=/Aldeavision_Reports_Definition/JIRA_Reports/NetworkMaintenance')

    last_monday = today - datetime.timedelta(days=today.weekday(), weeks = week + 1)
    this_sunday = today - datetime.timedelta(days=today.weekday() + 1, weeks = week)

    driver.find_element('id', 'start_date').send_keys(last_monday.strftime('%m-%d-%Y'))
    driver.find_element('id', 'end_date').send_keys(this_sunday.strftime('%m-%d-%Y'))
    buttons = driver.find_elements(By.CLASS_NAME, "insidebutton")
    buttons[1].click()

    sleep(10)

    img_tag_elements = driver.find_elements(By.TAG_NAME, 'img')
    img_tag_elements[9].click()
    week += 1