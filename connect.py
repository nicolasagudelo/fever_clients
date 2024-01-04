# Author: Nicolas Agudelo.

import os.path
from  selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
import datetime
from pathlib import Path
from os import rename, mkdir, remove
from shutil import rmtree

downloads_path = str(Path.home() / "Downloads")

today = datetime.date.today()

username = "nmc"

try:
    driver = webdriver.Edge('msedgedriver')
except:
    input("The webdriver needed to run the script is outdated, please download the last stable version from: \nhttps://developer.microsoft.com/en-us/microsoft-edge/tools/webdriver/\nReplace the webdriver on the script folder and run it once again.\nPress Enter to exit the script.")
weeks_list = []

def main(working_directory):

    ##################################
    ############# LOGIN ##############
    ##################################

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

    try:
        rmtree('{desktop_path}/GeneratedCSVs'.format(desktop_path = working_directory))
    except OSError as error:
        pass
        # print(error)

    try:
        mkdir('{desktop_path}/GeneratedCSVs'.format(desktop_path = working_directory))
    except OSError as error:
        pass
        # print(error)

    for index in range (4, 0, -1):
        driver.get('http://avjaspsrvr/jasperserver/flow.html?_flowId=viewReportFlow&reportUnit=/Aldeavision_Reports_Definition/JIRA_Reports/NetworkMaintenance/MonthlyServiceAvailability&standAlone=true&ParentFolderUri=/Aldeavision_Reports_Definition/JIRA_Reports/NetworkMaintenance')

        last_monday = today - datetime.timedelta(days=today.weekday(), weeks = week + 1)
        this_sunday = today - datetime.timedelta(days=today.weekday() + 1, weeks = week)

        weeks_list.append('{last_monday} - {this_sunday}'.format(last_monday = last_monday.strftime('%d-%b'), this_sunday = this_sunday.strftime('%d-%b')))

        driver.find_element('id', 'start_date').send_keys(last_monday.strftime('%m-%d-%Y'))
        driver.find_element('id', 'end_date').send_keys(this_sunday.strftime('%m-%d-%Y'))
        buttons = driver.find_elements(By.CLASS_NAME, "insidebutton")
        buttons[1].click()

        if os.path.exists('{downloads_path}/MonthlyServiceAvailability.csv'.format(downloads_path = downloads_path)):
                   remove('{downloads_path}/MonthlyServiceAvailability.csv'.format(downloads_path = downloads_path))

         # sleep(10)
        while True:
            try:
                img_tag_elements = driver.find_elements(By.TAG_NAME, 'img')
                img_tag_elements[9].click()
                break
            except:
                continue

        while True:
            
            try:
                rename("{downloads_path}/MonthlyServiceAvailability.csv".format(downloads_path = downloads_path), '{desktop_folder}/GeneratedCSVs/{index}. Report from {monday} to {sunday}.csv'.format(desktop_folder = working_directory, monday = last_monday, sunday = this_sunday, index = index - 1))
                break
            # If Source is a file 
            # but destination is a directory
            except IsADirectoryError:
                print("Source is a file but destination is a directory.")
            
            # If source is a directory
            # but destination is a file
            except NotADirectoryError:
                print("Source is a directory but destination is a file.")
            
            # For permission related errors
            except PermissionError:
                print("Operation not permitted.")
            
            # For other errors
            except OSError as error:
                # print(error)
                continue
        week += 1
    driver.quit()
    return weeks_list