from tkinter import *
import os.path
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time
from threading import *
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import csv


# Load Cookies Manually
okta_cookies = [
    {"name": "DT", "value": ""},
    {"name": "proximity_00ab843eaa804e450d5881391302ad92", "value": ""},
    {"name": "ln", "value": ""},
    {"name": "enduser_version", "value": ""},
    {"name": "t", "value": ""},
    {"name": "okta_user_lang", "value": ""},
    {"name": "autolaunch_triggered", "value": ""},
    {"name": "srefresh", "value": ""},
    {"name": "okta-oauth-nonce", "value": ""},
    {"name": "okta-oauth-state", "value": ""},
    {"name": "sid", "value": ""},
]

# to run on the UAT, put https://wayfairtest.service-now.com/sc_task.do?sysparm_query=number=
# to run on Dev put https://wayfairdev.service-now.com/sc_task.do?sysparm_query=number=
# to run on regular just have it wayfair

based_url = "https://wayfairdev.service-now.com/sc_task.do?sysparm_query=number="
test_task_link = based_url+"SCTASK0379642"
# CSV File
csv_location = "UPS Batch Shipping Template V2.0 - 4future (2).csv"

# options=webdriver.ChromeOptions()
#options.add_argument("user-data-dir=C:\\Users\\as112c\\AppData\\Local\\Google\\Chrome\\User Data")
#options.add_argument("profile-directory=Profile 2")

stop = FALSE


def open_browser():
    global browser
    if not browser:
        # CHANGE THIS FILEPATH IN YOUR VERSION
        filepath = os.path.dirname(__file__)
        filepath = filepath + '\\chromedriver.exe'
        browser = webdriver.Chrome(filepath)
        browser.get(test_task_link)
    return


def load_okta():
    global browser
    global okta_cookies
    for cookie in okta_cookies:
        browser.add_cookie(cookie)
    browser.refresh()
    return


def open_task():
    global browser
    browser.get(test_task_link)


def stop_process():
    global stop
    stop = TRUE


def start_process():
    global browser
    global stop

    stop = FALSE
    with open(csv_location, newline='') as csvfile:
        reader = csv.DictReader(csvfile)
        error_assets = []
        for row in reader:
                task_number = row['Reference #1 (Tasks)']
                #mailing_address = row['']
                tracking_number = row['Tracking #']
                asset_tag = row['Asset Tag']
                comment = row['AddComment']
                if task_number == "SCTASK0398778":
                    print("continue1")
                    continue
                if stop == FALSE:
                    print('Processing :', task_number, tracking_number, asset_tag)
                    browser.get(based_url+task_number)

                try:
                    wait_element = WebDriverWait(browser, 5).until(EC.presence_of_element_located(
                        (By.XPATH, '//*[@ac_columns="asset_tag;serial_number"]')))
                    print("Page is ready!")
                except TimeoutException:
                    print("Loading took too much time!")
                    continue
                except Exception:
                    print("There was another error!")
                    continue

                asset_tag_element = browser.find_element(
                    by=By.XPATH, value='//*[@ac_columns="asset_tag;serial_number"]')
                tracking_number_element = browser.find_element(
                    by=By.XPATH, value='/html/body/div[2]/form/span[1]/span/div[5]/div[4]/div/div[5]/table/tbody/tr[2]/td/div[2]/div/table/tbody/tr[6]/td/span/div[2]/div[2]/table/tbody/tr[2]/td/div/div/div/div[2]/input[2]')
                comment_element = browser.find_element(
                    by=By.XPATH, value='//*[@id="activity-stream-comments-textarea"]')
                #mailing_address_element = browser.find_element(by=By.XPATH, value='//*[@id="ni.VEa1ce566bdb0541108cc2e439d39619a3"]')
                
                #if mailing_address_element.get_attribute('value') == '':
                #    mailing_address_element.send_keys(mailing_address)
                
                
                if asset_tag_element.get_attribute('value') == '':
                    asset_tag_element.send_keys(asset_tag)
                    time.sleep(1)
                    asset_tag_element.send_keys(Keys.TAB)

                if tracking_number_element.get_attribute('value') == '':
                    tracking_number_element.send_keys(tracking_number)

                comment_element.send_keys(comment)
                time.sleep(1)
                comment_element.send_keys(Keys.TAB)

                if "ref_invalid" not in asset_tag_element.get_attribute('class'):
                    save_btn = browser.find_element(
                        by=By.ID, value="sysverb_update_and_stay")
                    save_btn.click()

                else:
                    asset_tag_element.clear()
                    save_btn.click()
                    error_assets.append([task_number, asset_tag])
                    print("passed")
                    continue

                try:
                    wait_element = WebDriverWait(browser, 5).until(EC.presence_of_element_located(
                        (By.XPATH, '//button[normalize-space()="Create UPS Tracking Record"]')))
                    print("UPS BTN Present")
                    ups_btn = browser.find_element(
                        by=By.XPATH, value='//button[normalize-space()="Create UPS Tracking Record"]')
                    ups_btn.click()
                except TimeoutException:
                    print("Error UPS BTN")
                    error_assets.append([task_number, asset_tag])
                    continue

                else:
                    print('"Process stop for:', task_number,
                        tracking_number, asset_tag)

    #######################################################
    #######################################################
    #######################################################
    #######################################################
    #######################################################


browser = None

window = Tk()
window.title("Program")
window.geometry("700x500")

theTekst = Label(window, text="Link:")
theTekst.pack()

input_link = Entry(window, width=110)
input_link.pack()

btn_open_browser = Button(window, text="Open Browser", command=open_browser)
btn_open_browser.pack()

btn_load_okta = Button(window, text="Load Okta", command=load_okta)
btn_load_okta.pack()

btn_open_task = Button(window, text="Open Task", command=open_task)
btn_open_task.pack()

btn_start_process = Button(window, text="Start Process", command=start_process)
btn_start_process.pack()

btn_stop_process = Button(window, text="Stop Process", command=stop_process)
btn_stop_process.pack()


window.mainloop()
