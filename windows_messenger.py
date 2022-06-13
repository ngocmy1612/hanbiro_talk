import time, sys, unittest, random, json
from datetime import datetime
from selenium import webdriver
from random import randint
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import TimeoutException, WebDriverException, NoSuchElementException
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.alert import Alert
from random import choice
from pathlib import Path
import os
import pyperclip
from framework_sample import *

from MN_functions import driver, data, ValidateFailResultAndSystem, Logging, TestCase_LogResult
json_file = os.path.dirname(Path(__file__).absolute())+"\\MN_groupware_auto.json"

# Start the web driver
service = webdriver.chrome.service.Service("C:\\Users\\Ngoc\\Desktop\\ngoc_automationtest\\auto_hanbiro_talk\\chromedriver_talk.exe")
service.start()

class bcolors:
    HEADER = '\033[95m'
    OKBLUE = '\033[94m'
    OKGREEN = '\033[92m'
    WARNING = '\033[93m'
    FAIL = '\033[91m'
    ENDC = '\033[0m'
    BOLD = '\033[1m'
    UNDERLINE = '\033[4m'

attachment = "C:\\Users\\Ngoc\\Desktop\\ngoc_automationtest\\auto_hanbiro_talk\\attachment\\background6.jpg"
file_text = "C:\\Users\\Ngoc\\Desktop\\ngoc_automationtest\\auto_hanbiro_talk\\attachment\\file_text.txt"

n = random.randint(1,1000)
now = datetime.now()
date = now.strftime("%m/%d/%y %H:%M:%S")

dept_org = "Selenium"    
contact_org = "AutomationTest"
forward_name = "AutomationTest2"
add_user = "AutomationTest"
chat_content = "This is content chat, date: " + date
quote_chat = "This is quote content: " + date
content_whisper = "This is message of whisper, date: " + date
content_board = "This is content of Board: " + date
content_edit = "Edit content Board: " + date

# start the app
driver = webdriver.remote.webdriver.WebDriver(
    command_executor=service.service_url,
    desired_capabilities={
        'browserName': 'chrome',
        'goog:chromeOptions': {
            'args': ['develop_mode'],
            'binary': 'C:\\Users\\Ngoc\\AppData\\Local\\Programs\\hanbiro-talk\\HanbiroTalk2.exe',
            'extensions': [],
            'windowTypes': ['webview']},
        'platform': 'ANY',
        'version': ''},
    browser_profile=None,
    proxy=None,
    keep_alive=False)

def login():
    #login
    Waits.Wait20s_ElementLoaded("//input[@id='domain']")
    domain = driver.find_element_by_xpath("//input[@id='domain']")
    if bool(domain.get_attribute("value")) == True:
        domain.clear()
        time.sleep(1)
    
    domain.send_keys("myngoc.hanbiro.net")

    time.sleep(1)

    user_id = driver.find_element_by_xpath("//input[@id='userid']")
    if bool(user_id.get_attribute("value")) == True:
        user_id.clear()
        time.sleep(1)
    user_id.send_keys("automationtest")
    time.sleep(1)
    Commands.InputElement("//input[@id='password']", "automationtest1!")
    time.sleep(1)
    driver.find_element_by_xpath("//span[text()='Sign In']").click()
    try:
        access_page = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "//div[@id='root']//ul/div")))
        print(bcolors.OKGREEN + "=> Login success" + bcolors.ENDC)
        TestCase_LogResult(**data["testcase_result"]["talk2"]["login"]["pass"])
    except:
        print(bcolors.OKGREEN + "=> Login fail" + bcolors.ENDC)
        TestCase_LogResult(**data["testcase_result"]["talk2"]["login"]["fail"])
    
def message():
    #access message tab
    access_page_chat = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "//div[@id='hanbiro_message_list_chat_input']")))
    time.sleep(2)
    
    if access_page_chat.is_displayed():
        print(bcolors.OKGREEN + ">> Access page chat success" + bcolors.ENDC)
        TestCase_LogResult(**data["testcase_result"]["talk2"]["access_message_page"]["pass"])
        try:
            write_content()
        except:
            pass   
        search_user_in_mess()        
    else:
        print(bcolors.OKGREEN + ">> Access page chat fail" + bcolors.ENDC)
        TestCase_LogResult(**data["testcase_result"]["talk2"]["access_message_page"]["fail"])
        pass

def write_content():
    #send message: hanbiro test + current time (send to myself) / send with attachment
    input_content = driver.find_element_by_xpath("//div[@id='textBox']")
    input_content.send_keys(chat_content)
    input_content.send_keys(Keys.ENTER)
    print(bcolors.OKGREEN + "- Input content chat" + bcolors.ENDC)

    result_chat = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "//*[@id='hanbiro_message_list_container']//div[@class='simplebar-content']//div//p[contains(text(),'" + str(chat_content) + "')]")))
    print(bcolors.OKGREEN + ">> Send message success" + bcolors.ENDC)
    TestCase_LogResult(**data["testcase_result"]["talk2"]["write_content"]["pass"])

    #attach clouddisk
    time.sleep(2)
    driver.find_element_by_xpath("//*[@id='hanbiro_message_list_chat_input']//button[2]").click()
    print("- Attach file Clouddisk")
    attach_clouddisk()

    #attach PC
    attach_file = driver.find_element_by_xpath("//*[@id='hanbiro_message_list_chat_input']//div[2]/input")
    attach_file.send_keys(attachment)
    print(bcolors.OKGREEN + "- Attach file" + bcolors.ENDC)
    time.sleep(5)

def attach_clouddisk():
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "//div[contains(@class,'MuiDialog-paperScrollPaper')]//div[contains(@class,'Pane1')]//div[@class='MuiListItemText-root']")))
    time.sleep(2)
    try:
        no_items = driver.find_element_by_xpath("//div[contains(@class,'MuiDialog-paperScrollPaper')]//div[contains(@class,'Pane2')]//div/p[contains(.,'No items')]")
        if no_items.is_displayed():
            print("=> No file in Clouddisk to attach")
            driver.find_element_by_xpath("//button/span[text()='Close']").click()
    except:
        count_file = int(len(driver.find_elements_by_xpath("//div[contains(@class,'MuiDialog-paperScrollPaper')]//div[contains(@class,'Pane2')]//div[@class='simplebar-content']//div/div[@role='rowgroup']/div")))
        if count_file > 2:
            driver.find_element_by_xpath("//div[contains(@class,'MuiDialog-paperScrollPaper')]//div[contains(@class,'Pane2')]//div[@class='simplebar-content']//div[2]/div[contains(@class,'MuiListItem-button')]//span/input").click()
            driver.find_element_by_xpath("//div[contains(@class,'MuiDialog-paperScrollPaper')]//div[contains(@class,'Pane2')]//div[@class='simplebar-content']//div[3]/div[contains(@class,'MuiListItem-button')]//span/input").click()
            print("=> Select file")
            time.sleep(2)
        driver.find_element_by_xpath("//button/span[text()='SEND']").click()
        print("=> Send attach file clouddisk")
        
def attach_clouddisk_whisper():
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "//div[contains(@class,'MuiDialog-paperScrollPaper')]//div[contains(@class,'Pane1')]//div[@class='MuiListItemText-root']")))
    time.sleep(2)
    try:
        no_items = driver.find_element_by_xpath("//div[contains(@class,'MuiDialog-paperScrollPaper')]//div[contains(@class,'Pane2')]//div/p[contains(.,'No items')]")
        if no_items.is_displayed():
            print("=> No file in Clouddisk to attach")
            driver.find_element_by_xpath("//button/span[text()='Close']").click()
    except:
        count_file = int(len(driver.find_elements_by_xpath("//div[contains(@class,'MuiDialog-paperScrollPaper')]//div[contains(@class,'Pane2')]//div[@class='simplebar-content']//div/div[@role='rowgroup']/div")))
        if count_file > 2:
            driver.find_element_by_xpath("//div[contains(@class,'MuiDialog-paperScrollPaper')]//div[contains(@class,'Pane2')]//div[@class='simplebar-content']//div[2]/div[contains(@class,'MuiListItem-button')]//span/input").click()
            driver.find_element_by_xpath("//div[contains(@class,'MuiDialog-paperScrollPaper')]//div[contains(@class,'Pane2')]//div[@class='simplebar-content']//div[3]/div[contains(@class,'MuiListItem-button')]//span/input").click()
            print("=> Select file")
            time.sleep(2)
        driver.find_element_by_xpath("//button/span[text()='Apply']").click()
        print("=> Send attach file clouddisk")

def search_user_in_mess():
    #Search user in message tab
    driver.find_element_by_xpath("//ul[contains(@class,'MuiList-padding')]//div[3]").click()
    time.sleep(3)
    search_contact = driver.find_element_by_xpath("//*[@id='root']//div[contains(@class,'MuiInputBase-root')]/input[contains(@class,'MuiInputBase-input')]")
    search_contact.send_keys(contact_org)
    search_contact.send_keys(Keys.ENTER)
    print(bcolors.OKGREEN + "- Search User" + bcolors.ENDC)
    time.sleep(2)

    contact_search = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH,"//*[@class='simplebar-mask']//div[@class='simplebar-content']//span[contains(.,'Contacts')]/following-sibling::div//div[contains(.,'"+ str(contact_org) +"')]")))
    if contact_search.is_displayed():
        print(bcolors.OKGREEN + ">> Search contact success" + bcolors.ENDC)
        contact_search.click()
        TestCase_LogResult(**data["testcase_result"]["talk2"]["message_search"]["pass"])
    else:
        print(bcolors.OKGREEN + ">> Search contact fail" + bcolors.ENDC)
        TestCase_LogResult(**data["testcase_result"]["talk2"]["message_search"]["fail"])
            
def whisper():
    driver.find_element_by_xpath("//ul[contains(@class,'MuiList-padding')]//div[1]").click()
    WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.XPATH, "//*[@class='simplebar-mask']//div[@class='simplebar-content']/div/div/div")))
    time.sleep(5)

    search_contact = driver.find_element_by_xpath("//*[@id='root']//div[contains(@class,'MuiInputBase-root')]/input[contains(@class,'MuiInputBase-input')]")
    search_contact.send_keys(contact_org)
    search_contact.send_keys(Keys.ENTER)
    time.sleep(2)

    contact_search = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "//*[@class='simplebar-mask']//div[@class='simplebar-content']//span[contains(.,'Contacts')]/following-sibling::div//div[contains(.,'"+ str(contact_org) +"')]")))
    if contact_search.is_displayed():
        print(bcolors.OKGREEN + ">> Search user success" + bcolors.ENDC)
        time.sleep(3)
        actionChains = ActionChains(driver)
        actionChains.context_click(contact_search).perform()
        print(bcolors.OKGREEN + "- Right click" + bcolors.ENDC)

        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH,"//div[contains(@class,'MuiPaper-rounded')]/ul[contains(@class,'MuiMenu-list')]//li//span[contains(.,'Send Whisper')]"))).click()
        time.sleep(3)
        print(bcolors.OKGREEN + "- Send whisper" + bcolors.ENDC)
        send_whisper()
        TestCase_LogResult(**data["testcase_result"]["talk2"]["access_whisper_page"]["pass"])
    else:
        print(bcolors.OKGREEN + ">> Search user fail" + bcolors.ENDC)
        TestCase_LogResult(**data["testcase_result"]["talk2"]["access_whisper_page"]["fail"])

def send_whisper():
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "//div[contains(@class,'MuiDialog-container')]//div[@class='MuiDialogContent-root']//h6[contains(.,'Whisper Write')]")))
    time.sleep(2)

    input_whisper = driver.find_element_by_xpath("//div[contains(@placeholder,'Enter a message')]")
    input_whisper.send_keys(content_whisper)
    time.sleep(3)
    driver.find_element_by_xpath("//div[contains(@class,'MuiDialog-container')]//div[@class='MuiDialogContent-root']//h6[contains(.,'Whisper Write')]//following::input[2]/../following-sibling::button").click()
    print("- Attach file Clouddisk")
    attach_clouddisk_whisper()
    time.sleep(2)

    attach_whisper = driver.find_element_by_xpath("//div[contains(@class,'MuiDialog-container')]//div[@class='MuiDialogContent-root']//h6[contains(.,'Whisper Write')]//following::input[2]")
    attach_whisper.send_keys(file_text)
    print(bcolors.OKGREEN + "- Attach file whisper" + bcolors.ENDC)
    time.sleep(2)
    driver.find_element_by_xpath("//div[contains(@class,'MuiDialog-container')]//h6[contains(.,'Whisper Write')]//following::button/span[contains(.,'SEND')]").click()
    print(bcolors.OKGREEN + "- Send whisper" + bcolors.ENDC)
    time.sleep(5)
    driver.find_element_by_xpath("//div[contains(@class,'MuiDialog-container')]//h6[contains(.,'Whisper Write')]//following-sibling::button").click()
    print(bcolors.OKGREEN + "- Close pop up write whisper" + bcolors.ENDC)
    time.sleep(3)

    driver.find_element_by_xpath("//ul[contains(@class,'MuiList-padding')]//div[4]").click()
    print(bcolors.OKGREEN + "- Access Whisper page" + bcolors.ENDC)
    time.sleep(3)


login()
message()
whisper()

