import time, sys, unittest, random, json, os, pyperclip, framework_sample
from datetime import datetime
from selenium import webdriver
from random import randint, choice
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import TimeoutException, WebDriverException, NoSuchElementException
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.alert import Alert
from pathlib import Path
from MN_functions import *
from framework_sample import *
current_path = os.path.dirname(Path(__file__).absolute())
json_file = current_path + "\\MN_groupware_auto.json"

# Start the web driver
service = webdriver.chrome.service.Service(current_path + "\\chromedriver_talk.exe")
service.start()

#data
attachment = current_path + "background6.jpg"
file_text = current_path + "file_text.txt"
dept_org = "Selenium"    
contact_org = "AutomationTest"
password_talk = "automationtest1!"
forward_name = "AutomationTest2"
add_user = "AutomationTest"
chat_content = "Hanbiro test " + objects.date_time
quote_chat = "This is quote content: " + objects.date_time
content_whisper = "This is message of whisper, date: " + objects.date_time
content_board = "This is content of Board: " + objects.date_time
content_edit = "Edit content Board: " + objects.date_time

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
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "//input[@id='domain']")))

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
    user_id.send_keys(contact_org.lower())
    time.sleep(1)
    driver.find_element_by_xpath("//input[@id='password']").send_keys(password_talk)
    time.sleep(1)
    driver.find_element_by_xpath("//span[text()='Sign In']").click()
    try:
        access_page = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "//div[@id='root']//ul/div")))
        print(bcolors.OKGREEN + "=> Login success" + bcolors.ENDC)
        TestCase_LogResult(**data["testcase_result"]["talk2"]["login"]["pass"])
    except:
        print(bcolors.OKGREEN + "=> Login fail" + bcolors.ENDC)
        TestCase_LogResult(**data["testcase_result"]["talk2"]["login"]["fail"])

def get_newest_mess():
    newest_mess = driver.find_element_by_xpath("//div[@class='simplebar-content']//div[@aria-label='grid']/div/div[1]//p/span")
    newest_mess_content= newest_mess.text
    newest_mess.click()
    print("- Get newest mess")
    time.sleep(8)
    get_text = driver.find_element_by_xpath("//div[@class='simplebar-content']//div[contains(@class,'hanbiroToFadeInAndOut')]/div[last()]//div[contains(.,'"+ str(newest_mess_content) +"')]")
    if get_text.is_displayed():
        print("=> Newest mess correct content")

def message():
    #access message tab
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH,"//ul[contains(@class,'MuiList-padding')]//div[contains(@aria-label,'Room list')]"))).click()
    time.sleep(5)
    try:
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "//div[@class='simplebar-content']//div[@aria-label='grid']//button[@class='MuiButtonBase-root']")))
        print("=> Access messenger tab")
        get_newest_mess()
    except:
        print("=> Cannot access messenger tab")
        pass

    try:
        searchuser()
    except:
        pass

def searchuser():
    driver.find_element_by_xpath("//ul[contains(@class,'MuiList-padding')]//div[1]").click()
    print("- Access ORG")
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "//div[@class='simplebar-content']//div[@aria-label='grid']/div/div")))
    time.sleep(3)
    search_contact = driver.find_element_by_xpath("//*[@id='root']//div[contains(@class,'MuiInputBase-root')]/input[contains(@class,'MuiInputBase-input')]")
    search_contact.send_keys(contact_org)
    search_contact.send_keys(Keys.ENTER)
    time.sleep(2)
    print("- Search myself")
    try:
        contact_search = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "//*[@class='simplebar-mask']//div[@class='simplebar-content']//span[contains(.,'Contacts')]/following-sibling::div//div[contains(.,'"+ str(contact_org) +"')]")))
        contact_search.click()
        print("=> Search user success")
        send_mess()
    except:
        print("=> search user fail")

def send_mess():
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
    ####################################################################################### ATTACH PC fail
    #attach PC
    attach_file = driver.find_element_by_xpath("//*[@id='hanbiro_message_list_chat_input']//div[2]/input")
    attach_file.send_keys(attachment)
    print(bcolors.OKGREEN + "- Attach file PC" + bcolors.ENDC)
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
#whisper()

