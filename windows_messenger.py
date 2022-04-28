import time, sys, unittest, random, json#, requests, testlink
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

from MN_functions import driver, data, ValidateFailResultAndSystem, Logging, TestCase_LogResult
json_file = os.path.dirname(Path(__file__).absolute())+"\\MN_groupware_auto.json"

# Start the web driver
#service = webdriver.chrome.service.Service("D:\\Ngoc\\ngoc_automationtest\\auto_hanbiro_talk\\chromedriver_talk.exe")
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

#data
#attachment = "D:\\Ngoc\\ngoc_automationtest\\auto_hanbiro_talk\\attachment\\background6.jpg"
attachment = "C:\\Users\\Ngoc\\Desktop\\ngoc_automationtest\\auto_hanbiro_talk\\attachment\\background6.jpg"
#file_text = "D:\\Ngoc\\ngoc_automationtest\\auto_hanbiro_talk\\attachment\\file_text.txt"
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
            #'binary': 'C:\\Users\\ADMIN\\AppData\\Local\\Programs\\hanbiro-talk\\HanbiroTalk2.exe',
            #'binary': 'C:\\Users\\Ngoc\\AppData\\Local\\Programs\\hanbiro-talk-test\\HanbiroTalk2-test.exe',
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
    #domain.send_keys("qa.hanbiro.net")
    #domain.send_keys("global3.hanbiro.com")
    #domain.send_keys("gw.hanbiro.vn")

    time.sleep(1)

    user_id = driver.find_element_by_xpath("//input[@id='userid']")
    if bool(user_id.get_attribute("value")) == True:
        user_id.clear()
        time.sleep(1)
    user_id.send_keys("automationtest")
    time.sleep(1)
    driver.find_element_by_xpath("//input[@id='password']").send_keys("automationtest1!")
    time.sleep(1)
    driver.find_element_by_xpath("//span[text()='Sign In']").click()
    try:
        access_page = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "//div[@id='root']//ul/div")))
        print(bcolors.OKGREEN + "=> Login success" + bcolors.ENDC)
        TestCase_LogResult(**data["testcase_result"]["talk2"]["login"]["pass"])
    except:
        print(bcolors.OKGREEN + "=> Login fail" + bcolors.ENDC)
        TestCase_LogResult(**data["testcase_result"]["talk2"]["login"]["fail"])

def organization():
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "//*[@class='simplebar-mask']//div[@class='simplebar-content']/div/div/div")))
    time.sleep(3)
        
    search_dept = driver.find_element_by_xpath("//*[@id='root']//div[contains(@class,'MuiInputBase-root')]/input[contains(@class,'MuiInputBase-input')]")
    search_dept.send_keys(dept_org)
    search_dept.send_keys(Keys.ENTER)
    print(bcolors.OKGREEN + "- Search Department" + bcolors.ENDC)
    time.sleep(2)

    dept_search = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "//*[@class='simplebar-mask']//div[@class='simplebar-content']//span[contains(.,'Departments')]/following-sibling::div//div[contains(.,'"+ str(dept_org) +"')]")))
    if dept_search.is_displayed():
        print(bcolors.OKGREEN + ">> Search department success" + bcolors.ENDC)
        TestCase_LogResult(**data["testcase_result"]["talk2"]["search_ORG"]["pass"])
    else:
        print(bcolors.OKGREEN + ">> Search department fail" + bcolors.ENDC)
        TestCase_LogResult(**data["testcase_result"]["talk2"]["search_ORG"]["fail"])

    driver.find_element_by_xpath("//*[@id='root']//div[contains(@class,'MuiInputBase-root')]/following-sibling::button/span[@class='MuiIconButton-label']").click()
    print(bcolors.OKGREEN + "- Delete key search" + bcolors.ENDC)
    search_contact = driver.find_element_by_xpath("//*[@id='root']//div[contains(@class,'MuiInputBase-root')]/input[contains(@class,'MuiInputBase-input')]")
    search_contact.send_keys(contact_org)
    search_contact.send_keys(Keys.ENTER)
    print(bcolors.OKGREEN + "- Search User" + bcolors.ENDC)
    time.sleep(2)

    contact_search = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH,"//*[@class='simplebar-mask']//div[@class='simplebar-content']//span[contains(.,'Contacts')]/following-sibling::div//div[contains(.,'"+ str(contact_org) +"')]")))
    if contact_search.is_displayed():
        print(bcolors.OKGREEN + ">> Search contact success" + bcolors.ENDC)
        TestCase_LogResult(**data["testcase_result"]["talk2"]["search_user"]["pass"])
        contact_search.click()
    else:
        print(bcolors.OKGREEN + ">> Search contact fail" + bcolors.ENDC)
        TestCase_LogResult(**data["testcase_result"]["talk2"]["search_user"]["fail"])
    
def message():
    access_page_chat = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "//div[@id='hanbiro_message_list_chat_input']")))
    time.sleep(2)
    
    if access_page_chat.is_displayed():
        print(bcolors.OKGREEN + ">> Access page chat success" + bcolors.ENDC)
        TestCase_LogResult(**data["testcase_result"]["talk2"]["access_message_page"]["pass"])
        try:
            write_content()
        except:
            pass

        try:
            add_new_member()
        except:
            pass            
    else:
        print(bcolors.OKGREEN + ">> Access page chat fail" + bcolors.ENDC)
        TestCase_LogResult(**data["testcase_result"]["talk2"]["access_message_page"]["fail"])
        pass

def write_content():
    input_content = driver.find_element_by_xpath("//div[@id='textBox']")
    input_content.send_keys(chat_content)
    input_content.send_keys(Keys.ENTER)
    print(bcolors.OKGREEN + "- Input content chat" + bcolors.ENDC)

    result_chat = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "//*[@id='hanbiro_message_list_container']//div[@class='simplebar-content']//div//p[contains(text(),'" + str(chat_content) + "')]")))
    print(bcolors.OKGREEN + ">> Send message success" + bcolors.ENDC)
    TestCase_LogResult(**data["testcase_result"]["talk2"]["write_content"]["pass"])
    actionChains = ActionChains(driver)
    time.sleep(3)
    actionChains.context_click(result_chat).perform()
    time.sleep(3)
    print(bcolors.OKGREEN + "- Right click" + bcolors.ENDC)

    quote_content = driver.find_element_by_xpath("//*[@id='c-menu']//li/span[contains(.,'Quote')]")
    quote_content.click()
    print(bcolors.OKGREEN + "- Quote content" + bcolors.ENDC)

    try:
        content_quote_chat = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "//*[@id='hanbiro_message_list_chat_input']//div[contains(text(),'" + str(chat_content) + "')]//following::div[@id='textBox']")))
        content_quote_chat.send_keys(quote_chat)
        content_quote_chat.send_keys(Keys.ENTER)

        result_quote = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "//*[@id='hanbiro_message_list_container']//div[@class='simplebar-content']//div//span[contains(.,'" + str(quote_chat) + "')]")))
        time.sleep(3)
        print(bcolors.OKGREEN + ">> Quote content success" + bcolors.ENDC)
        TestCase_LogResult(**data["testcase_result"]["talk2"]["quote_content"]["pass"])
    except:
        print(bcolors.OKGREEN + ">> Quote content fail" + bcolors.ENDC)
        TestCase_LogResult(**data["testcase_result"]["talk2"]["quote_content"]["fail"])

    time.sleep(2)
    driver.find_element_by_xpath("//*[@id='hanbiro_message_list_chat_input']//button[2]").click()
    print("- Attach file Clouddisk")
    attach_clouddisk()

    attach_file = driver.find_element_by_xpath("//*[@id='hanbiro_message_list_chat_input']//div[2]/input")
    #https://github.com/SeleniumHQ/selenium/pull/7446/files (Fix python file when run on new computer)
    #attach_file.send_keys(file_text)
    attach_file.send_keys(attachment)
    print(bcolors.OKGREEN + "- Attach file" + bcolors.ENDC)
    time.sleep(5)

    element_to_hover_over = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "//*[@id='hanbiro_message_list_container']//div[@class='simplebar-content']/div/div/div[last()]//span//img")))
    image_height = int(element_to_hover_over.value_of_css_property('height').replace("px", ""))
    if image_height > 50:
        print(bcolors.OKGREEN + ">> Image can be downloaded" + bcolors.ENDC)
        TestCase_LogResult(**data["testcase_result"]["talk2"]["attach_file_message"]["pass"])
        TestCase_LogResult(**data["testcase_result"]["talk2"]["view_file_message"]["pass"])
        TestCase_LogResult(**data["testcase_result"]["talk2"]["download_file_message"]["pass"])
    else:
        print(bcolors.OKGREEN + ">> Fail to preview image" + bcolors.ENDC)
        TestCase_LogResult(**data["testcase_result"]["talk2"]["attach_file_message"]["fail"])
        TestCase_LogResult(**data["testcase_result"]["talk2"]["view_file_message"]["fail"])
        TestCase_LogResult(**data["testcase_result"]["talk2"]["download_file_message"]["fail"])

    hover = ActionChains(driver).move_to_element(element_to_hover_over)
    hover.perform()
    print(bcolors.OKGREEN + "- Hover image" + bcolors.ENDC)
    time.sleep(3)
    driver.find_element_by_xpath("//*[@id='hanbiro_message_list_container']//div[@class='simplebar-content']/div/div/div[last()]//span//img/following-sibling::span//*[starts-with(@class,'MuiSvgIcon-root')][1]").click()
    print(bcolors.OKGREEN + "- Forward file" + bcolors.ENDC)
    time.sleep(3)
    forward_user = driver.find_element_by_xpath("//div[@class='MuiDialogContent-root']//input[@class='MuiInputBase-input']")
    forward_user.send_keys(forward_name)
    forward_user.send_keys(Keys.ENTER)
    print(bcolors.OKGREEN + "- Search forward user" + bcolors.ENDC)

    select_user = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH,"//div[contains(@class,'MuiPaper-root')]//div[3]/div/div[1]//h6[contains(text(),'" + str(forward_name) + "')]")))
    time.sleep(2)
    select_user.click()
    try:
        selected_user = driver.find_element_by_xpath("//div[contains(@class,'MuiPaper-root')]//div[3]/div/div[1]//h6[contains(text(),'" + str(forward_name) + "')]/../../../div[contains(@class,'Mui-selected')]")
        if selected_user.is_displayed():
            print(bcolors.OKGREEN + ">> Select user success" + bcolors.ENDC)
            driver.find_element_by_xpath("//div[@class='MuiDialogContent-root']//span[@class='MuiIconButton-label']/input").click()
            driver.find_element_by_xpath("//button//span[contains(@class,'MuiButton-label') and contains(.,'SEND')]").click()
            print(bcolors.OKGREEN + "- Send forward file" + bcolors.ENDC)
        TestCase_LogResult(**data["testcase_result"]["talk2"]["forward_file_message"]["pass"])
    except:
        print(bcolors.OKGREEN + ">> Select user fail" + bcolors.ENDC)
        driver.find_element_by_xpath("//button//span[contains(@class,'MuiButton-label') and contains(.,'Close')]").click()
        TestCase_LogResult(**data["testcase_result"]["talk2"]["forward_file_message"]["fail"])

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

def add_new_member():
    driver.find_element_by_xpath("//*[@id='root']//div[contains(@class,'Pane2')]//div[contains(@class,'MuiBox-root')][1]//button[1]/span[@class='MuiIconButton-label']").click()
    print(bcolors.OKGREEN + "- Add new member" + bcolors.ENDC)
    add_member = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "//*[@class='MuiDialogContent-root']//input[contains(@class,'MuiInputBase-input')]")))
    time.sleep(2)
    add_member.send_keys(add_user)
    add_member.send_keys(Keys.ENTER)
    print(bcolors.OKGREEN + "- Search new member" + bcolors.ENDC)
    
    total = driver.find_element_by_xpath("//*[@class='MuiDialogContent-root']//h6[contains(@class,'MuiTypography-colorTextSecondary')]")
    #print(bcolors.OKGREEN + total.text[0:1])
    #thêm case đếm user sau khi search và chọn tất cả những user hiển thị hoặc kiểm tra org, chọn user bất kì
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//div[contains(@class,'MuiDialog-paperScrollPaper')]/div[1]/div[2]/div/div[1]/div[1]/div[1]/div/div[1]/div[2]/div/div/div/div/div/div[2]"))).click()
    print(bcolors.OKGREEN + "- Select user 1" + bcolors.ENDC)
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//div[contains(@class,'MuiDialog-paperScrollPaper')]/div[1]/div[2]/div/div[1]/div[1]/div[1]/div/div[1]/div[2]/div/div/div/div/div/div[3]"))).click()
    print(bcolors.OKGREEN + "- Select user 2" + bcolors.ENDC)
    time.sleep(3)
    total_update = driver.find_element_by_xpath("//*[@class='MuiDialogContent-root']//h6[contains(@class,'MuiTypography-colorTextSecondary')]")
    total1 = total_update.text
    total1_up = int(total1[0:1])
    time.sleep(2)
    driver.find_element_by_xpath("//div[contains(@class,'MuiDialog-paperScrollPaper')]//button[2]/span[contains(.,'Add')]").click()
    print(bcolors.OKGREEN + "- Click add" + bcolors.ENDC)
    time.sleep(3)

    chat_room = driver.find_elements_by_xpath("//div[contains(@class,'Pane2')]//button[@class='MuiButtonBase-root']/div/div")
    count_chat_room = int(len(chat_room)) + 1
    if count_chat_room == total1_up:
        print(bcolors.OKGREEN + ">> Add new member success" + bcolors.ENDC)
        TestCase_LogResult(**data["testcase_result"]["talk2"]["add_new_member"]["pass"])
    else:
        print(bcolors.OKGREEN + ">> Add new member fail" + bcolors.ENDC)
        TestCase_LogResult(**data["testcase_result"]["talk2"]["add_new_member"]["fail"])

    #Search user at message page
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

def board():
    time.sleep(5)
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH,"//ul[contains(@class,'MuiList-padding')]//div[contains(@aria-label,'Room list')]"))).click()
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
        driver.find_element_by_xpath("//*[@id='hanbiro_message_list_chat_input']//div[3]").click()
        print(bcolors.OKGREEN + "- Click Board" + bcolors.ENDC)
        try:
            write_board()
            TestCase_LogResult(**data["testcase_result"]["talk2"]["access_board_page"]["pass"])
        except:
            pass

        '''count_icon = int(len(driver.find_elements_by_xpath("//*[@id='hanbiro_message_list_chat_input']/div[2]/div[1]/button")))
        if count_icon == 5:
            driver.find_element_by_xpath("//*[@id='hanbiro_message_list_chat_input']//button[5]").click()
            print(bcolors.OKGREEN + "- Click Board" + bcolors.ENDC)
            try:
                write_board()
                TestCase_LogResult(**data["testcase_result"]["talk2"]["access_board_page"]["pass"])
            except:
                pass

        elif count_icon == 3:
            driver.find_element_by_xpath("//*[@id='hanbiro_message_list_chat_input']//button[3]").click()
            print(bcolors.OKGREEN + "- Click Board" + bcolors.ENDC)
            try:
                write_board()
                TestCase_LogResult(**data["testcase_result"]["talk2"]["access_board_page"]["pass"])
            except:
                pass

        else:
            print(bcolors.OKGREEN + "- No button Board" + bcolors.ENDC)
            TestCase_LogResult(**data["testcase_result"]["talk2"]["access_board_page"]["fail"])
            pass'''
    else:
        print(bcolors.OKGREEN + ">> Search contact fail" + bcolors.ENDC)
        pass

def write_board():
    time.sleep(3)
    write_button = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH,"//ul[contains(@class,'MuiList-subheader')]/following-sibling::div//button[contains(@aria-label,'write')]")))
    write_button.click()
    time.sleep(5)
    print(bcolors.OKGREEN + "- Click write button" + bcolors.ENDC)
    input_board = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "//div[contains(@placeholder,'Enter a message')]")))
    time.sleep(2)
    input_board.send_keys(content_board)
    board_attach = driver.find_element_by_xpath("//*[@id='whisper_editor']/following-sibling::div//button[2]//input")
    board_attach.send_keys(attachment)
    print(bcolors.OKGREEN + "- Attach file" + bcolors.ENDC)
    driver.find_element_by_xpath("//h6[contains(text(),'Board write')]/following-sibling::button").click()
    print(bcolors.OKGREEN + "- Save board" + bcolors.ENDC)
    try:
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "//ul[contains(@class,'MuiList-subheader')]//div[@class='simplebar-content']//p[contains(.,'"+ content_board +"')]")))
        time.sleep(3)
        print(bcolors.OKGREEN + ">> Find Board in List" + bcolors.ENDC)
        TestCase_LogResult(**data["testcase_result"]["talk2"]["write_board"]["pass"])
        view_board = driver.find_element_by_xpath("//ul[contains(@class,'MuiList-subheader')]//div[@class='simplebar-content']//p[contains(.,'"+ content_board +"')]/../following-sibling::button[contains(.,'Comment')]")
        view_board.click()
        try:
            comment()
        except:
            pass
        try:
            edit_board()
        except:
            pass
        try:
            delete_board()
        except:
            pass
        time.sleep(3)
    except:
        print(bcolors.OKGREEN + ">> No Board in List" + bcolors.ENDC)
        TestCase_LogResult(**data["testcase_result"]["talk2"]["write_board"]["fail"])
        pass

def comment():
    comt = "Comment Board " + str(n)
    comment_board = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "//*[@id='outlined-multiline-static']")))
    print(bcolors.OKGREEN + "- View board" + bcolors.ENDC)
    comment_board.send_keys(comt)
    comment_board.send_keys(Keys.ENTER)
    
    try:
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "//ul[contains(@class,'MuiList-padding')]//p[contains(.,'"+ comt +"')]")))
        print(bcolors.OKGREEN + ">> Comment board Successfully" + bcolors.ENDC)
        TestCase_LogResult(**data["testcase_result"]["talk2"]["comment_board"]["pass"])
    except:
        print(bcolors.OKGREEN + ">> Comment board Fail" + bcolors.ENDC)
        TestCase_LogResult(**data["testcase_result"]["talk2"]["comment_board"]["fail"])

def edit_board():
    time.sleep(3)
    element_to_hover = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "//ul[contains(@class,'MuiList-padding')]//button[contains(@class,'MuiIconButton-root')]/span/*[@class='MuiSvgIcon-root']")))
    hover_but = ActionChains(driver).move_to_element(element_to_hover)
    hover_but.click().perform()
    print(bcolors.OKGREEN + "- List function" + bcolors.ENDC)
    time.sleep(2)
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "//*[@id='menu-list-grow']/li[contains(.,'Edit')]"))).click()
    print(bcolors.OKGREEN + "- Edit Board" + bcolors.ENDC)

    input_board = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "//div[contains(@placeholder,'Enter a message')]")))
    input_board.clear()
    time.sleep(2)
    input_board.send_keys(content_edit)
    driver.find_element_by_xpath("//h6[contains(text(),'Board write')]/following-sibling::button").click()
    print(bcolors.OKGREEN + "- Save board" + bcolors.ENDC)
    edit_board = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "//ul[contains(@class,'MuiList-subheader')]//div[@class='simplebar-content']//p[contains(.,'"+ content_edit +"')]")))
    time.sleep(3)
    if edit_board.is_displayed():
        print(bcolors.OKGREEN + ">> Edit board Successfully" + bcolors.ENDC)
        TestCase_LogResult(**data["testcase_result"]["talk2"]["edit_board"]["pass"])
    else:
        print(bcolors.OKGREEN + ">> Edit board Fail" + bcolors.ENDC)
        TestCase_LogResult(**data["testcase_result"]["talk2"]["edit_board"]["fail"])

def delete_board():
    time.sleep(3)
    element_to_hover = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "//ul[contains(@class,'MuiList-padding')]//button[contains(@class,'MuiIconButton-root')]/span/*[@class='MuiSvgIcon-root']")))
    hover_but = ActionChains(driver).move_to_element(element_to_hover)
    hover_but.click().perform()
    print(bcolors.OKGREEN + "- List function" + bcolors.ENDC)
    time.sleep(2)
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "//*[@id='menu-list-grow']/li[contains(.,'Delete')]"))).click()
    print(bcolors.OKGREEN + "- Delete Board" + bcolors.ENDC)

    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "alert-dialog-title")))
    print(bcolors.OKGREEN + "- Pop up notification" + bcolors.ENDC)
    time.sleep(2)
    driver.find_element_by_xpath("//*[@id='alert-dialog-title']//following-sibling::div//button[contains(.,'OK')]").click()
    print(bcolors.OKGREEN + "- Click OK" + bcolors.ENDC)

    time.sleep(3)
    try:
        edit_board = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//ul[contains(@class,'MuiList-subheader')]//div[@class='simplebar-content']//p[contains(.,'"+ content_edit +"')]")))
        print(bcolors.OKGREEN + ">> Delete board Fail" + bcolors.ENDC)
        TestCase_LogResult(**data["testcase_result"]["talk2"]["delete_board"]["fail"])
    except:
        print(bcolors.OKGREEN + ">> Delete board Successfully" + bcolors.ENDC)
        TestCase_LogResult(**data["testcase_result"]["talk2"]["delete_board"]["pass"])
            
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
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "//*[@id='root']//div[contains(@class,'MuiInputBase-root')]/following-sibling::button/span[@class='MuiIconButton-label']"))).click()  

    search_contact = driver.find_element_by_xpath("//*[@id='root']//div[contains(@class,'MuiInputBase-root')]/input[contains(@class,'MuiInputBase-input')]")
    search_contact.send_keys(contact_org)
    search_contact.send_keys(Keys.ENTER)

    try:
        whisper_search = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "//*[@class='simplebar-mask']//div[@class='simplebar-content']//div[1]//h6[contains(.,'"+ str(contact_org) +"')]")))
        whisper_search.click()
        print(bcolors.OKGREEN + ">> Search user success" + bcolors.ENDC)
        TestCase_LogResult(**data["testcase_result"]["talk2"]["whisper_search"]["pass"])
        WebDriverWait(driver, 40).until(EC.presence_of_element_located((By.XPATH, "//div[contains(@class,'Pane2')]/div[1]/div[1]/div[2]/div/div[@class='han-loading mini']")))
        time.sleep(3)
        print(bcolors.OKGREEN + "- View whisper" + bcolors.ENDC)
        view_whisper()
    except:
        print(bcolors.OKGREEN + ">> Search user fail" + bcolors.ENDC)
        TestCase_LogResult(**data["testcase_result"]["talk2"]["whisper_search"]["fail"])
        driver.find_element_by_xpath("//div[contains(@class,'Pane1')]//button[contains(@class,'MuiIconButton-sizeSmall')][2]/span[1]").click()
        print(bcolors.OKGREEN + "- Delete keys search" + bcolors.ENDC)

def view_whisper():
    try:
        whisper_content = driver.find_element_by_xpath("//div[contains(@class,'Pane2')]//div[@class='simplebar-content']//div[contains(@class,'MuiTypography-body2') and contains(text(),'"+ str(content_whisper) +"')]")
        print(bcolors.OKGREEN + ">> View whiper success" + bcolors.ENDC)
        TestCase_LogResult(**data["testcase_result"]["talk2"]["send_whisper"]["pass"])
        TestCase_LogResult(**data["testcase_result"]["talk2"]["view_whisper"]["pass"])
    except:
        print(bcolors.OKGREEN + ">> View whisper fail" + bcolors.ENDC)
        TestCase_LogResult(**data["testcase_result"]["talk2"]["send_whisper"]["fail"])
        TestCase_LogResult(**data["testcase_result"]["talk2"]["view_whisper"]["fail"])

    driver.find_element_by_xpath("//div[contains(@class,'Pane1')]//button[contains(@class,'MuiIconButton-sizeSmall')][2]/span[1]").click()
    print(bcolors.OKGREEN + "- Delete key search" + bcolors.ENDC)
    time.sleep(3)

    count_filter = int(len(driver.find_elements_by_xpath("//div[@class='simplebar-content']/div/div/div[contains(@class,'MuiListItem-gutters')]")))
    print(count_filter)
    if count_filter > 1:
        first_sort = driver.find_element_by_xpath("//*[@class='simplebar-mask']//div[@class='simplebar-content']//div[contains(@class,'MuiButtonBase-root')][1]//h6")
        first_sort_name = first_sort.text
        print(bcolors.OKGREEN + "--- Before filter: " + first_sort_name + bcolors.ENDC)
        time.sleep(2)
        driver.find_element_by_xpath("//div[contains(@class,'Pane1')]//button[4]/span[1]").click()
        time.sleep(2)
        print(bcolors.OKGREEN + "- Filter" + bcolors.ENDC)
        after_sort = driver.find_element_by_xpath("//*[@class='simplebar-mask']//div[@class='simplebar-content']//div[contains(@class,'MuiButtonBase-root')][1]//h6")
        after_sort_name = after_sort.text
        print(bcolors.OKGREEN + "--- After filter: " + after_sort_name + bcolors.ENDC)
        #need to check counter > 1 before compare
        if first_sort_name != after_sort_name:
            print(bcolors.OKGREEN + "=> Filter success" + bcolors.ENDC)
            TestCase_LogResult(**data["testcase_result"]["talk2"]["filter_whisper"]["pass"])
        else:
            print(bcolors.OKGREEN + "=> Filter fail" + bcolors.ENDC)
            TestCase_LogResult(**data["testcase_result"]["talk2"]["filter_whisper"]["fail"])
    else:
        print(bcolors.OKGREEN + "=> Not Filter" + bcolors.ENDC)
    





login()
organization()
message()
board()
whisper()

