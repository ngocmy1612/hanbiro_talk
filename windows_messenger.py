import time, sys, unittest, random, json, os, pyperclip
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

def PrintYellow(msg):
    '''• Usage: Color msg in yellow'''
    
    Logging(bcolors.WARNING + str(msg) + bcolors.ENDC)

    return msg

def PrintGreen(msg):
    '''• Usage: Color msg in green'''
    
    Logging(bcolors.OKGREEN + str(msg) + bcolors.ENDC)

    return msg

def PrintRed(msg):
    '''• Usage: Color msg in red'''
    
    Logging(bcolors.FAIL + str(msg) + bcolors.ENDC)

    return msg

class Waits():
    def WaitElementLoaded(time, xpath):
        '''• Usage: Wait until element VISIBLE in a selected time period'''
        
        WebDriverWait(driver, time).until(EC.presence_of_element_located((By.XPATH, xpath)))
        element = driver.find_element_by_xpath(xpath)

        return element

    def Wait20s_ElementLoaded(xpath):
        '''• Usage: Wait 20s until element VISIBLE'''
        
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, xpath)))
        element = driver.find_element_by_xpath(xpath)

        return element

    def Wait10s_ElementLoaded(xpath):
        '''• Usage: Wait 10s until element VISIBLE'''
        
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, xpath)))
        element = driver.find_element_by_xpath(xpath)

        return element

    def WaitElementInvisibility(time, xpath):
        '''• Usage: Wait until element INVISIBLE in a selected time period'''
        
        WebDriverWait(driver, time).until(EC.invisibility_of_element_located((By.XPATH, xpath)))
        element = driver.find_element_by_xpath(xpath)

        return element

    def Wait10s_ElementInvisibility(xpath):
        '''• Usage: Wait 10s until element INVISIBLE'''
        
        WebDriverWait(driver, 10).until(EC.invisibility_of_element_located((By.XPATH, xpath)))
        element = driver.find_element_by_xpath(xpath)

        return element
    
    def WaitUntilPageIsLoaded(page_xpath):
        if bool(page_xpath) == True:
            # wait until page's element is present
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, page_xpath)))

        # check if the loading icon is not present at the page -> page is completely loaded
        try:
            WebDriverWait(driver, 10).until(EC.invisibility_of_element_located((By.XPATH, "//div[@class='loading-dialog hide']")))
        except WebDriverException:
            pass

        '''If page_xpath=None/False -> only check if the loading icon is not present'''

class Commands():
    def FindElement(xpath):
        element = driver.find_element_by_xpath(xpath)

        return element

    def FindElements(xpath):
        element = driver.find_elements_by_xpath(xpath)

        return element

    def ClickElement(xpath):
        '''• Usage: Do the click on element
                return WebElement'''

        element = driver.find_element_by_xpath(xpath)
        element.click()

        return element

    def ClickElements(xpath, element_position):
        '''• Usage: Do the click on element
                return WebElement'''

        element = driver.find_elements_by_xpath(xpath)
        element[element_position].click()

        return element

    def Wait10s_ClickElement(xpath):
        '''• Usage: Wait until the element visible and do the click
                return WebElement'''

        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, xpath)))
        element = driver.find_element_by_xpath(xpath)
        element.click()

        return element

    def InputElementtest(xpath, value):
        driver.find_element_by_xpath(xpath).send_keys(value)

    def InputElement(xpath, value):
        '''• Usage: Send key value in input box
                return WebElement'''
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, xpath)))
        element = driver.find_element_by_xpath(xpath)
        element.send_keys(value)

        return element

    def InputEnterElement(xpath, value):
        '''• Usage: Send key value in input box and Enter
                return WebElement'''
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, xpath)))
        element = driver.find_element_by_xpath(xpath)
        element.send_keys(value)
        element.send_keys(Keys.RETURN)

        return element
    
    def InputElement_2Values(xpath, value1, value2):
        '''• Usage: Send key with 2 values in input box
                return WebElement'''

        element = driver.find_element_by_xpath(xpath)
        element.send_keys(value1)
        element.send_keys(value2)

        return element

    def Wait10s_InputElement(xpath, value):
        '''• Usage: Wait until the input box visible and send key value
                return WebElement'''

        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, xpath)))
        element = driver.find_element_by_xpath(xpath)
        element.send_keys(value)

        return element
    
    def SwitchToFrame(frame_xpath):
        WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.XPATH, frame_xpath)))
        frame = Commands.FindElement(frame_xpath)
        driver.switch_to.frame(frame)

        return frame
    
    def SwitchToDefaultContent():
        driver.switch_to.default_content()

    def ScrollDown():
        '''• Usuage: Scroll down, default height (0,-301)'''
        
        driver.execute_script("window.scrollTo(0,300)")
    
    def ScrollUp():
        '''• Usuage: Scroll down, default height (300,0)'''
        
        driver.execute_script("window.scrollTo(301, 0)")
    
    def Selectbox_ByValue(xpath, value):
        '''• Usage: Wait until select box is loaded
                select by value, return select box
                value = str()'''

        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, xpath)))
        element = driver.find_element_by_xpath(xpath)
        Select(element).select_by_value(value)

        return element
    
    def Selectbox_ByIndex(xpath, index_number):
        '''• Usage: Wait until select box is loaded
                select by the index, return select box
                index_number = int()'''

        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, xpath)))
        element = driver.find_element_by_xpath(xpath)
        Select(element).select_by_index(index_number)

        return element
    
    def Selectbox_ByVisibleText(xpath, selected_text):
        '''• Usage: Wait until select box is loaded
                select by visible text, return select box
                visible text = str()'''

        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, xpath)))
        element = driver.find_element_by_xpath(xpath)
        Select(element).select_by_index(selected_text)

        return element

    def MoveToElement(xpath):
        '''• Usage: Move to view element by ActionChains
                return WebElement'''

        element = driver.find_element_by_xpath(xpath)
        actions = ActionChains(driver)
        actions.move_to_element(element)
        actions.perform()
        time.sleep(1)

        return element

class Functions():
    def GetElementText(xpath):
        '''• Usage: Get and return element_text as str()'''

        element_text = str(driver.find_element_by_xpath(xpath).text)

        return element_text
    
    def GetInputValue(xpath):
        '''• Usage: Get and return input_value as str()
                 Use this function if element is input box'''

        input_element = driver.find_element_by_xpath(xpath)
        input_value = str(input_element.get_attribute("value"))

        return input_value
    
    def GetElementAttribute(xpath, attribute):
        '''• Usage: Get and return element_attribute as str()
                        (attribute can be value of 'class', 'style'... '''

        element = driver.find_element_by_xpath(xpath)
        element_attribute = str(element.get_attribute(attribute))

        return element_attribute

    def GetListLength(xpath):
        '''• Usage: Count how many elements are visible
                return a number int()'''

        list_length = int(len(driver.find_elements_by_xpath(xpath)))

        return list_length
    
    def xpath_ConvertXpath(xpath, replaced_value):
        '''• Usage: xpath which is being used must be written in style 'replaced_text'
                return str()'''

        if type(replaced_value) == int():
            '''It's used to define the order number of element
                        E.g: xpath + "[" + str(i) + "]" '''
                        # i=int()
            element_xpath = str(xpath).replace("order_number", str(replaced_value))
        
        elif type(replaced_value) == str():
            ''' It's used to replace the text in xpath
                        E.g: xpath = xpath + [contains(., 'replaced_text')] '''
                        # replaced_text=str()
            element_xpath = str(xpath).replace("replaced_text", str(replaced_value))
        
        else:
            print("replaced_value must be str() or int()")

        return element_xpath

    def getRandomNumber_fromSpecificRange(assigned_range):
        '''• Usage: Get a list of random numbers
                return a number int()'''

        random_number = int(random(randint(range(assigned_range))))

        return random_number

    def getRandomList_fromSpecificRange(picked_numbers, assigned_range):
        '''• Usage: Get a list of random numbers and remove duplicated number
                return a list()'''

        random_number = random(randint(range(assigned_range)))

        random_list = []
        i=1
        for i in range(assigned_range):
            random_number = random(randint(range(assigned_range)))
            random_list.append(random_number)
            
            random_list = list(dict.fromkeys(random_list))
            if len(random_list) == picked_numbers:
                break
            
            i+=1 

        return random_list

    def RemoveDuplicate_fromList(selected_list):
        '''• Usage: Remove duplicated items in the assigned list
                return the assigned list without duplicated item'''
        
        selected_list = list(dict.fromkeys(selected_list))

        return selected_list

    def checkIf_ElementVisible(xpath):
        '''• Usage: check element is visible
                    return True if element is visible'''
        
        try:
            driver.find_element_by_xpath(xpath)
            return True
        except WebDriverException:
            return False

    def waitIf_ElementVisible(xpath):
        '''• Usage: Wait 10s until element is visible
                    return True if element is visible'''
        
        try:
            Waits.Wait10s_ElementLoaded(xpath)
            return True
        except WebDriverException:
            return False

try:
    folder_execution = "C:\\Users\\Ngoc"
except:
    folder_execution = ""

current_path = os.path.dirname(Path(__file__).absolute())
json_file = current_path + "\\MN_groupware_auto.json"

# Start the web driver
service = webdriver.chrome.service.Service(current_path + "\\chromedriver_talk.exe")
service.start()

#data
attachment = current_path + "\\attachment\\background6.jpg"
file_text = current_path + "\\attachment\\file_text.txt"
user_name = "automationtest"
password_talk = "automationtest1!"
contact_org = "AutomationTest"
chat_content = "Hanbiro test messenger" + objects.date_time
content_whisper = "Hanbiro test whisper " + objects.date_time

# start the app
driver = webdriver.remote.webdriver.WebDriver(
    command_executor=service.service_url,
    desired_capabilities={
        'browserName': 'chrome',
        'goog:chromeOptions': {
            'args': ['develop_mode'],
            'binary': '%s\\AppData\\Local\\Programs\\hanbiro-talk\\HanbiroTalk2.exe' % folder_execution,
            'extensions': [],
            'windowTypes': ['webview']},
        'platform': 'ANY',
        'version': ''},
    browser_profile=None,
    proxy=None,
    keep_alive=False)

def login():
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["talk"]["domain"])))

    domain = driver.find_element_by_xpath(data["talk"]["domain"])
    if bool(domain.get_attribute("value")) == True:
        domain.clear()
        time.sleep(1)
    
    domain.send_keys("myngoc.hanbiro.net")

    time.sleep(1)

    user_id = driver.find_element_by_xpath(data["talk"]["talk_id"])
    if bool(user_id.get_attribute("value")) == True:
        user_id.clear()
        time.sleep(1)
    user_id.send_keys(user_name)
    Commands.InputElement(data["talk"]["password_talk"], password_talk)
    time.sleep(1)
    driver.find_element_by_xpath(data["talk"]["sign_in"]).click()
    try:
        access_page = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["talk"]["access_page"])))
        print(bcolors.OKGREEN + "=> Login success" + bcolors.ENDC)
        TestCase_LogResult(**data["testcase_result"]["talk2"]["login"]["pass"])
    except:
        print(bcolors.OKGREEN + "=> Login fail" + bcolors.ENDC)
        TestCase_LogResult(**data["testcase_result"]["talk2"]["login"]["fail"])

def get_newest_mess():
    # Tạo vòng lặp
    newest_mess = driver.find_element_by_xpath("//div[@class='simplebar-content']//div[@aria-label='grid']/div/div[1]//p/span")
    newest_mess_content= newest_mess.text
    newest_mess.click()
    print("- Get newest mess")
    time.sleep(5)
    get_text = driver.find_element_by_xpath("//div[@class='simplebar-content']//div[contains(@class,'hanbiroToFadeInAndOut')]/div[last()]//div[contains(.,'"+ str(newest_mess_content) +"')]")
    if get_text.is_displayed():
        print("=> Newest mess correct content")

def message():
    #access message tab
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["talk"]["room_list"]))).click()
    time.sleep(5)
    try:
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["talk"]["message_tab"])))
        print("=> Access messenger tab")
        #get_newest_mess()
    except:
        print("=> Cannot access messenger tab")
        pass

    try:
        searchuser()
    except:
        pass

def searchuser():
    driver.find_element_by_xpath(data["talk"]["org"]).click()
    print("- Access ORG")
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["talk"]["mess_page"])))
    time.sleep(3)
    search_contact = driver.find_element_by_xpath(data["talk"]["search_contact"])
    search_contact.send_keys(contact_org)
    search_contact.send_keys(Keys.ENTER)
    time.sleep(2)
    print("- Search myself")
    try:
        contact_search = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["talk"]["contact_search"]+ str(contact_org) +"')]")))
        contact_search.click()
        print("=> Search user success")
        send_mess()
    except:
        print("=> Search user fail")

def send_mess():
    access_page_chat = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["talk"]["access_page_chat"])))
    time.sleep(2)

    if access_page_chat.is_displayed():
        print(bcolors.OKGREEN + ">> Access page chat success" + bcolors.ENDC)
        TestCase_LogResult(**data["testcase_result"]["talk2"]["access_message_page"]["pass"])
        try:
            write_content()
        except:
            pass   

        try:
            srcoll_mess()
        except:
            pass

        try:
            search_user_in_mess()
        except:
            pass    
    else:
        print(bcolors.OKGREEN + ">> Access page chat fail" + bcolors.ENDC)
        TestCase_LogResult(**data["testcase_result"]["talk2"]["access_message_page"]["fail"])
        pass

def write_content():
    #send message: hanbiro test + current time (send to myself) / send with attachment
    input_content = driver.find_element_by_xpath(data["talk"]["input_content"])
    input_content.send_keys(chat_content)
    input_content.send_keys(Keys.ENTER)
    print(bcolors.OKGREEN + "- Input content chat" + bcolors.ENDC)

    result_chat = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["talk"]["result_chat"] + str(chat_content) + "')]")))
    print(bcolors.OKGREEN + ">> Send message success" + bcolors.ENDC)
    TestCase_LogResult(**data["testcase_result"]["talk2"]["write_content"]["pass"])

    #attach clouddisk
    time.sleep(2)
    driver.find_element_by_xpath(data["talk"]["attach_clouddisk"]).click()
    print("- Attach file Clouddisk")
    attach_clouddisk()
    time.sleep(3)

    #attach PC
    attach_file = driver.find_element_by_xpath(data["talk"]["attach_pc"])
    attach_file.send_keys(attachment)
    print(bcolors.OKGREEN + "- Attach file PC" + bcolors.ENDC)
    time.sleep(5)

def attach_clouddisk():
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["talk"]["clouddisk_page"])))
    time.sleep(2)
    try:
        no_items = driver.find_element_by_xpath(data["talk"]["no_items"])
        if no_items.is_displayed():
            print("=> No file in Clouddisk to attach")
            driver.find_element_by_xpath(data["talk"]["close_button"]).click()
    except:
        count_file = int(len(driver.find_elements_by_xpath(data["talk"]["count_file"])))
        if count_file > 2:
            driver.find_element_by_xpath(data["talk"]["file1"]).click()
            driver.find_element_by_xpath(data["talk"]["file2"]).click()
            print("=> Select file")
            time.sleep(2)
        driver.find_element_by_xpath(data["talk"]["send_file"][0]).click()
        print("=> Send attach file clouddisk")

def srcoll_mess():
    count_list = int(len(driver.find_elements_by_xpath(data["talk"]["count_list"])))
    scroll_newmess = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["talk"]["scroll_newmess"])))
    ActionChains(driver).move_to_element(scroll_newmess).release().perform()
    print(bcolors.OKGREEN + "scroll success" + bcolors.ENDC)
    time.sleep(2)
    coutn_list1 = int(len(driver.find_elements_by_xpath(data["talk"]["count_list"])))

    if coutn_list1 < count_list:
        print("=> Scroll up to view older messages fail")
    else:
        print("=> Scroll up to view older messages success")

def attach_clouddisk_whisper():
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["talk"]["clouddisk_page"])))
    time.sleep(2)
    try:
        no_items = driver.find_element_by_xpath(data["talk"]["no_items"])
        if no_items.is_displayed():
            print("=> No file in Clouddisk to attach")
            driver.find_element_by_xpath(data["talk"]["close_button"]).click()
    except:
        count_file = int(len(driver.find_elements_by_xpath(data["talk"]["count_file"])))
        if count_file > 2:
            driver.find_element_by_xpath(data["talk"]["file1"]).click()
            driver.find_element_by_xpath(data["talk"]["file2"]).click()
            print("=> Select file")
            time.sleep(2)
        driver.find_element_by_xpath(data["talk"]["send_file"][1]).click()
        print("=> Send attach file clouddisk")

def search_user_in_mess():
    #Search user in message tab
    driver.find_element_by_xpath(data["talk"]["search_user"]).click()
    time.sleep(3)
    search_contact = driver.find_element_by_xpath(data["talk"]["search_contact"])
    search_contact.send_keys(contact_org)
    search_contact.send_keys(Keys.ENTER)
    print(bcolors.OKGREEN + "- Search User" + bcolors.ENDC)
    time.sleep(2)

    contact_search = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH,data["talk"]["contact_search"]+ str(contact_org) +"')]")))
    if contact_search.is_displayed():
        print(bcolors.OKGREEN + ">> Search contact success" + bcolors.ENDC)
        contact_search.click()
        TestCase_LogResult(**data["testcase_result"]["talk2"]["message_search"]["pass"])
    else:
        print(bcolors.OKGREEN + ">> Search contact fail" + bcolors.ENDC)
        TestCase_LogResult(**data["testcase_result"]["talk2"]["message_search"]["fail"])
            
def whisper():
    driver.find_element_by_xpath(data["talk"]["org"]).click()
    WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.XPATH, data["talk"]["whisper"])))
    time.sleep(5)

    search_contact = driver.find_element_by_xpath(data["talk"]["search_contact"])
    search_contact.send_keys(contact_org)
    search_contact.send_keys(Keys.ENTER)
    time.sleep(2)

    contact_search = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["talk"]["contact_search"]+ str(contact_org) +"')]")))
    if contact_search.is_displayed():
        print(bcolors.OKGREEN + ">> Search user success" + bcolors.ENDC)
        time.sleep(3)
        actionChains = ActionChains(driver)
        actionChains.context_click(contact_search).perform()
        print(bcolors.OKGREEN + "- Right click" + bcolors.ENDC)

        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH,data["talk"]["send_whisper"]))).click()
        time.sleep(3)
        print(bcolors.OKGREEN + "- Send whisper" + bcolors.ENDC)
        send_whisper()
        TestCase_LogResult(**data["testcase_result"]["talk2"]["access_whisper_page"]["pass"])
    else:
        print(bcolors.OKGREEN + ">> Search user fail" + bcolors.ENDC)
        TestCase_LogResult(**data["testcase_result"]["talk2"]["access_whisper_page"]["fail"])

def send_whisper():
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["talk"]["write_whisper"])))
    time.sleep(2)

    input_whisper = driver.find_element_by_xpath(data["talk"]["input_whisper"])
    input_whisper.send_keys(content_whisper)
    time.sleep(3)
    driver.find_element_by_xpath(data["talk"]["clouddisk_button"]).click()
    print("- Attach file Clouddisk")
    attach_clouddisk_whisper()
    time.sleep(2)

    attach_whisper = driver.find_element_by_xpath(data["talk"]["attach_whisper"])
    attach_whisper.send_keys(file_text)
    print(bcolors.OKGREEN + "- Attach file whisper" + bcolors.ENDC)
    time.sleep(2)
    driver.find_element_by_xpath(data["talk"]["send_whis"]).click()
    print(bcolors.OKGREEN + "- Send whisper" + bcolors.ENDC)
    time.sleep(5)
    driver.find_element_by_xpath(data["talk"]["close_whisper"]).click()
    print(bcolors.OKGREEN + "- Close pop up write whisper" + bcolors.ENDC)
    time.sleep(3)

    driver.find_element_by_xpath(data["talk"]["access_whisper_page"]).click()
    print(bcolors.OKGREEN + "- Access Whisper page" + bcolors.ENDC)
    time.sleep(3)


login()
message()
#whisper()

