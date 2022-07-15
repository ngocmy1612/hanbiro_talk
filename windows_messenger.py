import time, sys, unittest, random, json, os, pyperclip, openpyxl
from tracemalloc import stop
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

try:
    folder_execution = "C:\\Users\\Ngoc"
except:
    folder_execution = ""

current_path = os.path.dirname(Path(__file__).absolute())
json_file = current_path + "\\MN_groupware_auto.json"
with open(json_file) as json_file:
        data = json.load(json_file)

# Start the web driver
service = webdriver.chrome.service.Service(current_path + "\\chromedriver_talk.exe")
service.start()

class objects:
    n = random.randint(1,1000)
    now = datetime.now()
    year = now.strftime("%Y")
    month = now.strftime("%m")
    day = now.strftime("%d")
    time1 = now.strftime("%H:%M:%S")
    date_time = now.strftime("%Y/%m/%d, %H:%M:%S")
    date_id = date_time.replace("/", "").replace(", ", "").replace(":", "")[2:]
    testcase_pass = "Test case status: pass"
    testcase_fail = "Test case status: fail"

#data
domain_execution = "myngoc.hanbiro.net"
contact_org = "automationtest1"
password_talk = "automationtest1!"

attachment = current_path + "\\attachment\\background6.jpg"
file_text = current_path + "\\attachment\\file_text.txt"
chat_content = "Hanbiro test messenger " + objects.date_time
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
    
class bcolors:
    HEADER = '\033[95m'
    OKBLUE = '\033[94m'
    OKGREEN = '\033[92m'
    WARNING = '\033[93m'
    FAIL = '\033[91m'
    ENDC = '\033[0m'
    BOLD = '\033[1m'
    UNDERLINE = '\033[4m'

class functions():
    chrome_options = webdriver.ChromeOptions()

    log_testcase = "\\testcase_log\\"
    testcase_log = current_path + log_testcase + "MN_testcase_result_" + str(objects.date_id) + ".xlsx"

    logs = [testcase_log]
    for log in logs:
        if ".txt" in log:
            open(log, "x").close()
        else:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.cell(row=1, column=1).value = "Menu"
            ws.cell(row=1, column=2).value = "Sub-Menu"
            ws.cell(row=1, column=3).value = "Test Case Name"
            ws.cell(row=1, column=4).value = "Status"
            ws.cell(row=1, column=5).value = "Description"
            ws.cell(row=1, column=6).value = "Date"
            ws.cell(row=1, column=7).value = "Tester"
            wb.save(log)

def TestCase_LogResult(menu, sub_menu, testcase, status, description, tester):

    if status == "Pass":
        print(objects.testcase_pass)
    else:
        print(objects.testcase_fail)
    
    wb = openpyxl.load_workbook(functions.testcase_log)
    current_sheet = wb.active
    start_row = len(list(current_sheet.rows)) + 1

    current_sheet.cell(row=start_row, column=1).value = menu
    current_sheet.cell(row=start_row, column=2).value = sub_menu
    current_sheet.cell(row=start_row, column=3).value = testcase
    current_sheet.cell(row=start_row, column=4).value = status
    current_sheet.cell(row=start_row, column=5).value = description
    current_sheet.cell(row=start_row, column=6).value = objects.date_time
    current_sheet.cell(row=start_row, column=7).value = tester

    wb.save(functions.testcase_log)

def Logging(msg):
    print(msg)

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
    def testcasepass(value):
        TestCase_LogResult(**data["testcase_result"]["talk2"][value]["pass"])

    def testcasefail(value):
        TestCase_LogResult(**data["testcase_result"]["talk2"][value]["fail"])

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

def login():
    domain = Waits.Wait20s_ElementLoaded(data["talk"]["domain"])
    if bool(domain.get_attribute("value")) == True:
        domain.clear()
        time.sleep(1)
    
    domain.send_keys(domain_execution)

    time.sleep(1)

    user_id = driver.find_element_by_xpath(data["talk"]["talk_id"])
    if bool(user_id.get_attribute("value")) == True:
        user_id.clear()
        time.sleep(1)
    user_id.send_keys(contact_org)
    Commands.InputElement(data["talk"]["password_talk"], password_talk)
    time.sleep(1)
    Commands.ClickElement(data["talk"]["sign_in"])
    try:
        Waits.Wait20s_ElementLoaded(data["talk"]["access_page"])
        PrintGreen("=> Login success")
        Commands.testcasepass("login")
    except:
        PrintRed("=> Login fail")
        Commands.testcasefail("login")

def get_newest_mess():
    list_mess = Functions.GetListLength(data["talk"]["list_mess"])

    if list_mess == 0:
        PrintYellow("=> List messenger empty")
    else:
        msg_list = []
        i = 0
        for i in range(list_mess):
            i += 1
            talk_msg = Functions.GetElementText(data["talk"]["talk_msg"] % str(i))
            msg_list.append(talk_msg)
            
        list_file = ["jpg", "jpeg", "png"]
        msg_text_list = []
        for msg in msg_list:
            split_smg = msg.split(".")[-1]
            if split_smg in list_file:
                PrintYellow(">>> Msg is image -> cannot define")
            elif msg == "[EMOTICON]":
                PrintYellow(">>> Msg is emoticon -> cannot define")
            else:
                PrintYellow(">>> Msg is text")
                msg_text_list.append(msg)

        for target_msg in msg_text_list:
            target_position = msg_list.index(target_msg) + 1
            try:
                Commands.ClickElement(data["talk"]["talk_user"] % str(target_position))
                PrintYellow("- Talk user")
                newest_content = Functions.GetElementText(data["talk"]["newest_mess_content"] % str(target_position))
                time.sleep(3)
                try:
                    Commands.FindElement(data["talk"]["newest_mess"] % str(newest_content))
                    PrintGreen("=> Newest mess correct content")
                except WebDriverException:
                    PrintRed("=> Newest mess wrong content")

                break
            except WebDriverException:
                PrintYellow("- Talk room")

def message():
    #access message tab
    Commands.Wait10s_ClickElement(data["talk"]["room_list"])
    time.sleep(3)
    try:
        Waits.Wait20s_ElementLoaded(data["talk"]["message_tab"])
        PrintGreen("=> Access messenger tab")
        Commands.testcasepass("access_message_page")
        time.sleep(2)
        get_newest_mess()
    except:
        PrintRed("=> Cannot access messenger tab")
        Commands.testcasefail("access_message_page")

    try:
        searchuser()
    except:
        pass

def searchuser():
    Commands.ClickElement(data["talk"]["org"])
    PrintYellow("- Access ORG")
    Waits.Wait20s_ElementLoaded(data["talk"]["mess_page"])
    time.sleep(3)
    Commands.InputEnterElement(data["talk"]["search_contact"], contact_org)
    time.sleep(2)
    PrintYellow("- Search myself")
    Commands.ClickElement(data["talk"]["contact_search"])
    try:
        Waits.Wait10s_ElementLoaded(data["talk"]["my_room"])
        PrintGreen("=> Search user success")
        Commands.testcasepass("message_search")
        send_mess()
    except:
        PrintRed("=> Search user fail")
        Commands.testcasefail("message_search")

def send_mess():
    try:
        Waits.Wait20s_ElementLoaded(data["talk"]["access_page_chat"])
        time.sleep(2)
        PrintGreen(">> Access page chat success")
        Commands.testcasepass("access_message_page")
        try:
            write_content()
        except:
            pass   
        
        try:
            srcoll_mess()
        except:
            pass 
    except WebDriverException:
        PrintRed(">> Access page chat fail")
        Commands.testcasefail("access_message_page")

def write_content():
    try:
        Commands.Wait10s_ClickElement(data["talk"]["attach_clouddisk"])
        PrintYellow("- Attach file Clouddisk")
        attach_clouddisk()
        time.sleep(3)
    except WebDriverException:
        PrintRed("=> Attach Clouddisk fail")

    try:
        Commands.InputElement(data["talk"]["attach_pc"], attachment)
        PrintYellow("- Attach file PC")
        time.sleep(3)
    except WebDriverException:
        PrintRed("=> Attach PC fail")

    Commands.InputEnterElement(data["talk"]["input_content"], chat_content)
    PrintYellow("- Input content chat")

    try:
        Waits.Wait20s_ElementLoaded(data["talk"]["result_chat"] % str(chat_content))
        PrintGreen(">> Send message success")
        Commands.testcasepass("write_content")
    except:
        PrintRed(">> Send message fail")
        Commands.testcasefail("write_content")

def attach_clouddisk():
    Waits.Wait20s_ElementLoaded(data["talk"]["clouddisk_page"])
    time.sleep(3)
    try:
        Commands.FindElement(data["talk"]["no_items"])
        PrintYellow("=> No file in Clouddisk to attach")
        Commands.ClickElement(data["talk"]["close_button"])
    except:
        count_file = Functions.GetListLength(data["talk"]["count_file"])
        if count_file > 2:
            Commands.ClickElement(data["talk"]["file1"])
            Commands.ClickElement(data["talk"]["file2"])
            PrintYellow("- Select file")
            time.sleep(2)
        Commands.ClickElement(data["talk"]["send_file"][0])
        PrintYellow("=> Send attach file clouddisk")

def srcoll_mess():
    count_list = Functions.GetListLength(data["talk"]["count_list"])
    if count_list > 40:
        Commands.MoveToElement(data["talk"]["scroll_newmess"])
        PrintYellow("- Scroll messenger")
        time.sleep(2)
        coutn_list1 = Functions.GetListLength(data["talk"]["count_list"])

        if coutn_list1 > count_list:
            PrintGreen("=> Scroll up to view older messages success")
        else:
            PrintRed("=> Scroll up to view older messages fail")
    else:
        target = Waits.Wait20s_ElementLoaded(data["talk"]["scroll_bar"])
        scroll_newmess = Waits.Wait20s_ElementLoaded(data["talk"]["scroll_newmess"])
        time.sleep(2)
        hover_bar = ActionChains(driver).move_to_element(target).click_and_hold(target)
        hover_bar.move_to_element(scroll_newmess).release().perform()
        PrintYellow("- Scroll messenger")
        time.sleep(2)

def attach_clouddisk_whisper():
    Waits.Wait20s_ElementLoaded(data["talk"]["clouddisk_page"])
    time.sleep(3)
    try:
        Commands.FindElement(data["talk"]["no_items"])
        PrintYellow("=> No file in Clouddisk to attach")
        Commands.ClickElement(data["talk"]["close_button"])
    except:
        count_file = Functions.GetListLength(data["talk"]["count_file"])
        if count_file > 2:
            Commands.ClickElement(data["talk"]["file1"])
            Commands.ClickElement(data["talk"]["file2"])
            PrintYellow("- Select file")
            time.sleep(2)
        Commands.ClickElement(data["talk"]["send_file"][1])
        PrintYellow("=> Send attach file clouddisk")
            
def whisper_page():
    contact_search = Commands.ClickElement(data["talk"]["contact_search"])
    actionChains = ActionChains(driver)
    actionChains.context_click(contact_search).perform()
    PrintYellow("- Right click")
    Commands.Wait10s_ClickElement(data["talk"]["send_whisper"])
    PrintYellow("- Send whisper")
    time.sleep(2)
    Waits.Wait20s_ElementLoaded(data["talk"]["write_whisper"])
    time.sleep(2)
    
    return contact_search

def whisper():
    try:
        contact_search = whisper_page()
    except:
        contact_search = None

    if bool(contact_search) == True:
        send_whisper()
    else:
        PrintRed(">>> Don't show pop up write whisper")

def send_whisper():
    Commands.InputElement(data["talk"]["input_whisper"], content_whisper)
    time.sleep(2)
    try:
        Commands.Wait10s_ClickElement(data["talk"]["clouddisk_button"])
        PrintYellow("- Attach file Clouddisk")
        attach_clouddisk_whisper()
        time.sleep(2)
    except WebDriverException:
        PrintRed("=> Attach Clouddisk fail")

    try:
        Commands.InputElement(data["talk"]["attach_whisper"], file_text)
        PrintYellow("- Attach file whisper")
    except WebDriverException:
        PrintRed("=> Attach PC fail")

    Commands.Wait10s_ClickElement(data["talk"]["send_whis"])
    PrintYellow("- Send whisper")
    Commands.Wait10s_ClickElement(data["talk"]["close_whisper"])
    PrintYellow("- Close pop up write whisper")
    Commands.Wait10s_ClickElement(data["talk"]["access_whisper_page"])
    PrintYellow("- Access Whisper page")
    time.sleep(5)

    try:
        Waits.Wait10s_ElementLoaded(data["talk"]["whisper_page"])
        PrintGreen("=> Access whisper tab success")
        Commands.testcasepass("access_whisper_page")
    except:
        PrintRed("=> Access whisper tab fail")
        Commands.testcasefail("access_whisper_page")

def log_out():
    Commands.ClickElement(data["talk"]["out_but"])
    PrintYellow("- Log out button")
    Commands.Wait10s_ClickElement(data["talk"]["clear_out"])
    time.sleep(2)
    try:
        Waits.Wait10s_ElementLoaded(data["talk"]["sign_in_but"])
        PrintGreen("=> Log out success")
    except WebDriverException:
        PrintRed("=> Log out fail")

login()
message()
whisper()
log_out()
