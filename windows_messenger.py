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
attachment = current_path + "\\attachment\\background6.jpg"
file_text = current_path + "\\attachment\\file_text.txt"
password_talk = "automationtest1!"
contact_org = "AutomationTest"
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
    
    domain.send_keys("myngoc.hanbiro.net")

    time.sleep(1)

    user_id = driver.find_element_by_xpath(data["talk"]["talk_id"])
    if bool(user_id.get_attribute("value")) == True:
        user_id.clear()
        time.sleep(1)
    user_id.send_keys(contact_org.lower())
    Commands.InputElement(data["talk"]["password_talk"], password_talk)
    time.sleep(1)
    Commands.ClickElement(data["talk"]["sign_in"])
    try:
        access_page = Waits.Wait20s_ElementLoaded(data["talk"]["access_page"])
        PrintGreen("=> Login success")
        Commands.testcasepass("login")
    except:
        PrintRed("=> Login fail")
        Commands.testcasefail("login")

def get_newest_mess():
    # Tạo vòng lặp (chưa fix)
    list_mess = int(len(driver.find_elements_by_xpath("//div[@class='simplebar-content']//div[contains(@role,'rowgroup')]//div[contains(@class,'MuiListItem-button')]")
    ))
    print(list_mess)
    i = 0
    for i in range(list_mess):
        i += 1
        try:
            ngoc = driver.find_element_by_xpath("//div[@class='simplebar-content']//div[contains(@role,'rowgroup')]//div[contains(@class,'MuiListItem-button')]["+ str(i) +"]//button/div/img")
            print(i)
            if ngoc.is_displayed():
                newest_mess = driver.find_element_by_xpath("//div[@class='simplebar-content']//div[@aria-label='grid']/div/div["+ str(i) +"]//p/span")
                newest_mess_content = newest_mess.text
                PrintRed(newest_mess_content)
                newest_mess.click()
                try:
                    text_mess = driver.find_element_by_xpath("//div[@id='hanbiro_message_list_container']//div[@class='simplebar-content']//div[contains(@class,'hanbiroToFadeInAndOut')]/div[last()]//div/span/p")
                    if text_mess.is_displayed():
                        get_text = driver.find_element_by_xpath("//div[@class='simplebar-content']//div[contains(@class,'hanbiroToFadeInAndOut')]/div[last()]//div[contains(.,'"+ str(newest_mess_content) +"')]")
                        if get_text.is_displayed():
                            print("=> Newest mess correct content")
                except:

            break
        except:
            continue






#//div[@class='simplebar-content']//div[contains(@role,'rowgroup')]//div[contains(@class,'MuiListItem-button')][3]//button/div/img




def get_text():
    newest_mess = driver.find_element_by_xpath("//div[@class='simplebar-content']//div[@aria-label='grid']/div/div[1]//p/span")
    newest_mess_content = newest_mess.text
    PrintRed(newest_mess_content)
    # newest_mess.click()
    # print("- Get newest mess")
    # time.sleep(5)
    # get_text = driver.find_element_by_xpath("//div[@class='simplebar-content']//div[contains(@class,'hanbiroToFadeInAndOut')]/div[last()]//div[contains(.,'"+ str(newest_mess_content) +"')]")
    # if get_text.is_displayed():
    #     print("=> Newest mess correct content")

def message():
    #access message tab
    Commands.Wait10s_ClickElement(data["talk"]["room_list"])
    time.sleep(5)
    get_newest_mess()
    # try:
    #     Waits.Wait20s_ElementLoaded(data["talk"]["message_tab"])
    #     PrintGreen("=> Access messenger tab")
    #     Commands.testcasepass("access_message_page")
    #     #get_newest_mess()
    # except:
    #     PrintRed("=> Cannot access messenger tab")
    #     Commands.testcasefail("access_message_page")
    #     pass

    # try:
    #     searchuser()
    # except:
    #     pass

def searchuser():
    Commands.ClickElement(data["talk"]["org"])
    PrintYellow("- Access ORG")
    Waits.Wait20s_ElementLoaded(data["talk"]["mess_page"])
    time.sleep(3)
    Commands.InputEnterElement(data["talk"]["search_contact"], contact_org.lower())
    time.sleep(2)
    PrintYellow("- Search myself")
    try:
        Commands.ClickElement(data["talk"]["contact_search"]+ str(contact_org) +"')]")
        PrintGreen("=> Search user success")
        Commands.testcasepass("message_search")
        send_mess()
    except:
        PrintRed("=> Search user fail")
        Commands.testcasefail("message_search")

def send_mess():
    access_page_chat = Waits.Wait20s_ElementLoaded(data["talk"]["access_page_chat"])
    time.sleep(2)

    if access_page_chat.is_displayed():
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
    else:
        PrintRed(">> Access page chat fail")
        Commands.testcasefail("access_message_page")
        pass

def write_content():
    #send message: hanbiro test + current time (send to myself) / send with attachment
    Commands.InputEnterElement(data["talk"]["input_content"], chat_content)
    PrintYellow("- Input content chat")

    result_chat = Waits.Wait20s_ElementLoaded(data["talk"]["result_chat"] + str(chat_content) + "')]")
    PrintGreen(">> Send message success")
    Commands.testcasepass("write_content")

    #attach clouddisk
    time.sleep(2)
    Commands.ClickElement(data["talk"]["attach_clouddisk"])
    PrintYellow("- Attach file Clouddisk")
    attach_clouddisk()
    time.sleep(3)

    #attach PC
    Commands.InputElement(data["talk"]["attach_pc"], attachment)
    PrintYellow("- Attach file PC")
    time.sleep(5)

def attach_clouddisk():
    Waits.Wait20s_ElementLoaded(data["talk"]["clouddisk_page"])
    time.sleep(2)
    try:
        no_items = driver.find_element_by_xpath(data["talk"]["no_items"])
        if no_items.is_displayed():
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
    Commands.MoveToElement(data["talk"]["scroll_newmess"])
    # scroll_newmess = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, data["talk"]["scroll_newmess"])))
    # ActionChains(driver).move_to_element(scroll_newmess).release().perform()
    PrintYellow("- Scroll messenger")
    time.sleep(2)
    coutn_list1 = Functions.GetListLength(data["talk"]["count_list"])

    if coutn_list1 > count_list:
        PrintGreen("=> Scroll up to view older messages success")
    else:
        PrintRed("=> Scroll up to view older messages fail")

def attach_clouddisk_whisper():
    Waits.Wait20s_ElementLoaded(data["talk"]["clouddisk_page"])
    time.sleep(2)
    try:
        no_items = driver.find_element_by_xpath(data["talk"]["no_items"])
        if no_items.is_displayed():
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
            
def whisper():
    Commands.ClickElement(data["talk"]["org"])
    Waits.Wait20s_ElementLoaded(data["talk"]["whisper"])
    time.sleep(5)

    Commands.InputEnterElement(data["talk"]["search_contact"], contact_org.lower())
    time.sleep(2)

    contact_search = Waits.Wait20s_ElementLoaded(data["talk"]["contact_search"]+ str(contact_org) +"')]")
    if contact_search.is_displayed():
        PrintGreen(">> Search user success")
        time.sleep(3)
        actionChains = ActionChains(driver)
        actionChains.context_click(contact_search).perform()
        PrintYellow("- Right click")

        Commands.Wait10s_ClickElement(data["talk"]["send_whisper"])
        time.sleep(3)
        PrintYellow("- Send whisper")
        send_whisper()
    else:
        PrintRed(">> Search user fail")

def send_whisper():
    Waits.Wait20s_ElementLoaded(data["talk"]["write_whisper"])
    time.sleep(2)

    Commands.InputElement(data["talk"]["input_whisper"], content_whisper)
    time.sleep(2)
    Commands.Wait10s_ClickElement(data["talk"]["clouddisk_button"])
    PrintYellow("- Attach file Clouddisk")
    attach_clouddisk_whisper()
    time.sleep(2)

    Commands.InputElement(data["talk"]["attach_whisper"], file_text)
    PrintYellow("- Attach file whisper")
    Commands.Wait10s_ClickElement(data["talk"]["send_whis"])
    PrintYellow("- Send whisper")
    Commands.Wait10s_ClickElement(data["talk"]["close_whisper"])
    PrintYellow("- Close pop up write whisper")
    Commands.Wait10s_ClickElement(data["talk"]["access_whisper_page"])
    PrintYellow("- Access Whisper page")
    time.sleep(2)

    try:
        driver.find_element_by_xpath(data["talk"]["whisper_page"])
        PrintGreen("=> Access whisper tab success")
        Commands.testcasepass("access_whisper_page")
    except:
        PrintRed("=> Access whisper tab fail")
        Commands.testcasefail("access_whisper_page")

login()
message()
#whisper()

