# -*- coding: Shift-JIS -*-#
#
#  File Name: ManHourSystemWrapper.py
# ------------------------------------------------------------------------------------
#
# History ########################################
#	Aug.1 ,2022 Creation
##############################################
#**** MODULES ****
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome import service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time

#**** FUNCTIONS ****
class MHSUtils:
#
#
#
    driver = ""
    def __init__(self, DATEFROM, DATETO, GID, SEIBAN):
        global driver
        print("from/to/GID/Seiban",DATEFROM, DATETO, GID, SEIBAN)
        options = Options()
        #options.add_argument('--headless')
        options.add_argument('--disable-logging')
        options.add_argument('--log-level=3')
        options.add_experimental_option('excludeSwitches', ['enable-logging'])
        CHROMEDRIVER="C:/Users/0000107049/Documents/Programming/Python/Scraping/chromedriver104_win32/chromedriver.exe"
        chrome_service = service.Service(executable_path=CHROMEDRIVER)
        driver = webdriver.Chrome(service=chrome_service,options=options)
        driver.get('https://ebisap.emcs.sony.co.jp:8092/view/login.aspx')
        time.sleep(5)
        LoginButtons=driver.find_elements(By.NAME, 'Button1')
        print("LoginButtons=",LoginButtons)
        for Button_Login in LoginButtons:
            Button_Login.send_keys(Keys.ENTER)
        #Change Window Handle
        handle_array = driver.window_handles
        # 一番最後のdriverに切り替える
        driver.switch_to.window(handle_array[-1])
        ReferLinks=driver.find_elements(By.ID, 'btnApproveRefer')
        print("ReferLinks=",ReferLinks)
        for ReferLink in ReferLinks:
            print("ReferLink=",ReferLink)
            ReferLink.send_keys(Keys.ENTER)
        print("current_URL=",driver.current_url)
        #Change Window Handle
        handle_array = driver.window_handles
        # 一番最後のdriverに切り替える
        driver.switch_to.window(handle_array[-1])

        InputItems=driver.find_elements(By.NAME, 'txtSearchWorkDateFrom')
        for InputItem in InputItems:
            InputItem.clear()
#            InputItem.send_keys("2022/07/01")
            InputItem.send_keys(DATEFROM)

        InputItems=driver.find_elements(By.NAME, 'txtSearchWorkDateTo')
        for InputItem in InputItems:
            InputItem.clear()
#            InputItem.send_keys("2022/07/15")
            InputItem.send_keys(DATETO)

        InputItems=driver.find_elements(By.NAME, 'txtSearchGlobalId')
        for InputItem in InputItems:
            InputItem.clear()
#            InputItem.send_keys("0000107049")
            InputItem.send_keys(GID)

        InputItems=driver.find_elements(By.NAME, 'txtSearchJoNum')
        for InputItem in InputItems:
            InputItem.clear()
            InputItem.send_keys(SEIBAN)

        InputItems=driver.find_elements(By.NAME, 'btnSearch')
        for InputItem in InputItems:
            InputItem.send_keys(Keys.ENTER)

    def get_date(self):
        Date_List=[]
        num=2
        key= "MyGrid_ctl02_lblWorkDate"
        while driver.find_elements(By.ID, key):
            InputItems=driver.find_elements(By.ID, key)
            for InputItem in InputItems:
#                print("Item=",InputItem.text)
                Date_List.append(InputItem.text)
            num += 1
            N=f'{num:02}'
            key= 'MyGrid_ctl'+N+'_lblWorkDate'
        return Date_List

    def get_seiban(self):
        Seiban_List=[]
        num=2
        key= "MyGrid_ctl02_lblJoNum"
        while driver.find_elements(By.ID, key):
            InputItems=driver.find_elements(By.ID, key)
            for InputItem in InputItems:
#                print("Item=",InputItem.text)
                Seiban_List.append(InputItem.text)
            num += 1
            N=f'{num:02}'
            key= 'MyGrid_ctl'+N+'_lblJoNum'
#        print("key=",key)
        return Seiban_List

    def get_gids(self):
        GID_List=[]
        num=2
        key= "MyGrid_ctl02_lblEmpCode"
        while driver.find_elements(By.ID, key):
            InputItems=driver.find_elements(By.ID, key)
            for InputItem in InputItems:
#                print("Item=",InputItem.text)
                GID_List.append(InputItem.text)
            num += 1
            N=f'{num:02}'
            key= 'MyGrid_ctl'+N+'_lblEmpCode'
#            print("key=",key)
        return GID_List

    def get_manhour(self):
        ManPower_List=[]
        num=2
        key= "MyGrid_ctl02_lblManHour"
        while driver.find_elements(By.ID, key):
            InputItems=driver.find_elements(By.ID, key)
            for InputItem in InputItems:
#                print("Item=",InputItem.text)
                ManPower_List.append(InputItem.text)
            num += 1
            N=f'{num:02}'
            key= 'MyGrid_ctl'+N+'_lblManHour'
#            print("key=",key)
        return ManPower_List

    def __del__(self):
        driver.quit()

