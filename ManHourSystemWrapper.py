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
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys

# Chromeの場合
#from selenium.webdriver.chrome.options import Options
#from selenium.webdriver.chrome import service

# Edgeの場合
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options

import time
from selenium.webdriver.common.alert import Alert

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
#        CHROMEDRIVER="C:/Users/0000107049/Documents/Programming/Python/Scraping/chromedriver104_win32/chromedriver.exe"

        CHROMEDRIVER="C:/Users/0000107049/Documents/Programming/Python/Scraping/edgedriver_win64/msedgedriver.exe"

        chrome_service = service.Service(executable_path=CHROMEDRIVER)
#        driver = webdriver.Chrome(service=chrome_service,options=options)

        driver = webdriver.Edge(executable_path=CHROMEDRIVER)


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



class MHSInputUtils:
#
#
#
    driver = ""
    date_target=""
    key_tag=""
    def __init__(self, DATEFROM, GID, SEIBAN):
        global driver
        global date_target
        date_target=DATEFROM
#        print("from/to/GID/Seiban",DATEFROM, GID, SEIBAN)
        options = Options()
        options.add_argument('--headless')
        options.add_argument('--disable-logging')
        options.add_argument('--log-level=3')
        options.add_experimental_option('excludeSwitches', ['enable-logging'])

        options.use_chromium = True

#        CHROMEDRIVER="C:/Users/0000107049/Documents/Programming/Python/Scraping/chromedriver104_win32/chromedriver.exe"
        CHROMEDRIVER="C:/Users/0000107049/Documents/Programming/Python/Scraping/edgedriver_win64/msedgedriver.exe"
#        chrome_service = service.Service(executable_path=CHROMEDRIVER)
#        driver = webdriver.Chrome(service=chrome_service,options=options)
#        driver = webdriver.Edge(executable_path=CHROMEDRIVER)
        driver = webdriver.Edge(service=Service(CHROMEDRIVER),options=options)
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
        ReferLinks=driver.find_elements(By.ID, 'btnJissekiInput')
        for ReferLink in ReferLinks:
            ReferLink.send_keys(Keys.ENTER)
        #Change Window Handle
        handle_array = driver.window_handles
        # 一番最後のdriverに切り替える
        driver.switch_to.window(handle_array[-1])

        # 年月を切り替え

    def get_colums(self):
        """
        try:
            tableElem = driver.find_element_by_class_name("jisseki_input_table")
            trs = tableElem.find_elements(By.TAG_NAME, "tr")
         # ヘッダ行は除いて取得
            for i in range(5,len(trs)):
                tds = trs[i].find_elements(By.TAG_NAME, "td")
                line = ""
                for j in range(0,len(tds)):
                    if j < len(tds)-1:
                        line += "%s\t" % (tds[j].text)
                    else:
                        line += "%s" % (tds[j].text)
                print(line+"\r\n")
        except:
            print(traceback.format_exc())
        """

        Seiban_List=[]
        num=0
        key= "hdnJoNum00"
        while driver.find_elements(By.ID, key):
            InputItems=driver.find_elements(By.ID, key)
            for InputItem in InputItems:
                print("Item=",InputItem.get_attribute('value'))
                Seiban_List.append(InputItem.get_attribute('value'))
            num += 1
            N=f'{num:02}'
            key= 'hdnJoNum'+N

        SeibanName_List=[]
        num=2
        key= "/html/body/form/div[4]/table/tbody/tr[7]/td[2]"
        while driver.find_elements(By.XPATH,key):
            InputItems=driver.find_elements(By.XPATH,key)
            for InputItem in InputItems:
                print("Item(SeibanName)=",InputItem.text)
                SeibanName_List.append(InputItem.text)
            num += 1
            key="/html/body/form/div[4]/table/tbody/tr[7]/td["+str(num)+"]"

        SeibanSubItem_List=[]
        num=2
        key= "/html/body/form/div[4]/table/tbody/tr[11]/td[2]"
        while driver.find_elements(By.XPATH,key):
            InputItems=driver.find_elements(By.XPATH,key)
            for InputItem in InputItems:
                print("Item(SeibanName)=",InputItem.text)
                SeibanSubItem_List.append(InputItem.text)
            num += 1
            key="/html/body/form/div[4]/table/tbody/tr[11]/td["+str(num)+"]"

        return Seiban_List,SeibanName_List,SeibanSubItem_List

    def get_date(self,num_SeibanItems):
        """
        # 最初の行のXPATH↓
        #   //*[@id="tblJissekiInput"]/tbody/tr[14]/td[1]
        # 2番目の行のXPATH↓
        #   //*[@id="tblJissekiInput"]/tbody/tr[15]/td[1]
        # 最後の行のXPATH↓
        #   //*[@id="tblJissekiInput"]/tbody/tr[50]/td[1]
        ↓2月の最終（2/28）
        //*[@id="tblJissekiInput"]/tbody/tr[47]/td[1]
        """
        # tr[14]～最後回して、日にちの一致したtrを探す。
        element = driver.find_element(By.XPATH,"/html/body/form/div[4]/table/tbody/tr[14]/td[1]")
        print("最初の日付",element.text)

        Date_List=[]
        num=14
        key="/html/body/form/div[4]/table/tbody/tr[14]/td[1]"
        try:
            while driver.find_elements(By.XPATH,key):
                Items=driver.find_elements(By.XPATH,key)
                for Item in Items:
                    Date_List.append(Item.text)
                    if Item.text == date_target:
                        print('日付一致 %d(%s)' % (num,date_target))
                        num_target=num
                num += 1
                key="/html/body/form/div[4]/table/tbody/tr["+str(num)+"]" +"/td[1]"
            print("Date_List",Date_List)
        except NoSuchElementException:
            print("Not found")

        #一致する行の要素を配列に入れる
        Kousuu_List=[]
        key="/html/body/form/div[4]/table/tbody/tr["+str(num_target)+"]" +"/td[4]/input[1]"
        num_seiban=4
        num_temp=1
        print("key=",key)
        while driver.find_elements(By.XPATH,key):
            if num_temp > num_SeibanItems:
                print("num_temp=",num_temp)
                break
            Items=driver.find_elements(By.XPATH,key)
            for Item in Items:
                Kousuu_List.append(Item.get_attribute('value'))
                num_seiban+=1
                key="/html/body/form/div[4]/table/tbody/tr["+str(num_target)+"]" +"/td["+str(num_seiban)+"]/input[1]"
                num_temp = num_temp +1
                if num_temp > num_SeibanItems:
                    print("num_temp=",num_temp)
                    break

        print("Kousuu_List=",Kousuu_List)

    def set_date(self, date_year,date_month,date_new,num_SeibanItems):
        global key_tag
        # tr[14]～最後回して、日にちの一致したtrを探す。
#        element = driver.find_element(By.XPATH,"/html/body/form/div[4]/table/tbody/tr[14]/td[1]")
#        print("最初の日付",element.text)

        print("date_new=",date_new)

#        InputItems=driver.find_elements(By.ID, 'ddlYearMonth')
#        for InputItem in InputItems:
#            InputItem.send_keys(date_year+'/'+date_month)
#            InputItem.send_keys("2022/07")

        Date_List=[]
        num=14
        key="/html/body/form/div[4]/table/tbody/tr[14]/td[1]"
        try:
            while driver.find_elements(By.XPATH,key):
                Items=driver.find_elements(By.XPATH,key)
                for Item in Items:
                    Date_List.append(Item.text)
                    if Item.text == date_new:
                        print('日付一致 %d(%s)' % (num,date_new))
                        num_target=num
                num += 1
                key="/html/body/form/div[4]/table/tbody/tr["+str(num)+"]" +"/td[1]"
            print("Date_List",Date_List)
        except NoSuchElementException:
            print("Not found")

        #一致する行の要素を配列に入れる
        Kousuu_List=[]
        key_list=[]
        key="/html/body/form/div[4]/table/tbody/tr["+str(num_target)+"]" +"/td[4]/input[1]"
        key_tag=key
        num_seiban=4
        num_temp=1
        print("key=",key)
        while driver.find_elements(By.XPATH,key):
            key_list.append(key)
            Items=driver.find_elements(By.XPATH,key)
            if num_temp > num_SeibanItems:
                break;
            for Item in Items:
                print("Item:set_date = ",Item)

                Kousuu_List.append(Item.get_attribute('value'))
                num_seiban+=1
                key="/html/body/form/div[4]/table/tbody/tr["+str(num_target)+"]" +"/td["+str(num_seiban)+"]/input[1]"
#                key="/html/body/form/div[4]/table/tbody/tr["+str(num_target)+"]" +"/td[4]/input[1]"
                num_temp = num_temp +1
                if num_temp > num_SeibanItems:
                    print("num_temp=",num_temp)
                    break
        print("Kousuu_List=",Kousuu_List)

#   Get Lysithea Data
        key = "/html/body/form/div[4]/table/tbody/tr["+str(num_target)+"]" +"/td[2]"
        Items=driver.find_elements(By.XPATH,key)
        for Item in Items:
            lysithea_data = Item.text
            print("lysithea_data=",lysithea_data)

        return key_list,key_tag, Kousuu_List,lysithea_data

    def send_updated_kousuu(self,key,new_kousuu):
#        global key_tag
        global driver
#        key=key_tag
        print("key：",key)
        if len(key) == 0 :
            print("key is empty")
            return
        """
        #Change Window Handle
        handle_array = driver.window_handles
        # 一番最後のdriverに切り替える
        driver.switch_to.window(handle_array[-1])
        """
#        print("driver=",driver)
#        while driver.find_elements(By.XPATH,key):
#        key="/html/body/form/div[4]/table/tbody/tr[25]/td[4]/input[1]"
#        key='//*[@id="txt00_12"]'
        Items=driver.find_elements(By.XPATH, key)
        for Item in Items:
            print("Item=", Item.text)
            Item.clear()
            Item.send_keys(new_kousuu)
            print("新しい工数は：",new_kousuu)

    def click_register(self):
        RegisterButtons=driver.find_elements(By.NAME, 'btnEdit')
        for RegisterButton in RegisterButtons:
#            RegisterButton.send_keys(Keys.ENTER)
#            RegisterButton.submit()
            RegisterButton.click()
            print("登録押されました")
        time.sleep(3)
        Alert(driver).accept()

    def __del__(self):
        driver.quit()

