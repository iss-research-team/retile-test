import time
# import PyQt5.sip
import os
import pyperclip as cp
import pyautogui as p
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import time, traceback
import xlrd
import csv


def readconfig():
    ans = []
    with open('conf/account.txt', 'r', encoding='utf-8') as f:
        line = f.readline()
        while line:
            ans.append(line)
            line = f.readline()
    return ans[0], ans[1]


def get_data():
    book = xlrd.open_workbook("input/company__name.xls")
    table = book.sheet_by_index(1)
    data = table.col_values(1)[1:]
    return data


def output_big_company(company_name):
    csv_write = csv.writer(open('output/big_company.csv', 'a+', encoding='utf-8', newline=''))
    for item in company_name:
        csv_write.writerow([item])


def output_not_export_company(data_list_):
    csv_write = csv.writer(open('output/big_company_not.csv', 'a+', encoding='utf-8', newline=''))
    for item in data_list_:
        csv_write.writerow([item])


def login():
    browser.get(login_url)
    browser.maximize_window()
    login_in = wait.until(EC.presence_of_element_located((By.CLASS_NAME, 'tr-login-login_btn')))
    time.sleep(1)
    name_tag = browser.find_element(by=By.ID,value='tr-login-username')
    name_tag.send_keys(user)
    time.sleep(1)
    pwd_tag = browser.find_element(by=By.ID,value='tr-login-password')
    pwd_tag.send_keys(pwd)
    time.sleep(1)
    login_in.click()
    time.sleep(3)
    for _ in range(3):
        try:
            time.sleep(1)
            cookie_button = browser.find_element(by=By.ID,value='onetrust-accept-btn-handler')
            cookie_button.click()
            break
        except:
            p.hotkey('ctrlleft', 'r')
            time.sleep(3)


def check_query(query_list):
    time.sleep(2)
    # browser.get('https://derwentinnovation.clarivate.com.cn/ui/zh/#/home/patent-search')
    # for _ in range(3):
    #     try:
    #         time.sleep(1)
    #         cookie_button = browser.find_element(by=By.ID,value='onetrust-accept-btn-handler')
    #         cookie_button.click()
    #         break
    #     except:
    #         p.hotkey('ctrlleft', 'r')
    #         time.sleep(3)
    home_button = wait.until(EC.element_to_be_clickable(
        (By.XPATH, "//body/di-app/di-app-header/div/header/div/nav/div/a")))
    home_button.click()
    patent_query_button = wait.until(EC.element_to_be_clickable(
        (By.XPATH,
         "//body/di-app/div[1]/div[3]/di-app-home/main/section[1]/div/di-app-smart-search/div/div/di-app-search-type/div[1]/mat-card/a/div[3]")))  # ????????????????????????
    patent_query_button.click()
    expert_button = wait.until(
        EC.presence_of_element_located((By.XPATH, '//button/span/span[contains(text(), "??????")]')))  # ????????????????????????
    expert_button.click()
    query_all_company = ''
    for i, company in enumerate(query_list):
        if i == 0:
            query_all_company = 'PA = ' + '(\"' + str(company) + '\")'
        else:
            query_all_company += 'OR PA = ' + '(\"' + str(company) + '\")'
    cp.copy(query_all_company)
    time.sleep(1)
    # expert_button = wait300.until(
    #     EC.presence_of_element_located((By.XPATH, '//button/span/span[contains(text(), "??????")]')))  # ????????????????????????

    p.scroll(-300)
    time.sleep(1)
    p.click(400, 900, 1)
    time.sleep(1)
    p.hotkey('ctrl', 'a')
    time.sleep(1)
    p.press('delete')
    time.sleep(1)
    p.hotkey('ctrlleft', 'v')
    query_button = wait.until(EC.element_to_be_clickable(
        (By.XPATH, "//mat-card//div[@class = 'content-footer-buttons fixed-container']/button[3]")))
    query_button.click()
    number = wait300.until(EC.element_to_be_clickable(
        (By.XPATH, "//body/di-app/div[1]/div[3]/di-app-search-results/div/div/div/div/div[1]/span[1]")))  # ???????????????????????????
    time.sleep(1)
    if int(number.text) > 59900:
        return False
    else:
        return True


def query(query_list):
    time.sleep(2)
    # browser.get('https://derwentinnovation.clarivate.com.cn/ui/zh/#/home/patent-search')
    home_button = wait.until(EC.element_to_be_clickable(
        (By.XPATH, "// body/di-app/di-app-header/div/header/div/nav/div/a")))
    home_button.click()
    patent_query_button = wait.until(EC.element_to_be_clickable(
        (By.XPATH, "//body/di-app/div[1]/div[3]/di-app-home/main/section[1]/div/di-app-smart-search/div/div/di-app-search-type/div[1]/mat-card/a/div[3]")))  # ????????????????????????
    patent_query_button.click()
    expert_button = wait.until(
        EC.presence_of_element_located((By.XPATH, '//button/span/span[contains(text(), "??????")]')))  # ????????????????????????
    expert_button.click()
    query_all_company = ''
    for i, company in enumerate(query_list):
        if i == 0:
            query_all_company = 'PA = ' + '(\"' + str(company) + '\")'
        else:
            query_all_company += ' OR PA = ' + '(\"' + str(company) + '\")'
    cp.copy(query_all_company)
    p.scroll(-300)
    time.sleep(1)
    p.click(400, 900, 1)
    time.sleep(1)
    p.hotkey('ctrl', 'a')
    time.sleep(1)
    p.press('delete')
    time.sleep(1)
    p.hotkey('ctrlleft', 'v')
    query_button = wait.until(EC.element_to_be_clickable(
        (By.XPATH, "//mat-card//div[@class = 'content-footer-buttons fixed-container']/button[3]")))
    query_button.click()
    time.sleep(1)
    expert_button = wait300.until(EC.element_to_be_clickable(
        (By.XPATH, '//span/span[contains(text(), "??????")]')))
    expert_button.click()


def export(search_list, order_):
    time.sleep(1)
    iframe = wait.until(
        EC.presence_of_element_located((By.XPATH, "//div[@class='em-iframe-container']/iframe")))
    time.sleep(1)
    browser.switch_to.frame(iframe)
    time.sleep(1)
    browser.switch_to.frame("statusframe")  # ?????????3???frame  statusframe
    print("????????????iframe")
    build_button = wait.until(EC.presence_of_element_located((By.ID, 'createButton')))  # ????????????????????????
    build_button.click()  # ??????
    browser.switch_to.default_content()
    print("??????")

    time.sleep(1)
    iframe = wait.until(
        EC.presence_of_element_located((By.XPATH, "//main/iframe[@name = 'di-iframe']")))  # ????????????????????????
    browser.switch_to.frame("di-iframe")
    order_id = wait300.until(EC.element_to_be_clickable(
        (By.XPATH,
         '//td[@id="twt:account-orders"]/table / tbody / tr[1] / td / table / tbody / tr[2] / td / table / tbody / tr / td / table / tbody / tr[1] / td / div / div / table / tbody / tr[2] / td[3] / div')))
    if order_ != order_id.text:
        order_ = order_id.text
    else :
        output_not_export_company(search_list)
    print("OK")



def run():
    try:
        login()
        for _ in range(3):
            try:
                time.sleep(1)
                cookie_button = browser.find_element(by=By.ID,value='onetrust-accept-btn-handler')
                cookie_button.click()
                print("cookie??????")
                break
            except:
                p.hotkey('ctrlleft', 'r')
                time.sleep(10)
        i = 196
        order_identify = ''
        while True:
            search_list = []
            search_list.append(data_list[i])
            i += 1
            while check_query(search_list):
                search_list.append(data_list[i])
                i += 1
            if len(search_list) == 1:
                output_big_company([search_list[0]])
                print("big")
                print(search_list)
                continue
            else:
                search_list.pop()
                i -= 1
            print(i)
            query(search_list)
            print(search_list)
            export(search_list, order_identify)
            print("chenggong")
    except:
        f = traceback.format_exc()
        browser.quit()
        p.alert('??????????????? Error???' + '\n' + f)

def rerun():
    login()
    browser.refresh()
    time.sleep(10)
    browser.get(login_url)

if __name__ == '__main__':
    login_url = 'https://www.derwentinnovation.com/login/'
    search_url = 'https://www.derwentinnovation.com/ui/zh/#/home'
    t = 1
    user, pwd = readconfig()
    # ???????????????????????????
    p.FAILSAFE = True
    # ????????????PyAutoGUI??????????????????????????????????????????0.1??????
    p.PAUSE = 0.1
    driver_path = 'drive/chromedriver.exe'
    browser = webdriver.Chrome(executable_path=driver_path)

    wait = WebDriverWait(browser, 10)
    wait300 = WebDriverWait(browser, 300)
    data_list = get_data()
    print('start')
    rerun()
