from selenium import webdriver
from selenium.webdriver.common.by import By
from datetime import date
import time
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from readexcel import get_values_from_excel
from selenium.common.exceptions import StaleElementReferenceException



def selenium_script_add_values():
    chromedriver_location = r"chromedriver\chromedriver.exe"
    excel_file = r'excel files\PIS.xlsx'
    driver = webdriver.Chrome(chromedriver_location)
    pis_link = 'http://prod.pistime.com/ClockViewR5/LoginR5.aspx'
    driver.get(pis_link)
    login(get_id(excel_file),get_pass(excel_file),driver)
    go_to_time_card(driver)
    add_to_time_card(excel_file,driver)
    time.sleep(20)
    driver.close()

def get_id(excel_file):
    for i in get_values_from_excel(excel_file):
        id = i[0]
        return id

def get_pass(excel_file):
    for i in get_values_from_excel(excel_file):
        password = i[1]
        return password


def login(id, passw,driver):

    #select login components on login page
    company_component = '//*[@id="txtCompany"]'
    username_component = '//*[@id="txtUsername"]'
    password_component = '//*[@id="txtPassword"]'
    login_button = '//*[@id="cmdLogin"]'


    #login values
    company = "AGS"
    username = id
    password = passw
    
    #login
    driver.find_element_by_xpath(company_component).send_keys(company)
    driver.find_element_by_xpath(username_component).send_keys(username)
    driver.find_element_by_xpath(password_component).send_keys(password)
    driver.find_element_by_xpath(login_button).click()


def go_to_time_card(driver):
    time_card = '//*[@id="mainMenu"]/ul/li[2]/a/span'
    labor_time = '//*[@id="mainMenu"]/ul/li[2]/div'
    driver.find_element_by_xpath(time_card).click()
    time.sleep(0.25)
    driver.find_element_by_xpath(labor_time).click()


def add_to_time_card(excel_file,driver):
    #click on today's add button
    today = date.today()
    day_today = today.weekday()
    date_today = today.strftime("%m/%d")
    current_frame = driver.current_window_handle
    

    for values in get_values_from_excel(excel_file):
        driver.switch_to.frame(driver.find_element_by_tag_name("iframe"))
        try:
            time.sleep(10)
            day_add = driver.find_element_by_xpath('//*[@id="ctl00_MainContent_rptDays_ctl0'+str(day_today+2)+'_imgbtnAdd1"]')
            day_add.click()
        except:
            wait = WebDriverWait(driver, 10, ignored_exceptions = StaleElementReferenceException)
            xpath_day = '//*[@id="ctl00_MainContent_rptDays_ctl0'+str(day_today+2)+'_imgbtnAdd1"]'
            day_add = wait.until(EC.element_to_be_clickable((By.XPATH, xpath_day)))
            day_add.click()
        


        #select component to add to
        driver.switch_to.frame(driver.find_element_by_tag_name("iframe"))
        hours = driver.find_element_by_xpath('//*[@id="txtHours"]')
        project = driver.find_element_by_xpath('//*[@id="rptLevels_ctl01_cboLevel_Input"]')
        task = driver.find_element_by_xpath('//*[@id="rptLevels_ctl02_cboLevel_Input"]')
        notes = driver.find_element_by_xpath('//*[@id="txtNotes"]')
        update_button = driver.find_element_by_xpath('//*[@id="cmdUpdate_input"]')
    
        #add the values and submit

        hours.send_keys(values[2])
        project.click()
        project.send_keys(values[3])
        time.sleep(0.25)
        project.send_keys(Keys.ENTER)
        task.click()
        time.sleep(1)
        task.send_keys(values[4])
        time.sleep(1)
        task.send_keys(Keys.ENTER)
        notes.send_keys(values[5])
        time.sleep(0.25)
        update_button.click()

        #switch to parent frame
        driver.switch_to.default_content()



