import openpyxl
from openpyxl import load_workbook
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
import time
import datetime
#Test
####USER INPUT####
clientcode = 'YONDU'
employeecode = '22-02762'
password = 'Aa051498!'
filingremarks = 'No Time In/Time Out'
excel_files = 'D:\RPA\Yondu Timesheet_Novemberr 1-15 2021_Monica Cho.xlsx'

####FROM EXCEL####
timeinvalues = []
timeoutvalues = []
datevalues = []
timeinvaluefinal = []
timeoutvaluefinal = []

####EXCEL####
#LOAD WORKSHEET
sheet = load_workbook(excel_files)
worksheet = sheet.active

#DATE
date_column =  worksheet['B10' : 'B24']

for cell in date_column:
    for date in cell:  
        datevalues.append(date.value)
#TIME IN
timein_column =  worksheet['N10' : 'N24']

for cell in timein_column:
    for timein in cell:  
        timeinvalues.append(timein.value)
#TIME OUT
timeout_column =  worksheet['O10' : 'O24']

for cell in timeout_column:
    for timeout in cell:  
        timeoutvalues.append(timeout.value)
#REMOVE NONE
timeinfinal = [i for i in timeinvalues if i]
timeoutfinal = [i for i in timeoutvalues if i]
datefinal = [i for i in datevalues if i]
iteration = len(datefinal)
#PRINT VALUES
for i in range(iteration):
    timeinformat = timeinfinal[i].strftime("%I:%M %p")
    timeinvaluefinal.append(timeinformat)
    timeoutformat = timeoutfinal[i].strftime("%I:%M %p")
    timeoutvaluefinal.append(timeoutformat)
#print("TIME IN: " + str(timeinvaluefinal))
#print("TIME OUT: " + str(timeoutvaluefinal))
#print("DATE: " + str(datefinal))

####WEB AUTOMATION####
url = 'https://ess.payroll.ph/app'
driver = webdriver.Chrome()
driver.get(url)

driver.find_element(By.ID, "loginClientCode").send_keys(clientcode)
driver.find_element(By.ID, "loginEmployeeCode").send_keys(employeecode)
driver.find_element(By.ID, "loginPassword").send_keys(password)
driver.find_element(By.ID, "loginForm").click()
#Click OB FILING
driver.implicitly_wait(15)
driver.find_element(By.XPATH, '//*[@class="x-portlet-button-img"][@src="/images/flat-icons/clock2.png"]').click()
for i in range(iteration):
    ###VARIABLES####
    datemonth = datefinal[i].strftime("%b")
    dateyear = datefinal[i].strftime("%Y")
    dateday = datefinal[i].day
    #Click NEW
    time.sleep(10)
    driver.find_element(By.XPATH, '//*[@class=" x-btn-text icon-new"][text()="New"]').click()
    #Client
    driver.find_element(By.XPATH, '//*[@src="/images/default/s.gif"][@class="x-form-trigger x-form-arrow-trigger"]').click()
    WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@class="x-layer x-combo-list "]/div/div[normalize-space(text()) = "Work from Home"]'))).click()
    #Date
    driver.find_element(By.XPATH, '//div/div/img[@src="/images/default/s.gif"][@class="x-form-trigger x-form-date-trigger"]').click()
    driver.find_element(By.XPATH, '//*/em[@class=" x-btn-arrow"]').click()
    WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@class="x-date-mp"]/table/tbody/tr/td/a[text() = "' + datemonth + '"]'))).click()
    WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@class="x-date-mp"]/table/tbody/tr/td/a[text() = "' + dateyear + '"]'))).click()
    driver.find_element(By.CLASS_NAME, "x-date-mp-ok").click()
    WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@class="x-date-picker x-unselectable"]/table/tbody/tr/td/table/tbody/tr/td/a/em/span[text() = "' + str(dateday) + '"]'))).click()
    time.sleep(1)
    #Time in
    driver.find_element(By.ID, "TimeIn").send_keys(timeinvaluefinal[i])
    time.sleep(1)
    #Time out
    driver.find_element(By.ID, "TimeOut").send_keys(timeoutvaluefinal[i])
    time.sleep(1)
    #Input Remarks
    driver.find_element(By.XPATH, '//*[@name="Comment"]').click()
    driver.find_element(By.XPATH, '//*[@name="Comment"]').send_keys(filingremarks)
    time.sleep(1)
    #Save
    driver.find_element(By.XPATH, '//*[@id="maintenance-save"]/tbody/tr[2]/td[2]/em/button[text()="Save"]').click()
    driver.find_element(By.XPATH, '//body').click()
    WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@class="x-btn   x-btn-noicon "]/tbody/tr[2]/td[2]/em/button[text()="Yes"]'))).click()
#driver.quit()