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
#login
clientcode = 'YONDU'
employeecode = '22-02762'
password = 'Aa051498!'
#OB filing
filingremarks = 'No Time In/Time Out'
#Change shift schedule filing
changeshiftremarks = 'Change Shift'
daytype = 'Regular Day'
excel_files = 'D:\RPA\Yondu Timesheet_Novemberr 1-15 2021_Monica Cho.xlsx'

####FROM EXCEL####
timeinvalues = []
timeoutvalues = []
datevalues = []
changetimein = []
changetimeout = []

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

####WEB AUTOMATION####
url = 'https://ess.payroll.ph/app'
driver = webdriver.Chrome()
driver.get(url)

driver.find_element(By.ID, "loginClientCode").send_keys(clientcode)
driver.find_element(By.ID, "loginEmployeeCode").send_keys(employeecode)
driver.find_element(By.ID, "loginPassword").send_keys(password)
driver.find_element(By.ID, "loginForm").click()
print("Login DONE")
#Click OB FILING
driver.implicitly_wait(15)
driver.find_element(By.XPATH, '//*[@class="x-portlet-button-img"][@src="/images/flat-icons/clock2.png"]').click()
for i in range(iteration):
    ###VARIABLES####
    #Click NEW
    time.sleep(5)
    driver.find_element(By.XPATH, '//*[@class=" x-btn-text icon-new"][text()="New"]').click()
    #Client
    driver.find_element(By.XPATH, '//*[@src="/images/default/s.gif"][@class="x-form-trigger x-form-arrow-trigger"]').click()
    WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@class="x-layer x-combo-list "]/div/div[normalize-space(text()) = "Work from Home"]'))).click()
    #Date
    driver.find_element(By.XPATH, '//div/div/img[@src="/images/default/s.gif"][@class="x-form-trigger x-form-date-trigger"]').click()
    driver.find_element(By.XPATH, '//*/em[@class=" x-btn-arrow"]').click()
    WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@class="x-date-mp"]/table/tbody/tr/td/a[text() = "' + datefinal[i].strftime("%b") + '"]'))).click()
    WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@class="x-date-mp"]/table/tbody/tr/td/a[text() = "' + datefinal[i].strftime("%Y") + '"]'))).click()
    driver.find_element(By.CLASS_NAME, "x-date-mp-ok").click()
    WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@class="x-date-picker x-unselectable"]/table/tbody/tr/td/table/tbody/tr/td/a/em/span[text() = "' + str(datefinal[i].day) + '"]'))).click()
    time.sleep(1)
    #Time in
    driver.find_element(By.ID, "TimeIn").send_keys(timeinvalues[i].strftime("%I:%M %p"))
    time.sleep(1)
    #Time out
    driver.find_element(By.ID, "TimeOut").send_keys(timeoutvalues[i].strftime("%I:%M %p"))
    time.sleep(1)
    #Input Remarks
    driver.find_element(By.XPATH, '//*[@name="Comment"]').click()
    driver.find_element(By.XPATH, '//*[@name="Comment"]').send_keys(filingremarks)
    time.sleep(1)
    #Save
    driver.find_element(By.XPATH, '//*[@id="maintenance-save"]/tbody/tr[2]/td[2]/em/button[text()="Save"]').click()
    driver.find_element(By.XPATH, '//body').click()
    WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@class="x-btn   x-btn-noicon "]/tbody/tr[2]/td[2]/em/button[text()="Yes"]'))).click()
print("OB Filing DONE")
#CLICK Change Shift Schedule Filing
time.sleep(5)
driver.find_element(By.XPATH, '//*/li[@id="tabpanel__tab-obfiling"]/a[@class="x-tab-strip-close"]').click()
driver.find_element(By.XPATH, '//*/div[@id="portal-second-column"]/div/div[2]/div[1]/div/div/div/div[4]/center/img[@src="/images/flat-icons/week.png"]').click()
for i in range(iteration):
    ###VARIABLES####
    #Click NEW
    time.sleep(5)
    driver.find_element(By.XPATH, '//*[@class=" x-btn-text icon-new"][text()="New"]').click()
    #Date
    driver.find_element(By.XPATH, '//*/div/div/img[@src="/images/default/s.gif"][@class="x-form-trigger x-form-date-trigger"]').click()
    driver.find_element(By.XPATH, '//*/em[@class=" x-btn-arrow"]').click()
    WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@class="x-date-mp"]/table/tbody/tr/td/a[text() = "' + datefinal[i].strftime("%b") + '"]'))).click()
    WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@class="x-date-mp"]/table/tbody/tr/td/a[text() = "' + datefinal[i].strftime("%Y") + '"]'))).click()
    driver.find_element(By.CLASS_NAME, "x-date-mp-ok").click()
    WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@class="x-date-picker x-unselectable"]/table/tbody/tr/td/table/tbody/tr/td/a/em/span[text() = "' + str(datefinal[i].day) + '"]'))).click()
    time.sleep(1)
    #New Schedule
    driver.find_element(By.XPATH, '//*/form/div[2]/div[1]/div/input[2]').send_keys(str(int(timeinfinal[i].strftime("%I"))) + timeinfinal[i].strftime("%p") + "-" + str(int(timeoutfinal[i].strftime("%I"))) + timeoutfinal[i].strftime("%p"))
    time.sleep(1)
    #Day Type
    driver.find_element(By.XPATH, '//*/form/div[3]/div[1]/div/input[2]').send_keys(daytype)
    time.sleep(1)
    #Input Remarks
    driver.find_element(By.XPATH, '//*[@name="Remarks"]').click()
    driver.find_element(By.XPATH, '//*[@name="Remarks"]').send_keys(changeshiftremarks)
    time.sleep(1)
    #Save
    driver.find_element(By.XPATH, '//*[@id="maintenance-save"]/tbody/tr[2]/td[2]/em/button[text()="Save"]').click()
    driver.find_element(By.XPATH, '//body').click()
    WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@class="x-btn   x-btn-noicon "]/tbody/tr[2]/td[2]/em/button[text()="Yes"]'))).click()
print("Change Shift DONE")