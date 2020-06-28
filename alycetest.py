from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from datetime import datetime
import ExcelFunctions
import openpyxl
import platform
import time

driver = webdriver.Chrome(executable_path="C:\Drivers\chromedriver.exe")
#driver = webdriver.Edge(executable_path="C:\Drivers\msedgedriver.exe")
#driver = webdriver.Firefox(executable_path="C:\Drivers\geckodriver.exe")
#driver = webdriver.Ie(executable_path="C:\Drivers\IEDriverServer.exe")
#driver = webdriver.Opera(executable_path="C:\Drivers\operadriver.exe")

path = "C://Users//Mikhail//PycharmProjects//Alyce//CheckListAlyce.xlsx"

#write result at next column
book = openpyxl.load_workbook(path)
sheet = book.active
num = sheet.max_column + 1

#Header of checklist
ExcelFunctions.writeData(path, "List1", 1, num, platform.platform())
ExcelFunctions.writeData(path, "List1", 2, num, driver.capabilities['browserName'])
ExcelFunctions.writeData(path, "List1", 3, num, driver.capabilities['browserVersion'])
#ExcelFunctions.writeData(path, "List1", 3, num, driver.capabilities['version']) #case opera
ExcelFunctions.writeData(path, "List1", 4, num, datetime.now().strftime('%d.%m.%Y_%H.%M.%S'))

#Opens main page
driver.maximize_window()
driver.get("http://hrtest.alycedev.com/")
ExcelFunctions.writeData(path, "List1", 6, num, "Pass")

#get screenshot of main page
driver.implicitly_wait(5)
driver.save_screenshot("C://Users//Mikhail//PycharmProjects//Alyce//homepage_default.png")
ExcelFunctions.writeData(path, "List1", 7, num, "Done")
time.sleep(5)

#one apple per minute
Jonathan = driver.find_element_by_xpath('/html/body/div/div[2]/div[2]/div[1]/div/section[1]/ul/li[1]/div/span/button')

Jonathan.click()
time.sleep(6)
if Jonathan.click():
    ExcelFunctions.writeData(path, "List1", 8, num, "Fail")
else:
    time.sleep(54)
    Jonathan.click()
    ExcelFunctions.writeData(path, "List1", 8, num, "Pass")

time.sleep(10)
driver.quit()
