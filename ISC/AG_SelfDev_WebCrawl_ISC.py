import xlrd
import xlsxwriter
import selenium
import time
from selenium import webdriver
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By

#set parameter
ChromeDriver = "/Users/yiougao/Python/chrome_driver/chromedriver"
url = "https://www.isc.ca/Pages/default.aspx"
input_workbook = '/Users/yiougao/Python/Web_Crawl/ISC/PIN.xlsx'
output_workbook = '/Users/yiougao/Python/Web_Crawl/ISC/property_info.xlsx'


#create output xlsx
write_workbook = xlsxwriter.Workbook(output_workbook)
write_sheet = write_workbook.add_worksheet("Sheet 1")
write_sheet.write(0,0,"Parcel Number")
write_sheet.write(0,1,"Type")
write_sheet.write(0,2,"Status")
write_sheet.write(0,3,"Class")                
write_sheet.write(0,4,"Total Units")
write_sheet.write(0,5,"LLD")
write_sheet.write(0,6,"Parcel Value")
write_sheet.write(0,7,"Mucinipality")
write_sheet.write(0,8,"Title Type")
write_sheet.write(0,9,"Last Amendment Date")
write_sheet.write(0,10,"Commodity")
write_sheet.write(0,11,"Owners")
write_sheet.write(0,12,"Share")
write_sheet.write(0,13,"Lock Information")


#login on ISC website
driver = webdriver.Chrome(ChromeDriver)
driver.get(url)
HomeSignIn = driver.find_element_by_xpath ('//*[@id="sign_in"]/div/div/div/div/ul/li[1]/a')
HomeSignIn.click()
driver.switch_to_frame(driver.find_element_by_name('landIFrame'))
Username = driver.find_element_by_xpath ('//*[@id="c_username"]')
Username.click()
Username.send_keys('username')
Password = driver.find_element_by_xpath ('//*[@id="c_Password"]')
Password.click()
Password.send_keys('password')
ClientNumber = driver.find_element_by_xpath ('//*[@id="c_clientNumber"]')
ClientNumber.click()
ClientNumber.send_keys('clientid')
SignIn = driver.find_element_by_xpath ('//*[@id="c_btnSignIn"]')
SignIn.click()
Search = driver.find_element_by_xpath ('//*[@id="side_menu"]/div/div/div/div/ul/li[5]/a')
Search.click()
driver.switch_to_frame(driver.find_element_by_name('landIFrame'))
SearchBy = driver.find_element_by_xpath ('//*[@id="searchby"]')
SearchBy.click()
SearchBy.send_keys('Parcel Number')


#read LINC from xlsx
read_workbook = xlrd.open_workbook(input_workbook)
read_sheet = read_workbook.sheet_by_index(0)
for i in range (1,2):
    LINC_input = read_sheet.cell(i,0).value

#Search LINC number on SPIN2
    ParcelNumber = driver.find_element_by_xpath ('//*[@id="ParcelNumber"]')
    ParcelNumber.click()
    ParcelNumber.clear()
    ParcelNumber.send_keys(LINC_input)
    driver.find_element_by_xpath ('//*[@id="AsOfDateSearchOption"]').click()
    driver.find_element_by_xpath ('//*[@id="main"]/table/tbody/tr[3]/td/table/tbody/tr/td[2]/a/span').click()
    try:
        parcelpage = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, '//*[@id="main"]/table[1]/tbody/tr[1]/td/table/tbody/tr/td[2]')))
    except TimeoutException:
        print "Timed out waiting for page to load"
#write to the output xlsx
    #driver.switch_to_frame(driver.find_element_by_name('landIFrame'))
    
    test = driver.find_element_by_xpath ('//*[@id="c1fe70e39c01450dbb6c5029cf7d2ad7"]/table/tbody/tr[1]/td/table/tbody/tr/td[1]')
    print test.text
    '''write_sheet.write(i,0,driver.find_element_by_xpath ('//*[@id="c1fe70e39c01450dbb6c5029cf7d2ad7"]/table/tbody/tr[1]/td/table/tbody/tr/td[1]').text)
    write_sheet.write(i,1,driver.find_element_by_xpath ('//*[@id="c1fe70e39c01450dbb6c5029cf7d2ad7"]/table/tbody/tr[1]/td/table/tbody/tr/td[2]').text)
    write_sheet.write(i,2,driver.find_element_by_xpath ('//*[@id="c1fe70e39c01450dbb6c5029cf7d2ad7"]/table/tbody/tr[3]/td/table/tbody/tr/td[1]').text)
    write_sheet.write(i,3,driver.find_element_by_xpath ('').text)
    write_sheet.write(i,4,driver.find_element_by_xpath ('').text)
    write_sheet.write(i,5,driver.find_element_by_xpath ('').text)
    write_sheet.write(i,6,driver.find_element_by_xpath ('').text)

print "Done"
write_workbook.close()
driver.close()
'''
