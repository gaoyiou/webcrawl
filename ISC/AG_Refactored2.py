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
ChromeDriver = 'C:\\Python27\\web_driver\\chromedriver_win32\\chromedriver.exe'
url = "https://www.isc.ca/Pages/default.aspx"
input_workbook = 'D:\\web_craw\\webcrawl\\ISC\\PIN.xlsx'
output_workbook = 'D:\\web_craw\\webcrawl\\ISC\\property_info_.xlsx'


#create output xlsx
write_workbook = xlsxwriter.Workbook(output_workbook)
write_sheet = write_workbook.add_worksheet("Sheet 1")
write_sheet.write(0,0,"Parcel Number")                
write_sheet.write(0,1,"Share")
write_sheet.write(0,2,"Land Description")
write_sheet.write(0,3,"Municipality")
write_sheet.write(0,4,"Title Number")
write_sheet.write(0,5,"Owners")


#login on ISC website
driver = webdriver.Chrome(ChromeDriver)
driver.get(url)
HomeSignIn = driver.find_element_by_xpath ('//*[@id="sign_in"]/div/div/div/div/ul/li[1]/a')
HomeSignIn.click()
driver.switch_to_frame(driver.find_element_by_name('landIFrame'))
Username = driver.find_element_by_xpath ('//*[@id="c_username"]')
Username.click()
Username.send_keys('gaoyiou')
Password = driver.find_element_by_xpath ('//*[@id="c_Password"]')
Password.click()
Password.send_keys('Gyo910701')
ClientNumber = driver.find_element_by_xpath ('//*[@id="c_clientNumber"]')
ClientNumber.click()
ClientNumber.send_keys('132392707')
SignIn = driver.find_element_by_xpath ('//*[@id="c_btnSignIn"]')
SignIn.click()
QuickSearch = driver.find_element_by_xpath ('//*[@id="side_menu"]/div/div/div/div/ul/li[4]/a')
QuickSearch.click()
driver.switch_to_frame(driver.find_element_by_name('landIFrame'))


#read LINC from xlsx
read_workbook = xlrd.open_workbook(input_workbook)
read_sheet = read_workbook.sheet_by_index(0)
#define the row to write record
row = 1

for i in range (1,read_sheet.nrows):
    LINC_input = read_sheet.cell(i,0).value

#Search LINC number on SPIN2
    SearchBy = driver.find_element_by_xpath ('//*[@id="quicksearchby"]')
    SearchBy.send_keys('Parcel Number')
    ParcelNumber = driver.find_element_by_xpath ('//*[@id="ParcelNumber"]')
    ParcelNumber.click()
    ParcelNumber.clear()
    ParcelNumber.send_keys(LINC_input)
    driver.find_element_by_xpath ('//*[@id="main"]/form/table[3]/tbody/tr[3]/td/table/tbody/tr/td[2]/a/span').click()

    try:
        parcelpage = WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.ID, "main")))
    except TimeoutException:
        print "Timed out waiting for page to load"
#write to the output xlsx
    
    NumberResults = driver.find_element_by_xpath ('//*[@class="land_table_body_row land_search_recordcount"]')
    n = int(NumberResults.text.split(' ')[0])

    write_sheet.write(row,0,driver.find_element_by_xpath ('//*[@id="main"]/table[3]/tbody/tr[11]/td/table/tbody/tr/td[1]/table/tbody/tr/td/table/tbody/tr/td[1]/table/tbody/tr/td/table/tbody/tr/td[1]').text)
    write_sheet.write(row,1,driver.find_element_by_xpath ('//*[@id="main"]/table[3]/tbody/tr[7]/td/table/tbody/tr/td[1]/table/tbody/tr/td/table/tbody/tr/td[5]').text)
    write_sheet.write(row,2,driver.find_element_by_xpath ('//*[@id="main"]/table[3]/tbody/tr[3]/td/table/tbody/tr/td[1]').text)
    write_sheet.write(row,3,driver.find_element_by_xpath ('//*[@id="main"]/table[3]/tbody/tr[11]/td/table/tbody/tr/td[1]/table/tbody/tr/td/table/tbody/tr/td[5]').text)
    title = driver.find_element_by_xpath ('//*[@id="main"]/table[3]/tbody/tr[7]/td/table/tbody/tr/td[1]/table/tbody/tr/td/table/tbody/tr/td[3]/table/tbody/tr/td/table/tbody/tr/td[3]').text
    owner = driver.find_element_by_xpath ('//*[@id="main"]/table[3]/tbody/tr[5]/td/table/tbody/tr/td[1]').text
    write_sheet.write(row,4,title)
    write_sheet.write(row,5,owner)
    
    

    if n > 1:
        for m in range (0, n-1):
            title = title +"; "+ driver.find_element_by_xpath ('//*[@id="main"]/table['+str(4+m)+']/tbody/tr[10]/td/table/tbody/tr/td[1]/table/tbody/tr/td/table/tbody/tr/td[3]/table/tbody/tr/td/table/tbody/tr/td[3]').text
            owner = owner +"; "+ driver.find_element_by_xpath ('//*[@id="main"]/table['+str(4+m)+']/tbody/tr[8]/td/table/tbody/tr/td[1]').text
            write_sheet.write(row,4,title)
            write_sheet.write(row,5,owner)
    
    row = row +1
    
    #back to search another property
    driver.find_element_by_xpath ('//*[@id="main"]/table['+str(4+n)+']/tbody/tr[3]/td/table/tbody/tr/td[2]').click()

print "Done"
write_workbook.close()
driver.close()