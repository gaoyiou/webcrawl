# import sys
# sys.path.append("/usr/local/lib/python2.7/site-packages")
import xlrd
import xlsxwriter
import selenium
from selenium import webdriver

#set parameter
ChromeDriver = "/Users/yiougao/Python/chrome_driver/chromedriver"
url = "https://alta.registries.gov.ab.ca/spinii/logon.aspx"
input_workbook = '/Users/yiougao/Python/Web_Crawl/SPIN2/PIN.xlsx'
output_workbook = '/Users/yiougao/Python/Web_Crawl/SPIN2/property_info.xlsx'

#create output xlsx
write_workbook = xlsxwriter.Workbook(output_workbook)
write_sheet = write_workbook.add_worksheet("Sheet 1")
write_sheet.write(0,0,"LINC")
write_sheet.write(0,1,"Title Number")
write_sheet.write(0,2,"Type")
write_sheet.write(0,3,"Short Legal")                
write_sheet.write(0,4,"Rights")
write_sheet.write(0,5,"Registration Date")
write_sheet.write(0,6,"Change/Cancel Date")
                  
#login on SPIN2 website
driver = webdriver.Chrome(ChromeDriver)
driver.get(url)
GuestLogin = driver.find_element_by_xpath ('//*[@id="uctrlLogon_cmdLogonGuest"]')
GuestLogin.click()
Disclaimer = driver.find_element_by_xpath ('//*[@id="cmdYES"]')
Disclaimer.click()
Search = driver.find_element_by_xpath ('//*[@id="MenuBar_pnlMenuItems"]/a[3]/img')
Search.click()
TitleDocs = driver.find_element_by_xpath ('//*[@id="SelectSearch1_tblSearchItems"]/tbody/tr[1]/td[1]/table/tbody/tr[1]/td[2]/a/font')
TitleDocs.click()
LINCSearch = driver.find_element_by_xpath ('//*[@id="SelectSearch_tblSearchItems"]/tbody/tr[3]/td[2]/table/tbody/tr[1]/td[2]/a/font')
LINCSearch.click()

#read LINC from xlsx
read_workbook = xlrd.open_workbook(input_workbook)
read_sheet = read_workbook.sheet_by_index(0)
for i in range (1,read_sheet.nrows):
    LINC_input = read_sheet.cell(i,0).value

#Search LINC number on SPIN2
    LINC = driver.find_element_by_xpath ('//*[@id="TitleLinc_ctlLinc_txtLincNumber"]')
    LINC.click()
    LINC.clear()
    LINC.send_keys(LINC_input)
    Surface = driver.find_element_by_xpath ('//*[@id="Table1"]/tbody/tr/td/table/tbody/tr/td/table/tbody/tr/td/span[2]/label')
    Surface.click()
    SearchLinc = driver.find_element_by_xpath ('//*[@id="TitleLinc_cmdSubmitSearch"]')
    SearchLinc.click()
#write to the output xlsx
    write_sheet.write(i,0,driver.find_element_by_xpath ('//*[@id="TitleResult_dgResults_lblLINC_0"]').text)
    write_sheet.write(i,1,driver.find_element_by_xpath ('//*[@id="TitleResult_dgResults"]/tbody/tr[2]/td[3]').text)
    write_sheet.write(i,2,driver.find_element_by_xpath ('//*[@id="TitleResult_dgResults"]/tbody/tr[2]/td[4]').text)
    write_sheet.write(i,3,driver.find_element_by_xpath ('//*[@id="TitleResult_dgResults"]/tbody/tr[2]/td[6]').text)
    write_sheet.write(i,4,driver.find_element_by_xpath ('//*[@id="TitleResult_dgResults"]/tbody/tr[2]/td[7]').text)
    write_sheet.write(i,5,driver.find_element_by_xpath ('//*[@id="TitleResult_dgResults_lblRegDate_0"]').text)
    write_sheet.write(i,6,driver.find_element_by_xpath ('//*[@id="TitleResult_dgResults_lblChangeDate_0"]').text)

print "Done"
write_workbook.close()
driver.close()
