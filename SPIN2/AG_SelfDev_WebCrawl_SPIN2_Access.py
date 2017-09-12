import selenium
import pyodbc
from selenium import webdriver

#set parameter
ChromeDriver = "C:\\Python27\\web_driver\\chromedriver_win32\\chromedriver.exe"
url = "https://alta.registries.gov.ab.ca/spinii/logon.aspx"
database_location = 'D:/web_craw/webcrawl/SPIN2/property_info.mdb'
DBdriver = '{Microsoft Access Driver (*.mdb)}'
                  
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

#connect to database and create a cursor
conn = pyodbc.connect('DRIVER={};DBQ={}'.format(DBdriver, database_location))
cursor = conn.cursor()

#read LINC from access database
cursor.execute("select LINC from property_info")
lincs = cursor.fetchall()
for linc in lincs:

#Search LINC number on SPIN2
    LINC = driver.find_element_by_xpath ('//*[@id="TitleLinc_ctlLinc_txtLincNumber"]')
    LINC.click()
    LINC.clear()
    LINC.send_keys(linc[0])
    Surface = driver.find_element_by_xpath ('//*[@id="Table1"]/tbody/tr/td/table/tbody/tr/td/table/tbody/tr/td/span[2]/label')
    Surface.click()
    SearchLinc = driver.find_element_by_xpath ('//*[@id="TitleLinc_cmdSubmitSearch"]')
    SearchLinc.click()
    
#write to the database
    cursor.execute("""
    UPDATE property_info
    SET
    Title_Number = ?,
    Type = ?,
    Short_Legal = ?,
    Rights = ?,
    Registration_Date = ?,
    Change_Cancel_Date = ?
    WHERE LINC = ?
    """,
    driver.find_element_by_xpath ('//*[@id="TitleResult_dgResults"]/tbody/tr[2]/td[3]').text,
    driver.find_element_by_xpath ('//*[@id="TitleResult_dgResults"]/tbody/tr[2]/td[4]').text,
    driver.find_element_by_xpath ('//*[@id="TitleResult_dgResults"]/tbody/tr[2]/td[6]').text,
    driver.find_element_by_xpath ('//*[@id="TitleResult_dgResults"]/tbody/tr[2]/td[7]').text,
    driver.find_element_by_xpath ('//*[@id="TitleResult_dgResults_lblRegDate_0"]').text,
    driver.find_element_by_xpath ('//*[@id="TitleResult_dgResults_lblChangeDate_0"]').text,
    str(linc[0]))

    print "finish for "+str(linc[0])
    conn.commit()

print "Done"
driver.close()
