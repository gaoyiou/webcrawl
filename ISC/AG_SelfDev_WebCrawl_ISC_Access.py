import pyodbc
import selenium
import time
from selenium import webdriver
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By

#set parameter
ChromeDriver = "C:\\Python27\\web_driver\\chromedriver_win32\\chromedriver.exe"
url = "https://www.isc.ca/Pages/default.aspx"
database_location = 'D:/web_craw/webcrawl/ISC/property_info.mdb'
DBdriver = '{Microsoft Access Driver (*.mdb)}'

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

#connect to database and create a cursor
conn = pyodbc.connect('DRIVER={};DBQ={}'.format(DBdriver, database_location))
cursor = conn.cursor()

#read LINC from access database
cursor.execute("select Parcel_Number from property_info")
lincs = cursor.fetchall()
for linc in lincs:

#Search LINC number on SPIN2
    SearchBy = driver.find_element_by_xpath ('//*[@id="quicksearchby"]')
    SearchBy.send_keys('Parcel Number')
    ParcelNumber = driver.find_element_by_xpath ('//*[@id="ParcelNumber"]')
    ParcelNumber.click()
    ParcelNumber.clear()
    ParcelNumber.send_keys(linc[0])
    driver.find_element_by_xpath ('//*[@id="main"]/form/table[3]/tbody/tr[3]/td/table/tbody/tr/td[2]/a/span').click()

    try:
        parcelpage = WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.ID, "main")))
    except TimeoutException:
        print "Timed out waiting for page to load"

    #write to database
    NumberResults = driver.find_element_by_xpath ('//*[@class="land_table_body_row land_search_recordcount"]')
    n = int(NumberResults.text.split(' ')[0])
 
    cursor.execute("""
    UPDATE property_info
    SET
    Share = ?,
    Land_Description = ?,
    Municipality = ?,
    Title_Number = ?,
    Owner = ?
    WHERE Parcel_Number = ?
    """,
    driver.find_element_by_xpath ('//*[@id="main"]/table[3]/tbody/tr[7]/td/table/tbody/tr/td[1]/table/tbody/tr/td/table/tbody/tr/td[5]').text,
    driver.find_element_by_xpath ('//*[@id="main"]/table[3]/tbody/tr[3]/td/table/tbody/tr/td[1]').text,
    driver.find_element_by_xpath ('//*[@id="main"]/table[3]/tbody/tr[11]/td/table/tbody/tr/td[1]/table/tbody/tr/td/table/tbody/tr/td[5]').text,
    driver.find_element_by_xpath ('//*[@id="main"]/table[3]/tbody/tr[7]/td/table/tbody/tr/td[1]/table/tbody/tr/td/table/tbody/tr/td[3]/table/tbody/tr/td/table/tbody/tr/td[3]').text,
    driver.find_element_by_xpath ('//*[@id="main"]/table[3]/tbody/tr[5]/td/table/tbody/tr/td[1]').text,
    str(linc[0]))
                   
    conn.commit()

    if n > 1:
        for m in range (0, n-1):
            cursor.execute("""
            insert into property_info (Parcel_Number,Share,Land_Description,Municipality,Title_Number,Owner)
            values (?,?,?,?,?,?)
            """,
            driver.find_element_by_xpath ('//*[@id="main"]/table['+str(4+m)+']/tbody/tr[14]/td/table/tbody/tr/td[1]/table/tbody/tr/td/table/tbody/tr/td[1]/table/tbody/tr/td/table/tbody/tr/td[1]').text,
            driver.find_element_by_xpath ('//*[@id="main"]/table['+str(4+m)+']/tbody/tr[10]/td/table/tbody/tr/td[1]/table/tbody/tr/td/table/tbody/tr/td[5]').text,
            driver.find_element_by_xpath ('//*[@id="main"]/table['+str(4+m)+']/tbody/tr[6]/td/table/tbody/tr/td[1]').text,
            driver.find_element_by_xpath ('//*[@id="main"]/table['+str(4+m)+']/tbody/tr[14]/td/table/tbody/tr/td[1]/table/tbody/tr/td/table/tbody/tr/td[5]').text,
            driver.find_element_by_xpath ('//*[@id="main"]/table['+str(4+m)+']/tbody/tr[10]/td/table/tbody/tr/td[1]/table/tbody/tr/td/table/tbody/tr/td[3]/table/tbody/tr/td/table/tbody/tr/td[3]').text,
            driver.find_element_by_xpath ('//*[@id="main"]/table['+str(4+m)+']/tbody/tr[8]/td/table/tbody/tr/td[1]').text)
            conn.commit()

    #back to search another property
    driver.find_element_by_xpath ('//*[@id="main"]/table['+str(4+n)+']/tbody/tr[3]/td/table/tbody/tr/td[2]').click()

print "Done"
driver.close()
