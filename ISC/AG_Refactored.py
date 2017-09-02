import xlrd
import xlsxwriter
import selenium
import time
from selenium import webdriver
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By

class ISCCrawler(object):

    def __init__(self):
        # set up static parameters
        self.xPaths = [
            "//*[@id=\"main\"]/table[%d]/tbody/tr[%d]/td/table/tbody/tr/td[1]/table/tbody/tr/td/table/tbody/tr/td[1]/table/tbody/tr/td/table/tbody/tr/td[1]",
            "//*[@id=\"main\"]/table[%d]/tbody/tr[%d]/td/table/tbody/tr/td[1]/table/tbody/tr/td/table/tbody/tr/td[5]",
            "//*[@id=\"main\"]/table[%d]/tbody/tr[%d]/td/table/tbody/tr/td[1]",
            "//*[@id=\"main\"]/table[%d]/tbody/tr[%d]/td/table/tbody/tr/td[1]/table/tbody/tr/td/table/tbody/tr/td[5]",
            "//*[@id=\"main\"]/table[%d]/tbody/tr[%d]/td/table/tbody/tr/td[1]/table/tbody/tr/td/table/tbody/tr/td[3]/table/tbody/tr/td/table/tbody/tr/td[3]",
            "//*[@id=\"main\"]/table[%d]/tbody/tr[%d]/td/table/tbody/tr/td[1]"
        ]

        self.chromeDriver = 'C:\\Python27\\web_driver\\chromedriver_win32\\chromedriver.exe'
        self.url = "https://www.isc.ca/Pages/default.aspx"
        #read LINC from xlsx
        self.readSheet = xlrd.open_workbook('D:\\web_craw\\webcrawl\\ISC\\PIN.xlsx').sheet_by_index(0)
        self.output_workbook = 'D:\\web_craw\\webcrawl\\ISC\\property_info.xlsx'

        # define parameters those will be set and used later
        self.writeSheet = None
        self.writeWorkBook = None
        self.driver = None
        
    def creatWorkSheet(self):
        #create output xlsx
        write_workbook = xlsxwriter.Workbook(self.output_workbook)
        write_sheet = write_workbook.add_worksheet("Sheet 1")
        head = ["Parcel Number", "Share", "Land Description", "Municipality", "Title Number", "Owners"]
        write_sheet.write_row('A1', head)
        
        self.writeWorkBook = write_workbook
        self.writeSheet = write_sheet
        return self

    def login(self):
        #login on ISC website
        driver = webdriver.Chrome(self.chromeDriver)
        driver.get(self.url)
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

        self.driver = driver
        return self

    def searchLINC(self, LINCinput):

            #Search LINC number on SPIN2
            SearchBy = self.driver.find_element_by_xpath ('//*[@id="quicksearchby"]')
            SearchBy.send_keys('Parcel Number')
            ParcelNumber = self.driver.find_element_by_xpath ('//*[@id="ParcelNumber"]')
            ParcelNumber.click()
            ParcelNumber.clear()
            ParcelNumber.send_keys(LINCinput)
            self.driver.find_element_by_xpath ('//*[@id="main"]/form/table[3]/tbody/tr[3]/td/table/tbody/tr/td[2]/a/span').click()

            try:
                WebDriverWait(self.driver, 30).until(EC.presence_of_element_located((By.ID, "main")))
            except TimeoutException:
                print "Timed out waiting for page to load"

    def driverRead(self, xPath):
        return self.driver.find_element_by_xpath(xPath).text

    def writeRow(self, tableIndex, rowIndex, moreThanN = False):
        row = []
        secondaryTableIndex = [11, 7 ,3 ,11 ,7, 5]

        for i in range(0, 6):
            row.append(self.driverRead(self.xPaths[i] % (tableIndex, secondaryTableIndex[i] if not moreThanN else secondaryTableIndex[i] + 3)))

        self.writeSheet.write_row("A%d" % rowIndex, row)

    def crawer(self):

        #define the row to write record
        row = 2

        for i in range (1, self.readSheet.nrows):
            LINCinput = self.readSheet.cell(i, 0).value

            self.searchLINC(LINCinput)
        #write to the output xlsx
            
            NumberResults = self.driver.find_element_by_xpath ('//*[@class="land_table_body_row land_search_recordcount"]')
            n = int(NumberResults.text.split(' ')[0])
         
            self.writeRow(3, row)
            
            row = row + 1

            if n > 1:
                for m in range (1, n):
                    self.writeRow(3+m, row, True)
                    row = row +1

            #back to search another property
            self.driver.find_element_by_xpath ('//*[@id="main"]/table['+str(4+n)+']/tbody/tr[3]/td/table/tbody/tr/td[2]').click()

        print "Done"
        self.writeWorkBook.close()
        self.driver.close()


if __name__ == '__main__':
    ISCCrawler().creatWorkSheet().login().crawer()