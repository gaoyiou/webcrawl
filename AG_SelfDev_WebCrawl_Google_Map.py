import sys
sys.path.append("/usr/local/lib/python2.7/site-packages")
from selenium import webdriver

#set parameter
ChromeDriver = "/Users/yiougao/Python/chrome_driver/chromedriver"
url = "https://google.ca/maps"
address_input = raw_input("Your address is: ")


driver = webdriver.Chrome(ChromeDriver)
driver.get(url)

address = driver.find_element_by_xpath ('//*[@id="searchboxinput"]')
address.click()
address.send_keys(address_input)
search_address = driver.find_element_by_xpath ('//*[@id="searchbox-searchbutton"]')
search_address.click()
