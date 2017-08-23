import requests as req
import xlrd
import csv
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from bs4 import BeautifulSoup
import time


workbook = xlrd.open_workbook("imo.xlsx")
sheet = workbook.sheet_by_index(0)

imos = []
for row in range(1, sheet.nrows):
	
	cols = sheet.row_values(row)
	
	imo = int(cols[0])
	imos.append(imo)



options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")

driver = webdriver.Chrome("/home/alexkott/Documents/YouDo/ships/chromedriver/chromedriver", chrome_options=options)



url = "https://www.korabel.ru/index.php"




for imo in imos:
	driver.get("https://www.korabel.ru/search?s=fleet")
	driver.find_element_by_id("headSearch_Input").click()
	driver.find_element_by_id("headSearch_Input").send_keys(imo)
	time.sleep(0.5)

	els = driver.find_elements_by_class_name("sItem")
	print(els)

	break



