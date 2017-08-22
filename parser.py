import xlrd
import csv
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from bs4 import BeautifulSoup
import time


options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")

driver = webdriver.Chrome("/home/alexkott/Documents/YouDo/shop-parsing/selenium_test/chromedriver", chrome_options=options)
driver.get("http://www.shippingexplorer.net/ru/")
driver.find_element_by_class_name("userlogin").click()
driver.find_element_by_id("UsernameOrEMail").send_keys("alexey.kott@gmail.com")
driver.find_element_by_id("Password").send_keys("4815162342")
driver.find_element_by_class_name("btn-primary").click()
driver.get("http://www.shippingexplorer.net/ru/ships")

site = "http://www.shippingexplorer.net{}"

workbook = xlrd.open_workbook("imo.xlsx")
sheet = workbook.sheet_by_index(0)

header = []
flag = 0 # для заголовка

l = 0


imos = []
for row in range(1, sheet.nrows):
	
	cols = sheet.row_values(row)
	
	imo = int(cols[0])
	imos.append(imo)


for imo in imos:
	driver.find_element_by_id("Name").send_keys(imo)
	driver.find_element_by_id("Name").send_keys(Keys.ENTER)
	time.sleep(1)
	
	content = driver.find_elements_by_class_name("odd")[0]
	soup = BeautifulSoup(content.get_attribute("innerHTML"), "lxml")
	try:
		a = soup.find_all("a")[0]
	except Exception as e:
		print(str(e))
		print(soup)
		print("Error IMO {}".format(imo))
		driver.get("http://www.shippingexplorer.net/ru/ships")
		time.sleep(1)
		continue
	driver.get(site.format(a['href']))

	page = BeautifulSoup(driver.page_source, "lxml")

	name = page.h1.contents[0]
	name = name.strip()

	line = dict()
	header.append("Название")
	line['Название'] = name
	sheets = page.find_all(class_="infosheet")
	for sheet in sheets:
		for li in sheet.findAll("li"):
			spans = li.findAll("span")
			field = str(spans[0].contents[0])
			if flag == 0:
				header.append(field)
			try:
				line[field] = str(spans[1].contents[0])
			except:
				line[field] = ' '
	
	if flag == 0:
		with open("result.csv", 'a') as f:
			writer = csv.writer(f, delimiter=';')
			writer.writerow(header)
		f.close()

	#print(line)
	res_line = []
	for x in header:
		if x in line:
			res_line.append(line[x])
		else:
			res_line.append('-')


	with open("result.csv", 'a') as f:
		writer = csv.writer(f, delimiter=';')
		writer.writerow(res_line)
	f.close()


	flag = 1
	l = l+1
	print(l)

	driver.get("http://www.shippingexplorer.net/ru/ships")
	time.sleep(1)
driver.close()
	