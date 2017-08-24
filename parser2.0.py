import requests as req
import xlrd
import csv
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from bs4 import BeautifulSoup
import time
import re


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


error_imos = []

vessel = {
	'name'	:	'',
	'type'	:	'',
	'imo'	:	'',
	'mmsi'	:	'',
	'dwt'	:	'',
	'gt'	:	'',
	'me'	:	'',
	'year'	:	'',
	'number':	'',
	'state'	:	'',
	'flag'	:	'',
	'port'	:	'',
	'country':	'',
	'owner'	:	'',
	'owner_country' : '',
	'owner_city'	:	'',
	'address':	'',
	'email'	:	'',
	'phone'	:	'',
	'owner_site' : ''
}

with open('report.csv', 'w') as f:
	w = csv.DictWriter(f, vessel.keys(), delimiter = ';')
	w.writeheader()
	f.close()

for imo in imos:

	vessel = {
		'name'	:	'',
		'type'	:	'',
		'imo'	:	'',
		'mmsi'	:	'',
		'dwt'	:	'',
		'gt'	:	'',
		'me'	:	'',
		'year'	:	'',
		'number':	'',
		'state'	:	'',
		'flag'	:	'',
		'port'	:	'',
		'country':	'',
		'owner'	:	'',
		'owner_country' : '',
		'owner_city'	:	'',
		'address':	'',
		'email'	:	'',
		'phone'	:	'',
		'owner_site' : ''
	}

	driver.get("https://www.korabel.ru/search?s=fleet")
	driver.find_element_by_id("headSearch_Input").click()
	driver.find_element_by_id("headSearch_Input").send_keys(imo)
	time.sleep(2)


	response = driver.page_source

	soup = BeautifulSoup(response, "lxml")

	searchResult = soup.find_all(class_="searchResult")
	try:
		#print(searchResult[0].a['href'])
		link = searchResult[0].a['href']
	except:
		error_imos.append(imo)

	driver.get(link)

	vessel['imo'] = imo

	page = BeautifulSoup(driver.page_source, "lxml")
	try:
		name = page.find_all(class_="svh-topl")[0].h1.contents[0]
	except:
		name = ''
	vessel['name'] = name

	try:
		type_ = page.find_all(class_="ship_types")[0].a.contents[0]
	except:
		type_ = ''
	vessel['type'] = type_

	try:
		mmsi = re.findall(r'<li title="">MMSI: <span>\d*<\/span><\/li>', driver.page_source)[0]
		mmsi = re.findall(r'(?<=<span>)\d*', mmsi)[0]
	except:
		mmsi = ''
	vessel['mmsi'] = mmsi

	try:
		dwt = re.findall(r'<li title="Дедвейт \(т\)">DWT: <span>\d*<\/span><\/li>', driver.page_source)[0]
		dwt = re.findall(r'(?<=<span>)\d*', dwt)[0]
	except:
		dwt = ''
	vessel['dwt'] = dwt

	try:
		gt = re.findall(r'<li title="Валовая вместимость \(т\)">GT: <span>\d*<\/span><\/li>', driver.page_source)[0]
		gt = re.findall(r'(?<=<span>)\d*', gt)[0]
	except:
		gt = ''
	vessel['gt'] = gt

	try:
		me = re.findall(r'<li title="Марка ГД">ME: <span>[^<]*<\/span><\/li>', driver.page_source)[0]
		me = re.findall(r'(?<=<span>)[^<]*(?=<\/span>)', me)[0]
	except:
		me = ''
	vessel['me'] = me

	try:
		year = re.findall(r'<li title="дата постройки">Year built: <span>\d*<\/span><\/li>', driver.page_source)[0]
		year = re.findall(r'(?<=<span>)\d*(?=<\/span>)', year)[0]
	except:
		year = ''
	vessel['year'] = year

	try:
		number = re.findall(r'http:\/\/info.rs-head.spb.ru\/webFS\/regbook\/vessel\?fleet_id=\d*"', driver.page_source)[0]
		number = re.findall(r'(?<=fleet_id=)\d*(?=")', number)[0]
	except:
		number = ''
	vessel['number'] = number

	try:
		state = re.findall(r'<tr><td style="width:160px;">Состояние<\/td><td>[^<]*<\/td><\/tr>', driver.page_source)[0]
		state = re.findall(r'(?<=<td>)[^<]*(?=<\/td>)', state)[0]
	except:
		state = ''
	vessel['state'] = state

	try:
		flag = re.findall(r'<tr><td style="width:160px;">Флаг<\/td><td>[^<]*<\/td><\/tr>', driver.page_source)[0]
		flag = re.findall(r'(?<=<td>)[^<]*(?=<\/td>)', flag)[0]
	except:
		flag = ''
	vessel['flag'] = flag

	try:
		port = re.findall(r'<tr><td style="width:160px;">Порт приписки<\/td><td>[^<]*<\/td><\/tr>', driver.page_source)[0]
		port = re.findall(r'(?<=<td>)[^<]*(?=<\/td>)', port)[0]
	except:
		port = ''
	vessel['port'] = port

	try:
		owner = re.findall(r'<tr><td style="width:160px;">Владелец<\/td><td><a href="https:\/\/www.korabel.ru\/catalogue\/item_view\/\d*.html" target="_blank">[^<]*<\/a><\/td><\/tr>', driver.page_source)[0]
		owner_link = re.findall(r'(?<=href=")[^"]*(?=")', owner)[0]
		owner = re.findall(r'(?<=target="_blank">)[^<]*(?=<\/a>)', owner)[0]
	except:
		owner = ''
	vessel['owner'] = owner

	try:
		country = re.findall(r'<tr><td style="width:160px;">Страна постройки<\/td><td>[^<]*<\/td><\/tr>', driver.page_source)[0]
		country = re.findall(r'(?<=<td>)[^<]*(?=<\/td>)', country)[0]
	except:
		country = ''
	vessel['country'] = country

	try:
		driver.get(owner_link)
		owner_page = driver.page_source
		try:
			owner_country = re.findall(r'(?<=<b>Страна:<\/b>)[^<]*(?=<b>)', owner_page)[0]
			owner_country = owner_country.strip(",")
		except:
			owner_country = ''

		try:
			owner_city = re.findall(r'(?<=<b>Город:<\/b>)[^<]*(?=<\/span>)', owner_page)[0]
		except:
			owner_city = ''

		try:
			address = re.findall(r'(?<=<b>Адрес:<\/b>)[^<]*(?=<\/span>)', owner_page)[0]
		except:
			address = ''

		try:
			phone = re.findall(r'(?<=<b>Телефон:<\/b>)[^<]*(?=<b>)', owner_page)[0]
			phone = phone.strip(',')
		except:
			phone = ''

		try:
			email = re.findall(r'(?<=<b>E-mail<\/b>: <a target="_blank" href="mailto:)[^"]*(?=">)', owner_page)[0]
		except:
			email = ''

		try:
			owner_site = re.findall(r'(?<=<b>URL<\/b>: <a href=")[^"]*(?=" target)', owner_page)[0]
		except:
			owner_site = ''


	except:
		print("Owner info error")

	vessel["owner_country"] = owner_country
	vessel["owner_city"] = owner_city
	vessel['address'] = address
	vessel['phone'] = phone
	vessel['email'] = email
	vessel['owner_site'] = owner_site

	print(imo)
	print(owner_country, owner_city, address, phone, email, owner_site)

	with open('report.csv', 'a') as f:
		w = csv.DictWriter(f, vessel.keys(), delimiter = ';')
		w.writerow(vessel)
		f.close()
	time.sleep(1)
	#print(name, imo, type_, mmsi, dwt, gt, me, year, number, state, flag, port, owner, country)

	
	'''
	with open("resp.txt", "w") as f:
		f.write(str(resp))
		f.close()

	searchResult = driver.find_element_by_id("searchResult")
	print(searchResult.get_attribute("innerHTML	"))
	break
	soup = BeautifulSoup(searchResult.get_attribute("innerHTML"), "lxml")
	links = soup.find_all("a")
	for a in links:
		print(a['href'])

	'''
with open("error_imos.csv", "w") as csvfile:
	writer = csv.writer(csvfile, delimiter = ";")
	writer.writerow(error_imos)
	csvfile.close()

#driver.close()

