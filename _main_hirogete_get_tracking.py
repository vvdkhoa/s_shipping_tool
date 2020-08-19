from selenium import webdriver # https://mylife8.net/install-selenium-and-run-on-windows/
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import re
from time import sleep
from set_spreadsheet import set_spreadsheet_ship_table, sheet_write_ship_table
from clean_data import to_float
# import lxml.html as lh
# import pandas as pd
# import json

###########################################################################
# https://mylife8.net/install-selenium-and-run-on-windows/
# https://stackoverflow.com/questions/49290704/python-save-html-from-browser
# Return all page from HIROGETE 発送待ち
def get_html_chrome(url):

	ListingId_list = []
	ItemLink_list = []

	driver = webdriver.Chrome()
	driver.get(url)

	# Wait Login form
	wait_xpath(driver= driver, xpath='/html/body/div/div/div/form/div/div[2]/input')

	# Send ID and pass and click login
	id = driver.find_element(By.XPATH, '/html/body/div/div/div/form/div/div[2]/input')
	id.send_keys("b.yodo.express@gmail.com")
	password = driver.find_element(By.XPATH, '/html/body/div/div/div/form/div/div[3]/input')
	password.send_keys("mmMm13547")
	login = driver.find_element(By.XPATH, '/html/body/div/div/div/form/div/div[5]/div/button')
	login.click()

	# Wait loading homepage
	wait_xpath(driver= driver, xpath='/html/body/ui-view/div[3]')

	# Close all note box:
	note_box_xpath = [	'//*[@id="modal-notice-content-53"]/div[1]/button', #　日本郵便...
					'//*[@id="modal-notice-content-36"]/div[1]/button',	# 同期できない...
					'//*[@id="modal-notice-content-54"]/div[1]/button',
					'//*[@id="modal-notice-content-71"]/div[1]/button'
				]
	for xpath in note_box_xpath:
		try:
			driver.find_element(By.XPATH, xpath).click()
		except:
			print ('    Closed all Hirogete note box')

	# Directions to 発送待ち
	navi_xpath = [	'//*[@id="navbarHeight"]/div[1]/ul[1]/li/a/i',					# Show list icon
					'/html/body/ui-view/div[2]/div/div[1]/div/div/div[1]/ul/li[2]',	# 注文管理
					'//*[@id="myTabs"]/li[2]/a'
				]
	for xpath in navi_xpath:
		try:
			driver.find_element(By.XPATH, xpath).click()
			wait_xpath(driver, xpath)
		except:
			print ('    Error encountered when directions to "発送待ち". Please try again')

	# Wait loading all tracking
	wait_xpath(driver= driver, xpath='//*[@id="jp_shipment_grid_0_wrapper"]/div[2]')

	all_html = []
	i = 1
	while True:

		sleep(5)
		all_html.append(driver.page_source)

		# Next page from page 2
		i += 1
		new_page_xpath = '//*[@id="jp_shipment_grid_0_paginate"]/span/a[' + str(i) + ']'
		try:
			driver.find_element_by_xpath(new_page_xpath).click()
			sleep(5)
		except:
			break

	driver.quit()

	return all_html


###########################################################################
# https://qiita.com/ColdFreak/items/f8f6a9e3c957862a53ba
def wait_xpath(driver, xpath):
	# driver = webdriver.Chrome()
	# driver.get(url)
	try:
		element = WebDriverWait(driver, 60).until(
			EC.presence_of_element_located((By.XPATH, xpath)))
		# print ('    Page loaded')
	except:
		print ('    Taking too much time in loading')

	sleep(2)


###########################################################################
# Scrapt Record, Tracking number... from str html
def scrapt(html):

	res = []

	all_row = re.findall(r'<input type="checkbox" name="select-row">.*?</tr>', html)
	for row in all_row:

		record = re.search(r'>  .*?</td>', row).group(0)
		record = record.replace('>  ', '').replace('</td>', '')
		record = int(record)

		price = re.search(r'>\$.*?</td>', row).group(0)
		price = price.replace('>$', '').replace('</td>', '')
		price = to_float(price)

		tracking = re.search(r'reqCodeNo1=.*?&amp', row).group(0)
		tracking = tracking.replace('reqCodeNo1=', '').replace('&amp', '')

		img = re.search(r'<img src=".*?"', row).group(0)
		img = img.replace('<img src="', '').replace('"', '')

		# Create a row like shipping tool
		res.append([record,'','','','','',price,'','','','','','','','','','',tracking, img])

	return res


###########################################################################
def write_tracking(ship_data):

	ShippingTool = set_spreadsheet_ship_table('ShippingTool')
	Old_record = ShippingTool.col_values(2) # Manual setting

	# Check and remove duplicate with old record:
	new_data = []
	for i in ship_data:
		if str(i[0]) not in Old_record:
			new_data.append(i)

	# Add all and write to sheet
	if new_data != []:
		w_data = []
		for i in new_data:
			w_data += i
		sheet_write_ship_table('ShippingTool', len(Old_record)+1, 2, len(Old_record)+ len(new_data), 20, w_data) # Manual setting
	else:
		print('    There is no new data')


###########################################################################
def main_scrap_store():

	url = 'https://app.hirogete.com/#!/order/shipment/list_jp'
	all_html = get_html_chrome(url)

	ship_data = [] # Write to shipping tool
	for html in all_html:
		ship_data += scrapt(html)

	write_tracking(ship_data)

###########################################################################
if __name__ == '__main__':
	main_scrap_store()
	