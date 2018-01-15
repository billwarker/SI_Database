import selenium
from selenium import webdriver
import os
from settings import *
from bs4 import BeautifulSoup
import openpyxl
import datetime


def convert_xml_to_xlsx(filename, name, inventory_dir):
	file = open(filename, 'r')
	soup = BeautifulSoup(file, 'lxml')
	wb = openpyxl.Workbook()
	sheet = wb.active
	table = []
	for row in soup.find_all("row"):
		new_row = []
		for cell in row.find_all("cell"):
			new_row.append(str(cell.get_text()))
		table.append(new_row)

	for r, row in enumerate(table):
		for c, col in enumerate(row):
			cell = sheet.cell(row=r+1, column=c+1)
			cell.value = col
	output = name + ".xlsx"
	wb.save(os.path.join(inventory_dir, output))

	return os.path.join(inventory_dir, output)

def download_inventory(name, payload, download_dir, driver_dir, inventory_dir):
	name = "{} Inventory {}".format(name, datetime.date.today().strftime("%m-%d-%Y"))
	portal_url = "http://4pl.leansupplysolutions.com/portal/login.php"
	inventory_url = "http://4pl.leansupplysolutions.com/portal/inventory.php"
	download_before = os.listdir(download_dir)

	driver = webdriver.Chrome(driver_dir)
	driver.get(portal_url)

	user = driver.find_element_by_name("username")
	user.send_keys(payload["username"])

	pw = driver.find_element_by_name("userpassword")
	pw.send_keys(payload["userpassword"])

	driver.find_element_by_id("cmdsave").click()
	driver.find_element_by_xpath('//a[@href="'+ "inventory.php" +'"]').click()
	driver.find_element_by_id("downloadbtn").click()
	print("Downloading...")
	downloaded_file = ''
	while downloaded_file.endswith('.xls') == False:
		download_after = os.listdir(download_dir)

		for file in download_after:
			if file not in download_before:
				downloaded_file = file

	driver.quit()
	filename = os.path.join(download_dir, downloaded_file)
	print(filename)

	output_path = convert_xml_to_xlsx(filename, name, inventory_dir)
	os.remove(os.path.join(download_dir, filename))
	print(name, "downloaded.")
	
	return output_path

if __name__ == '__main__':
	download_inventory("STAR", star_payload, download_dir, driver_dir, inventory_dir)
	download_inventory("SBW", sbw_payload, download_dir, driver_dir, inventory_dir)

