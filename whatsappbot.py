from selenium import webdriver 
from selenium.webdriver.support.ui import WebDriverWait 
from selenium.webdriver.support import expected_conditions as EC 
from selenium.webdriver.common.keys import Keys 
from selenium.webdriver.common.by import By 
import time
import openpyxl as excel

def readContacts(fileName):
    lst = []
    file = excel.load_workbook(fileName)
    sheet = file.active
    firstCol = sheet['A']
    for cell in range(len(firstCol)):
        contact = str(firstCol[cell].value)
        contact = "\"" + contact + "\""
        lst.append(contact)
    return lst

def setup():
	# driver = webdriver.Firefox()
	profileManager = webdriver.FirefoxProfile('C:/Users/Fatwa/AppData/Roaming/Mozilla/Firefox/Profiles/x6h2u93g.default')
	driver = webdriver.Firefox(profileManager)
	return driver

def send_message_with_contact(target, string, driver):
	driver.get('https://web.whatsapp.com/')
	wait = WebDriverWait(driver, 60)
	x_arg = '//span[contains(@title, \'' + target + '\')]'
	group_title = wait.until(EC.presence_of_element_located((By.XPATH, x_arg)))
	group_title.click()
	inp_xpath = '//div[@dir="ltr"][@data-tab="1"][@contenteditable="true"][@spellcheck="true"]'
	input_box = wait.until(EC.presence_of_element_located((By.XPATH, inp_xpath)))
	for i in range(1):
		input_box.send_keys(string + Keys.ENTER)
		time.sleep(1)
	driver.close()

def send_message_without_contact(target, string, driver):
	driver.get('https://web.whatsapp.com/send?phone=' + target)
	wait = WebDriverWait(driver, 60)
	inp_xpath = '//div[@dir="ltr"][@data-tab="1"][@contenteditable="true"][@spellcheck="true"]'
	input_box = wait.until(EC.presence_of_element_located((By.XPATH, inp_xpath)))
	for i in range(1):
		input_box.send_keys(string + Keys.ENTER)
		time.sleep(1)
	driver.close()

def send_message_with_excel(target, string, wait):
	x_arg = '//span[contains(@title, \'' + target + '\')]'
	group_title = wait.until(EC.presence_of_element_located((By.XPATH, x_arg)))
	group_title.click()
	inp_xpath = '//div[@dir="ltr"][@data-tab="1"][@contenteditable="true"][@spellcheck="true"]'
	input_box = wait.until(EC.presence_of_element_located((By.XPATH, inp_xpath)))
	for i in range(60):
		input_box.send_keys(string + Keys.ENTER)
		time.sleep(60)

pilihan = input('Kontak, Nomor Telepon atau Excel? (K/N/E)?')
if pilihan == 'K' or pilihan == 'k':
	target = input('Nama Kontak:')
	string = input('Pesan:')
	driver = setup()
	send_message_with_contact(target, string, driver)
elif pilihan == 'N' or pilihan == 'n':
	target = input('Nomor Telepon:')
	string = input('Pesan:')
	driver = setup()
	send_message_without_contact(target, string, driver)
elif pilihan == 'E' or pilihan == 'e':
	targets = readContacts("contacts.xlsx")
	string = input('Pesan:')
	driver = setup()
	driver.get('https://web.whatsapp.com/')
	wait = WebDriverWait(driver, 60)
	for target in targets:
		send_message_with_excel(target, string, wait)
	driver.close()