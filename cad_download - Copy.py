from selenium.webdriver.chrome.options import Options
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from selenium.webdriver.support.select import Select
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.remote.webelement import WebElement
from selenium.common.exceptions import NoSuchElementException
import time
import csv
import json
import os
from pprint import pprint

driver = webdriver.Chrome(r"chromedriver_win32\chromedriver.exe")
driver.get("https://www.ruland.com/amtcl-10-2-f.html")
driver.maximize_window()
driver.implicitly_wait(7)

login = driver.find_element(
    By.XPATH, "//body/div[6]/main/div[2]/div/div/div[7]/ul/li/div[3]")
login.click()

email = driver.find_element(
    By.XPATH, "//body/div[6]/main/div[2]/div/div/div[8]/div[2]/form/div[1]/div/input")
email.send_keys("adharva@gmail.com")
password = driver.find_element(
    By.XPATH, "//body/div[6]/main/div[2]/div/div/div[8]/div[2]/form/div[2]/div/input")
password.send_keys("adharva@123")
login_button = driver.find_element(
    By.XPATH, "//body/div[6]/main/div[2]/div/div/div[8]/div[2]/form/div[3]/div/button")
login_button.click()
time.sleep(5)

url_data = []
with open('input.csv') as f:
    csvr = csv.DictReader(f)
    for line in csvr:
        url_data.append(line)
for ind, val in enumerate(url_data):
    driver.get(val['url'])
    while (driver.execute_script("return document.readyState") != 'complete'):
        continue
    driver.implicitly_wait(8)
    if len(driver.find_elements(By.XPATH, "//body/div[6]/main/div[2]/div/div/div[7]/ul/li/div[2]/label/select")) > 0:
        try:
            cads = driver.find_element(
                By.XPATH, "//body/div[6]/main/div[2]/div/div/div[7]/ul/li/div[2]/label/select") 
            select_cad = Select(cads)
            select_cad.select_by_index(19)
            driver.implicitly_wait(5)
            driver.execute_script("window.scrollTo(0,window.scrollY + 400)")
            download = driver.find_element(
                By.XPATH, "//body/div[6]/main/div[2]/div/div/div[7]/ul/li/div[3]/a")
            download.click()
            time.sleep(2)
        except NoSuchElementException:
            continue
    else:
        continue
driver.close()
