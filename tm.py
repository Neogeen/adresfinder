import pandas as pd
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException, TimeoutException, UnexpectedAlertPresentException
from selenium.webdriver.firefox.firefox_binary import FirefoxBinary
from selenium.webdriver.firefox.options import Options as FirefoxOptions
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from time import sleep

driver = webdriver.Firefox()
path = 'канцелярия докторантура2 (2019-2020) на работу.xlsx'
table = pd.read_excel(path, sheet_name='work')

Op = table['adres'].to_list()
try:
    for i in range(1000):
        driver.get('https://post.kz/postcode')
        sleep(1)
        driver.find_element_by_xpath('/html/body/div/div[3]/div/main/div/div/form/input').send_keys(Op[i])
        sleep(7)
        ind = driver.find_element_by_xpath('/html/body/div/div[3]/div/main/div/div/form/ul/li[1]').text
        index = ind.split(' ')[0]
        print(index)
        driver.get('https://post.kz/postcode')
        k = Op
except:
    print(k)
table.to_excel('out.xlsx', index=False)