from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
import time
from selenium import webdriver
import requests
import os
import pandas
from bs4 import BeautifulSoup
import re
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.common.exceptions import TimeoutException
from datetime import datetime
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import StaleElementReferenceException
from selenium.common.exceptions import ElementNotInteractableException
from selenium.common.exceptions import ElementClickInterceptedException
from selenium.common.exceptions import WebDriverException
import pandas as pd
import csv
from dotenv import load_dotenv

Options = webdriver.ChromeOptions()
Options.add_argument('--no-sandbox')
Options.add_argument('--disable-dev-shm-usage')
Options.add_argument('--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3')
Options.add_argument('--start-maximized')

driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=Options)

df = pd.read_excel('USR.xlsx')
SKUs = df.iloc[9:, 0].tolist()  
Brands = df.iloc[9:, 1].tolist()  

# Function to scrape data and update Excel
def search_extract_data():
    driver.get('https://www.kitchenall.com/')
    time.sleep(2)
    
    for sku, brand in zip(SKUs, Brands):
        try:
            search_field = driver.find_element(By.XPATH, "//input[@aria-autocomplete='both']")
            search_field.clear()
            search_field.send_keys(sku, ' ', brand)
            search_field.send_keys(Keys.RETURN)
            time.sleep(3)
            

            item = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//a[@class='result']")))
            item.click()
            time.sleep(4)
            
            try:

                main_cat = driver.find_element(By.XPATH, "//div[@class='breadcrumbs']//li[2]").text
            except:
                main_cat = ''
            try:   
                sub_cat = driver.find_element(By.XPATH, "//div[@class='breadcrumbs']//li[3]").text
            except:
                sub_cat = ''
            try:
                sub_cat_2 = driver.find_element(By.XPATH, "//div[@class='breadcrumbs']//li[4]").text
            except:
                sub_cat_2 = ''
            price = driver.find_element(By.XPATH, "//div[@class='product-info-main mobi pdp-block']//span[@class='price']").text
            shipping = driver.find_element(By.XPATH, "//div[@class='product-info-main mobi pdp-block']//span[@class='shipping_price']").text.replace('Shipping', '').strip()
        
            try:
                spec_sheet = driver.find_element(By.XPATH, "//a[contains(.,'Spec sheet')]").get_attribute('href')
            except NoSuchElementException:
                spec_sheet = ''
                
            try:
                manual = driver.find_element(By.XPATH, "//a[contains(text(),'Manual')]").get_attribute('href')
            except NoSuchElementException:
                manual = ''
                
            try:
                warranty = driver.find_element(By.XPATH, "(//div[@itemprop='product_warranty_select'])[2]").text
            except NoSuchElementException:
                warranty = ''
            try:    
                approval = driver.find_element(By.XPATH, "(//td[@data-th='Approval'])[2]").text
            except:
                approval = ''
            # Update Excel with scraped data
            row_index = df[(df['SKU'] == sku) & (df['Brand'] == brand)].index[0]
            df.at[row_index, 'Price'] = price
            df.at[row_index, 'Shipping'] = shipping
            df.at[row_index, 'Specsheet'] = spec_sheet
            df.at[row_index, 'Approvals'] = approval
            df.at[row_index, 'Manual'] = manual
            df.at[row_index, 'Warranty'] = warranty
            df.at[row_index, 'Main'] = main_cat
            df.at[row_index, 'Sub Cat 1'] = sub_cat
            df.at[row_index, 'Sub Cat 2'] = sub_cat_2

            df.to_excel('USR.xlsx', index=False)
            
        except (NoSuchElementException, TimeoutException, StaleElementReferenceException, ElementNotInteractableException, ElementClickInterceptedException) as e:
            print(f"Error scraping data for SKU {sku} and Brand {brand} - {e}")
            pass


search_extract_data()


