"""
@author: Joey Leurs
Version 2.1
Date: 17/07/2022

Webscraping tool to extract price data for the Elder Scrolls Online.
"""

from selenium import webdriver
from selenium.webdriver.common.by import By
import time
from playsound import playsound
import re
import pandas as pd

df = pd.read_excel('esoprices.xlsx')


discount = 5#percent of average
lastseen = 15#minutes or less

#webdriver
driver = webdriver.Chrome('E:\Downloads\chromedriver_win32\chromedriver.exe')

#Scraping   
counter = 1


for i, x in enumerate(df['link']):
    driver.get(x)
    time.sleep(1)
    counter = 1
    deals_found = 0
    while counter < 19:
        for post in range(1, 20):
                tr_count = str(counter)
                
                item = driver.find_element(By.XPATH, '//*[@id="search-result-view"]/div[1]/div/table/tbody/tr['+tr_count+']/td[1]/div[1]').text
                location = driver.find_element(By.XPATH, '//*[@id="search-result-view"]/div[1]/div/table/tbody/tr['+tr_count+']/td[3]').text
                price = driver.find_element(By.XPATH, '//*[@id="search-result-view"]/div[1]/div/table/tbody/tr['+tr_count+']/td[4]/span[1]').text
                seen = driver.find_element(By.XPATH, '//*[@id="search-result-view"]/div[1]/div/table/tbody/tr['+tr_count+']/td[5]').text
                
                price_n = float(price.replace(',', ''))
                price_a = float(df['price'][i])
                price_d = price_a - (price_a * discount / 100)
                             
                if seen.find("Hour") != -1:
                    seen = '600'
                if seen.find("Now") != -1:
                    seen = '1'
                seentime = int(re.sub('[^0-9]', '', seen))
                
                if counter < 19:
                    counter += 2
                else:
                    counter = 1
                
                if price_n < price_d and seentime < lastseen:
                    print(item, location, price, seen+' minutes ago', sep='\n', end='\n\n')
                    deals_found = 1
                    playsound('C:/Windows/Media/tada.wav')
        if  deals_found == 0:
            print('no deals found for '+item, end='\n\n')
        else:
            deals_found = 0
#eof
print("Finished scanning for deals")