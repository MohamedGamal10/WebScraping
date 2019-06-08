# -*- coding: utf-8 -*-
"""
Created on Wed Dec 26 22:17:52 2018

@author: Mohamed
"""
from selenium import webdriver
from selenium.webdriver.common import keys
import pandas as pd
import os
import time
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
# Web Driver Config
driver = webdriver.Firefox(executable_path=r'C:\Users\mohamed.elzayat\.anaconda\geckodriver-v0.23.0-win64\geckodriver.exe')
#Site Crawling Path
driver.get('https://www.microchip.com/exportcontroldata/')
#Search Bar by xpath
search_bar = driver.find_element_by_xpath('//*[@id="radSearch_Input"]')
#Excel file of parts
df=pd.read_excel('C:\\Users\\mohamed.elzayat\\Desktop\\microchip2\\Products.xlsx')
#counter of names of images
num=0
#list of data scrap from site
total_data=list()
# main code 
for i in range (0,len(df)):
    #counter on parts in Excel input
    x=df.iloc[i].to_string(index=False)
    #send parts to search bar
    search_bar.send_keys(x)
    #identify saerch buton by link text and click on search button
    search_button=driver.find_element_by_link_text('Search')
    search_button.click()
    # wait until data loded and scrap data
    #XPath_with_data= WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[1]/form/section/div[1]/div/div[2]/section/div/div/div/div[2]')))
    #XPath_without_data= WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[1]/form/section/div[1]/div/div[2]/section/div/div/div[2]')))
    def not_found ():
        try:
            WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[1]/form/section/div[1]/div/div[2]/section/div/div/div/div[2]')))
        except Exception :
            return False
    if not_found() == False :
        data = WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[1]/form/section/div[1]/div/div[2]/section/div/div/div[2]')))
    else :
        data = WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[1]/form/section/div[1]/div/div[2]/section/div/div/div/div[2]')))
    #add every data scrap to the list 
    total_data.append(data.text)
    #screenshot of data , and change the name of image because it save on the same name
    driver.save_screenshot('C:\\Users\\mohamed.elzayat\\Desktop\\\microchip2\\m.png')
    os.chdir('C:\\Users\\mohamed.elzayat\\Desktop\\microchip2')
    b=str(num)+'.png'
    os.rename('m.png',b)
    num=num+1
    # wait until data loded
    data
    #clear search bar from last part to add new part on it
    search_bar.clear()
    data
    time.sleep(3)
#convert all data in list to data series and save it in excel 
final_total_data=pd.Series(total_data)
df['data']=final_total_data
df.to_excel("data.xlsx",sheet_name="new")