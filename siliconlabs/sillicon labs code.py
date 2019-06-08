# -*- coding: utf-8 -*-
"""
Created on Thu Jan 24 06:51:34 2019

@author: Mohamed
"""
#Required Lib
from selenium import webdriver
from selenium.webdriver.common import keys
import pandas as pd
import time
from selenium.webdriver.firefox.firefox_profile import FirefoxProfile
from selenium.webdriver.firefox.firefox_binary import FirefoxBinary
binary = FirefoxBinary("C:\\Program Files\\Mozilla Firefox\\firefox.exe")
profile = FirefoxProfile("C:\\Users\\Mohamed\\AppData\\Roaming\\Mozilla\\Firefox\\Profiles\\8hw2l42j.default")
driver = webdriver.Firefox(firefox_profile=profile, firefox_binary=binary, executable_path=r'C:\Users\Mohamed\Anaconda3\geckodriver-v0.23.0-win64\geckodriver.exe')

#parts that we need its data
df=pd.read_excel('D:\\Silicon labs\\Products.xlsx')

#open Browser
def browser() :
    driver.get('https://ops.silabs.com/rfi/Pages/PartSearch')
    password = driver.find_element_by_xpath('//*[@id="passwordBox"]')
    password.send_keys('3237258m')
    login = driver.find_element_by_xpath('//*[@id="mktoLogin"]')
    login.click()
    time.sleep(60)

#Select part and put it in the Search box then click search
def select_product(i):
    search_box=driver.find_element_by_xpath('//*[@id="MainContent_PartSearchUserControl1_txtPartNumber"]')
    x=df.iloc[i].to_string(index=False)
    search_box.send_keys(x)
    search_btn=driver.find_element_by_xpath('//*[@id="MainContent_PartSearchUserControl1_ImageButtonInput"]')
    search_btn.click()
    time.sleep(2)

#get the range of number parts that appear in search 
def get_range () :
    try :
        range_of_output = driver.find_element_by_xpath('//*[@id="MainContent_PartSearchListUserControl1_gvProjects"]')
        value = range_of_output.get_attribute('cellspacing')
        return value
    except Exception :
        value = 0
        return value

#click on the selected part
def get_data(value):
    counter=int(value)
    for i in range(1,counter):
        y='#MainContent_PartSearchListUserControl1_gvProjects > tbody:nth-child(1) > tr:nth-child({})'.format(i)
        result=driver.find_element_by_css_selector(y)
        result.click()
        show_all=driver.find_element_by_xpath('//*[@id="MainContent_PartSearchListUserControl1_btnShowAllData"]')
        show_all.click()
        time.sleep(10)
        data=driver.find_element_by_xpath('//*[@id="MainContent_ContainerSearchDataResultUserDataControl1_ImageBtnExportToExcel"]')
        data.click()
        driver.back() 
    try :
        driver.back()
    except Exception :
        time.sleep(2)
        driver.back()
list_multiple= [50,100,200,250,300,350,400,450,500,550,600,650,700,750,800,850,900,950]

#Begin of the code
browser()
for i in range (0,len(df)):
    try :
        select_product(i)
    except Exception :
        driver.get('https://ops.silabs.com/rfi/Pages/PartSearch')
        select_product(i)
    except Exception:
        browser()
        select_product(i)
    get_range()
    if int(get_range()) > 0 :
        try :
            get_data(int(get_range()))
        except Exception :
            driver.back()
            time.sleep(2)
            driver.back()
            select_product(i)
            get_range()
            get_data(int(get_range()))
    else:
        driver.back()
        time.sleep(5)
    if i in list_multiple :
        driver.execute_script("window.open('');")
        time.sleep(2)
        driver.switch_to.window(driver.window_handles[0])
        driver.close()
        time.sleep(1)
        driver.switch_to.window(driver.window_handles[0])
        try :
            browser()
        except Exception :
            driver.get('https://ops.silabs.com/rfi/Pages/PartSearch')