#!/usr/bin/env python
# coding: utf-8

#Install missing packages
import subprocess
import sys

def install_if_missing(package):
    try:
        __import__(package)
    except ImportError:
        subprocess.check_call([sys.executable, "-m", "pip", "install", package])

install_if_missing('selenium')
install_if_missing('pandas')
install_if_missing('openpyxl')




#Import
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from datetime import datetime
import os
import pandas as pd
import time
import tkinter as tk
from tkinter import ttk
from datetime import datetime
import threading
import time
import calendar
from pathlib import Path



#Services and Options Safe
service = Service(log_path=os.devnull)
options = webdriver.ChromeOptions()
options.add_argument("--headless=new")  # Use new headless mode (more stable)
options.add_argument("--disable-gpu")
options.add_argument("--window-size=1920,1080")
options.add_argument("--disable-extensions")
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")
options.add_argument("--disable-blink-features=AutomationControlled")
options.add_experimental_option("excludeSwitches", ["enable-automation"])
options.add_experimental_option("useAutomationExtension", False)

#Services and Options Un-Safe
"""
options = webdriver.ChromeOptions()
options.add_argument("--headless=new")           # Fastest headless mode
options.add_argument("--disable-gpu")
options.add_argument("--window-size=1920,1080")
options.add_argument("--disable-extensions")
options.add_argument("--blink-settings=imagesEnabled=false")  # No image load
options.page_load_strategy = 'eager'  # Don't wait for full load (faster)
"""

prefs = {"profile.managed_default_content_settings.images": 2}
options.add_experimental_option("prefs", prefs)

#Folder path
timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
file_name = f"{timestamp}_CompanyData.xlsx"

script_dir = os.path.dirname(os.path.abspath(__file__))
reports_folder = os.path.join(script_dir, "reports")
os.makedirs(reports_folder, exist_ok=True)
file_path = os.path.join(reports_folder, file_name)

#Open web
driver = webdriver.Chrome(service=service, options=options)
#driver = webdriver.Chrome()
driver.get("https://www.set.or.th/th/market/get-quote/stock/")
driver.implicitly_wait(15)

#Declaration and values
Keywords = ['ไม่','ดี']
KeywordsEN = ['yes','no']
DataDict = []
ListofLinks = []
CompanyList = []

maxpageamt = 0
maxcompanyamt = 0
newscardamt = 0
CurrentLanguague = 'TH'

#Progression
pageprg = 1
companyprg = 1
newsprg = 1

#Custom
#newscardamtoveride = 0
stopcommand = False
dateselectionmode = True
Enddateselectionmode = True
#Start Date
keyedDate = 1
keyedMonth = 1
keyedYear = 2567

#End Date
keyedEndDate = 1
keyedEndMonth = 1
keyedEndYear = 2567

#Special Command


# ========== General Function ==========

def log_to_gui(log_widget, message):
    log_widget.insert(tk.END, message + "\n")
    log_widget.see(tk.END)

def setnewscardamt():
    global newscardamt
    newscardamt = len(ListofLinks)

def setdate(day,month,year):
    global keyedDate
    global keyedMonth
    global keyedYear

    if day == '' or month == '' or year == '':
        keyedDate = datetime.now().strftime("%d")
        keyedMonth = datetime.now().strftime("%m")
        keyedYear = datetime.now().strftime("%Y")
        keyedYear = (int(keyedYear) + 543) - 5
        return


    keyedDate = day
    keyedMonth = month
    keyedYear = int(year) + 543

def setEnddate(day,month,year):
    global keyedEndDate
    global keyedEndMonth
    global keyedEndYear

    if day == '' or month == '' or year == '':
        keyedEndDate = datetime.now().strftime("%d")
        keyedEndMonth = datetime.now().strftime("%m")
        keyedEndYear = datetime.now().strftime("%Y")
        keyedEndYear = (int(keyedEndYear) + 543)
        return


    keyedEndDate = day
    keyedEndMonth = month
    keyedEndYear = int(year) + 543


#Get Max Page Amount
def GetMaxPageAmt():
    maxpage = driver.find_element(By.XPATH,'/html/body/div[1]/div/div/div[2]/div[2]/div[2]/div/div/div[3]/div[2]/div[2]/div[3]/ul/li[8]/button')
    global maxpageamt
    maxpageamt = maxpage.text
    driver.implicitly_wait(15)

#Get Max Company in 1 page
def GetMaxCompanyAmt():
    tbody = driver.find_element(By.XPATH, '/html/body/div[1]/div/div/div[2]/div[2]/div[2]/div/div/div[3]/div[2]/div[1]/div[2]/table/tbody')
    rows = tbody.find_elements(By.TAG_NAME, 'tr')
    global maxcompanyamt
    maxcompanyamt = len(rows)
    print('Found', len(rows), 'rows.')
    driver.implicitly_wait(15)

#Get Company Name
def EnterCompanyPage(PageNumber):
    if CurrentLanguague == 'TH':
        driver.get("https://www.set.or.th/th/market/product/stock/quote/"+ CompanyList[PageNumber - 1] +"/news")
    else:
        driver.get("https://www.set.or.th/en/market/product/stock/quote/"+ CompanyList[PageNumber - 1] +"/news")
    
    driver.implicitly_wait(15)

#Next Page
def SelectSubCompanyListPage():
    if pageprg == 1:
        return()

    if pageprg == 2:
        pageloc = 3
    elif pageprg == 3:
        pageloc = 4
    elif pageprg == 4:
        pageloc = 5
    else:
        pageloc = 6


    if pageprg > 5:
        for i in range(4):
            changepage = driver.find_element(By.XPATH, '/html/body/div[1]/div/div/div[2]/div[2]/div[2]/div/div/div[3]/div[2]/div[2]/div[3]/ul/li[' + str(i+2) + ']/button')
            driver.execute_script("arguments[0].click();", changepage)

        changepage = driver.find_element(By.XPATH, '/html/body/div[1]/div/div/div[2]/div[2]/div[2]/div/div/div[3]/div[2]/div[2]/div[3]/ul/li[' + str(6) + ']/button')
        driver.execute_script("arguments[0].click();", changepage)

    if pageprg > 6:
        amount = int(pageprg) - 5

        if (int(maxpageamt) - int(pageprg)) == 1:
            amount = amount - 1

        if (int(maxpageamt) - int(pageprg)) == 0:
            amount = amount - 2


        for i in range(int(amount)):
            changepage = driver.find_element(By.XPATH, '/html/body/div[1]/div/div/div[2]/div[2]/div[2]/div/div/div[3]/div[2]/div[2]/div[3]/ul/li[' + str(pageloc) + ']/button')
            driver.execute_script("arguments[0].click();", changepage)

        if (int(maxpageamt) - int(pageprg)) == 1:
            pageloc = 7
            changepage = driver.find_element(By.XPATH, '/html/body/div[1]/div/div/div[2]/div[2]/div[2]/div/div/div[3]/div[2]/div[2]/div[3]/ul/li[' + str(pageloc) + ']/button')
            driver.execute_script("arguments[0].click();", changepage)

        if (int(maxpageamt) - int(pageprg)) == 0:
            pageloc = 7
            changepage = driver.find_element(By.XPATH, '/html/body/div[1]/div/div/div[2]/div[2]/div[2]/div/div/div[3]/div[2]/div[2]/div[3]/ul/li[' + str(pageloc) + ']/button')
            driver.execute_script("arguments[0].click();", changepage)

            pageloc = 8
            changepage = driver.find_element(By.XPATH, '/html/body/div[1]/div/div/div[2]/div[2]/div[2]/div/div/div[3]/div[2]/div[2]/div[3]/ul/li[' + str(pageloc) + ']/button')
            driver.execute_script("arguments[0].click();", changepage)

    else:
        changepage = driver.find_element(By.XPATH, '/html/body/div[1]/div/div/div[2]/div[2]/div[2]/div/div/div[3]/div[2]/div[2]/div[3]/ul/li[' + str(pageloc) + ']/button')
        driver.execute_script("arguments[0].click();", changepage)

#Navigate to News Page
def GetPageCompany():
    global CompanyList
    for i in range(maxcompanyamt):
        comname = driver.find_element(By.XPATH,'/html/body/div[1]/div/div/div[2]/div[2]/div[2]/div/div/div[3]/div[2]/div[1]/div[2]/table/tbody/tr['+ str(i+1) +']/td[1]/div/span/div/div/a/div')
        CompanyList.append(comname.text)
        driver.implicitly_wait(15)
        time.sleep(0.1)

"""
#Select End Date 
def Enddateselection(Year,Month,Date):
    if Enddateselectionmode == False:
        return
    
    
    driver.implicitly_wait(15)
    if CurrentLanguague == 'TH':
        calendarbtn = driver.find_element(By.XPATH, '/html/body/div[1]/div/div/div[2]/div[2]/div[2]/div[2]/div/div[1]/div/div[3]/div/div/div/div/span[2]/input')
        driver.execute_script("arguments[0].click();", calendarbtn)
        MonthList = ['มกราคม','กุมภาพันธ์','มีนาคม','เมษายน','พฤษภาคม','มิถุนายน','กรกฎาคม','สิงหาคม','กันยายน','ตุลาคม','พฤศจิกายน','ธันวาคม']
        Month = MonthList[int(int(Month)-1)]
    else:
        calendarbtn = driver.find_element(By.XPATH, '/html/body/div[1]/div/div/div[2]/div[2]/div[2]/div/div/div[1]/div/div[3]/div/div/div/div/span[2]/input')
        driver.execute_script("arguments[0].click();", calendarbtn)
        MonthList = ['January','February','March','April','May','June','July','August','September','October','November','December']
        Month = MonthList[int(int(Month)-1)]
        Year = int(Year) - 543
    
    log_to_gui(app.log_box, "Searching through End Date")
    Endmonthandyearcycle(Year,Month,Date)

def Endmonthandyearcycle(Year,Month,Date):
    driver.implicitly_wait(15)
    Target = str(Month) + " " + str(Year)
    if CurrentLanguague == 'TH':
        monthanyear = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div/div/div[2]/div[2]/div[2]/div[2]/div/div[1]/div/div[3]/div/div/div/div/span[2]/div/div/div/div[2]/div[1]/div/div[1]')))
        valuenow = monthanyear.text
        time.sleep(0.2)
        if (valuenow == Target) == False:
            nextmonthbtn = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[1]/div/div/div[2]/div[2]/div[2]/div[2]/div/div[1]/div/div[3]/div/div/div/div/span[2]/div/div/div/div[2]/div[2]/div[1]')))
            driver.execute_script("arguments[0].click();", nextmonthbtn)
            time.sleep(0.2)
            Endmonthandyearcycle(Year,Month,Date)
            return
        else:
            time.sleep(0.3)
            Enddateselectioncycle(Date,8)
            return
    else:
        monthanyear = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div/div/div[2]/div[2]/div[2]/div/div/div[1]/div/div[3]/div/div/div/div/span[2]/div/div/div/div[2]/div[1]/div/div[1]/div')))
        valuenow = monthanyear.text
        time.sleep(0.2)
        if (valuenow == Target) == False:
            nextmonthbtn = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[1]/div/div/div[2]/div[2]/div[2]/div/div/div[1]/div/div[3]/div/div/div/div/span[2]/div/div/div/div[2]/div[2]/div[1]')))
            driver.execute_script("arguments[0].click();", nextmonthbtn)
            time.sleep(0.2)
            Endmonthandyearcycle(Year,Month,Date)
            return
        else:
            time.sleep(0.3)
            Enddateselectioncycle(Date,8)
            return

def Enddateselectioncycle(Date,DatePrg):
    driver.implicitly_wait(15)
    if CurrentLanguague == 'TH':
        selectingdate = driver.find_element(By.XPATH,'/html/body/div[1]/div/div/div[2]/div[2]/div[2]/div[2]/div/div[1]/div/div[3]/div/div/div/div/span[2]/div/div/div/div[2]/div[1]/div/div[2]/div['+ str(DatePrg) +']/span')
        valuenow = selectingdate.text
        driver.implicitly_wait(15)

        if valuenow == '':
            valuenow = 0
            

        if (int(valuenow) == int(Date)) == False:
            if int(valuenow) > int(Date):
                DatePrg = DatePrg - 1
            if int(valuenow) < int(Date):
                DatePrg = DatePrg + 1
            
            time.sleep(0.2)
            Enddateselectioncycle(Date,DatePrg)
            return
        else:
            wait = WebDriverWait(driver, 10)
            selectingdate = wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div/div/div[2]/div[2]/div[2]/div[2]/div/div[1]/div/div[3]/div/div/div/div/span[2]/div/div/div/div[2]/div[1]/div/div[2]/div['+ str(DatePrg) +']/span')))
            driver.execute_script("arguments[0].click();", selectingdate)
            driver.implicitly_wait(15)
            time.sleep(0.3)
            if dateselectionmode == False:
                searchbtn = driver.find_element(By.XPATH,'/html/body/div[1]/div/div/div[2]/div[2]/div[2]/div[2]/div/div[1]/div/button[3]')
                driver.execute_script("arguments[0].click();", searchbtn)
                driver.implicitly_wait(15)
                return
    else:
        selectingdate = driver.find_element(By.XPATH,'/html/body/div[1]/div/div/div[2]/div[2]/div[2]/div/div/div[1]/div/div[3]/div/div/div/div/span[2]/div/div/div/div[2]/div[1]/div/div[2]/div['+ str(DatePrg) +']/span')
        valuenow = selectingdate.text
        driver.implicitly_wait(15)

        if valuenow == '':
            valuenow = 0
            

        if (int(valuenow) == int(Date)) == False:
            if int(valuenow) > int(Date):
                DatePrg = DatePrg - 1
            if int(valuenow) < int(Date):
                DatePrg = DatePrg + 1
            
            time.sleep(0.2)
            Enddateselectioncycle(Date,DatePrg)
            return
        else:
            wait = WebDriverWait(driver, 10)
            selectingdate = wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div/div/div[2]/div[2]/div[2]/div/div/div[1]/div/div[3]/div/div/div/div/span[2]/div/div/div/div[2]/div[1]/div/div[2]/div['+ str(DatePrg) +']/span')))
            driver.execute_script("arguments[0].click();", selectingdate)
            driver.implicitly_wait(15)
            time.sleep(0.3)
            if dateselectionmode == False:
                searchbtn = driver.find_element(By.XPATH,'/html/body/div[1]/div/div/div[2]/div[2]/div[2]/div/div/div[1]/div/button[3]')
                driver.execute_script("arguments[0].click();", searchbtn)
                driver.implicitly_wait(15)
                return
        
#Select Start Date
def dateselection(Year,Month,Date):
    if dateselectionmode == False:
        return
    
    if CurrentLanguague == 'TH':
        calendarbtn = driver.find_element(By.XPATH, '/html/body/div[1]/div/div/div[2]/div[2]/div[2]/div[2]/div/div[1]/div/div[2]/div/div/div/span[2]/input')
        driver.execute_script("arguments[0].click();", calendarbtn)
        driver.implicitly_wait(15)
        MonthList = ['มกราคม','กุมภาพันธ์','มีนาคม','เมษายน','พฤษภาคม','มิถุนายน','กรกฎาคม','สิงหาคม','กันยายน','ตุลาคม','พฤศจิกายน','ธันวาคม']
        Month = MonthList[int(int(Month)-1)]
        
    else:
        calendarbtn = driver.find_element(By.XPATH, '/html/body/div[1]/div/div/div[2]/div[2]/div[2]/div/div/div[1]/div/div[2]/div/div/div/span[2]/input')
        driver.execute_script("arguments[0].click();", calendarbtn)
        driver.implicitly_wait(15)
        MonthList = ['January','February','March','April','May','June','July','August','September','October','November','December']
        Month = MonthList[int(int(Month)-1)]
        Year = int(Year) - 543
    
    log_to_gui(app.log_box, "Searching through Start Date")
    monthandyearcycle(Year,Month,Date)

def monthandyearcycle(Year,Month,Date):
    driver.implicitly_wait(15)
    Target = str(Month) + " " + str(Year) 
    if CurrentLanguague == 'TH':
        monthanyear = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div/div/div[2]/div[2]/div[2]/div[2]/div/div[1]/div/div[2]/div/div/div/span[2]/div/div/div/div[2]/div[1]/div/div[1]')))
        valuenow = monthanyear.text
        time.sleep(0.2)
        if (valuenow == Target) == False:
            nextmonthbtn = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[1]/div/div/div[2]/div[2]/div[2]/div[2]/div/div[1]/div/div[2]/div/div/div/span[2]/div/div/div/div[2]/div[2]/div[2]')))
            driver.execute_script("arguments[0].click();", nextmonthbtn)
            time.sleep(0.2)
            monthandyearcycle(Year,Month,Date)
            return
        else:
            time.sleep(0.3)
            dateselectioncycle(Date,8)
            return
    else:
        monthanyear = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div/div/div[2]/div[2]/div[2]/div/div/div[1]/div/div[2]/div/div/div/span[2]/div/div/div/div[2]/div[1]/div/div[1]/div')))
        valuenow = monthanyear.text
        time.sleep(0.2)
        if (valuenow == Target) == False:
            nextmonthbtn = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[1]/div/div/div[2]/div[2]/div[2]/div/div/div[1]/div/div[2]/div/div/div/span[2]/div/div/div/div[2]/div[2]/div[2]')))
            driver.execute_script("arguments[0].click();", nextmonthbtn)
            time.sleep(0.2)
            monthandyearcycle(Year,Month,Date)
            return
        else:
            time.sleep(0.3)
            dateselectioncycle(Date,8)
            return

def dateselectioncycle(Date,DatePrg):
    driver.implicitly_wait(15)
    if CurrentLanguague == 'TH':
        selectingdate = driver.find_element(By.XPATH,'/html/body/div[1]/div/div/div[2]/div[2]/div[2]/div[2]/div/div[1]/div/div[2]/div/div/div/span[2]/div/div/div/div[2]/div[1]/div/div[2]/div['+ str(DatePrg) +']/span')
        valuenow = selectingdate.text
        driver.implicitly_wait(15)

        if valuenow == '':
            valuenow = 0
            

        if (int(valuenow) == int(Date)) == False:
            if int(valuenow) > int(Date):
                DatePrg = DatePrg - 1
            if int(valuenow) < int(Date):
                DatePrg = DatePrg + 1
            
            time.sleep(0.2)
            dateselectioncycle(Date,DatePrg)
            return
        else:
            wait = WebDriverWait(driver, 10)
            selectingdate = wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div/div/div[2]/div[2]/div[2]/div[2]/div/div[1]/div/div[2]/div/div/div/span[2]/div/div/div/div[2]/div[1]/div/div[2]/div['+ str(DatePrg) +']/span')))
            driver.execute_script("arguments[0].click();", selectingdate)
            driver.implicitly_wait(15)
            time.sleep(0.3)
            searchbtn = driver.find_element(By.XPATH,'/html/body/div[1]/div/div/div[2]/div[2]/div[2]/div[2]/div/div[1]/div/button[3]')
            driver.execute_script("arguments[0].click();", searchbtn)
            driver.implicitly_wait(15)
            return
    else:
        selectingdate = driver.find_element(By.XPATH,'/html/body/div[1]/div/div/div[2]/div[2]/div[2]/div/div/div[1]/div/div[2]/div/div/div/span[2]/div/div/div/div[2]/div[1]/div/div[2]/div['+ str(DatePrg) +']/span')
        valuenow = selectingdate.text
        driver.implicitly_wait(15)

        if valuenow == '':
            valuenow = 0
            

        if (int(valuenow) == int(Date)) == False:
            if int(valuenow) > int(Date):
                DatePrg = DatePrg - 1
            if int(valuenow) < int(Date):
                DatePrg = DatePrg + 1
            
            time.sleep(0.2)
            dateselectioncycle(Date,DatePrg)
            return
        else:
            wait = WebDriverWait(driver, 10)
            selectingdate = wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div/div/div[2]/div[2]/div[2]/div/div/div[1]/div/div[2]/div/div/div/span[2]/div/div/div/div[2]/div[1]/div/div[2]/div['+ str(DatePrg) +']/span')))
            driver.execute_script("arguments[0].click();", selectingdate)
            driver.implicitly_wait(15)
            time.sleep(0.3)
            searchbtn = driver.find_element(By.XPATH,'/html/body/div[1]/div/div/div[2]/div[2]/div[2]/div/div/div[1]/div/button[3]')
            driver.execute_script("arguments[0].click();", searchbtn)
            driver.implicitly_wait(15)
            return
"""

#New Date Selection Method
def DirectDateSearch():
    if newscardamt == 0:
        return
    
    for i in range(newscardamt):
        if CurrentLanguague == 'TH':
            try:
                Datepointer = driver.find_element(By.XPATH,'/html/body/div[1]/div/div/div[2]/div[2]/div[2]/div[2]/div/div[2]/div[2]/div[' + str(i+1) + ']/div/div[1]/div[1]/div[1]/span[1]')
                StartYear = keyedYear
                EndYear = keyedEndYear
            except:
                return

        else:
            try:
                Datepointer = driver.find_element(By.XPATH,'/html/body/div[1]/div/div/div[2]/div[2]/div[2]/div/div/div[2]/div[2]/div[' + str(i+1) + ']/div/div[1]/div[1]/div[1]/span[1]')
                StartYear = keyedYear - 543
                EndYear = keyedEndYear - 543
            except:
                return
        
        DateText = Datepointer.text
        DateText = DateText.split()

        PointerDate = DateText[0]
        

        if PointerDate == 'ข่าววันนี้':
            MonthList = ['ม.ค.','ก.พ.','มี.ค.','เม.ย.','พ.ค.','มิ.ย.','ก.ค.','ส.ค.','ก.ย.','ต.ค.','พ.ย.','ธ.ค.']
            PointerDate = datetime.now().strftime("%d")
            PointerMonth = datetime.now().strftime("%m")
            PointerYear = datetime.now().strftime("%Y")
            PointerYear = (int(keyedEndYear) + 543)
            PointerMonth = MonthList[int(int(PointerMonth) - 1)]
        elif PointerDate == 'Today':
            PointerDate = datetime.now().strftime("%d")
            PointerMonth = datetime.now().strftime("%m")
            PointerYear = datetime.now().strftime("%Y")
            MonthList = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']
            PointerMonth = MonthList[int(int(PointerMonth) - 1)]
        else:
            PointerMonth = DateText[1]
            PointerYear = int(DateText[2])


        if CurrentLanguague == 'TH':
            MonthList = ['ม.ค.','ก.พ.','มี.ค.','เม.ย.','พ.ค.','มิ.ย.','ก.ค.','ส.ค.','ก.ย.','ต.ค.','พ.ย.','ธ.ค.']
            PointerMonth = int(MonthList.index(PointerMonth) + 1)

        else:
            MonthList = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']
            PointerMonth = int(MonthList.index(PointerMonth) + 1)


        

        if int(PointerYear) >= int(StartYear) and int(PointerYear) <= int(EndYear):
            if PointerMonth > int(keyedMonth) and PointerMonth < int(keyedEndMonth):
                GetNewsPageLinks(i+1)

            elif PointerMonth == int(keyedEndMonth):
                if int(PointerDate) <= int(keyedEndDate):
                    GetNewsPageLinks(i+1)
                else:
                    pass
            
            elif PointerMonth == int(keyedMonth):
                if int(PointerDate) >= int(keyedDate):
                    GetNewsPageLinks(i+1)
                else:
                    pass

            else:
                pass
        else:
            #Increase Speed by int(PointerYear) < int(StartYear) then return
            pass
    
#Get more News
def GetMoreNews():
    if CurrentLanguague == 'TH':
        try:
            morenews = driver.find_element(By.XPATH,'/html/body/div[1]/div/div/div[2]/div[2]/div[2]/div[2]/div/div[2]/div[2]/div[6]/div')
            driver.execute_script("arguments[0].click();", morenews)
        except:
            print("No More News")
            log_to_gui(app.log_box, "No Extra News")
            return
    else: 
        try:
            morenews = driver.find_element(By.XPATH,'/html/body/div[1]/div/div/div[2]/div[2]/div[2]/div/div/div[2]/div[2]/div[6]/div')
            driver.execute_script("arguments[0].click();", morenews)
        except:
            print("No More News")
            log_to_gui(app.log_box, "No Extra News")
            return
    
    driver.delete_all_cookies()
    driver.execute_script("window.localStorage.clear();")
    driver.execute_script("window.sessionStorage.clear();")

#Get all news cards
def GetNewsCardsAmt():
    global newscardamt
    if CurrentLanguague == 'TH':
        try:
            wait = WebDriverWait(driver, 10)
            container_xpath = "/html/body/div[1]/div/div/div[2]/div[2]/div[2]/div[2]/div/div[2]/div[2]"
            # Wait for the specific container
            container = wait.until(EC.presence_of_element_located((By.XPATH, container_xpath)))
            # Now only find cards inside that container
            news_cards = container.find_elements(By.CSS_SELECTOR, "div.card-quote-news")
            Cards = len(news_cards)
            newscardamt += Cards
            print('Card Found: ' + str(Cards))
            log_to_gui(app.log_box, 'News Card Found: ' + str(Cards))
            driver.implicitly_wait(15)
        except:
            print("Set Card")
            log_to_gui(app.log_box, 'no news found')
            newscardamt = 0
            return
    else:
        try:
            wait = WebDriverWait(driver, 10)
            container_xpath = "/html/body/div[1]/div/div/div[2]/div[2]/div[2]/div/div/div[2]/div[2]"
            # Wait for the specific container
            container = wait.until(EC.presence_of_element_located((By.XPATH, container_xpath)))
            # Now only find cards inside that container
            news_cards = container.find_elements(By.CSS_SELECTOR, "div.card-quote-news")
            Cards = len(news_cards)
            newscardamt += Cards
            print('Card Found: ' + str(Cards))
            log_to_gui(app.log_box, 'News Card Found: ' + str(Cards))
            driver.implicitly_wait(15)
        except:
            print("Set Card")
            log_to_gui(app.log_box, 'no news found')
            newscardamt = 0
            return

#Loop till all of news are all read
#Find News Page link
def GetNewsPageLinks(location):
    global ListofLinks
    if newscardamt == 0:
        return


    if CurrentLanguague == 'TH':
        newloc = driver.find_element(By.XPATH,'/html/body/div[1]/div/div/div[2]/div[2]/div[2]/div[2]/div/div[2]/div[2]/div['+ str(location) +']/div/div[1]/div[2]/div[1]/div/div/div/ul/li[2]/a')
    else:
        newloc = driver.find_element(By.XPATH,'/html/body/div[1]/div/div/div[2]/div[2]/div[2]/div/div/div[2]/div[2]/div['+ str(location) +']/div/div[1]/div[2]/div[1]/div/div/div/ul/li[2]/a')
    
    time.sleep(0.2)
    fullink = newloc.get_attribute('href')
    eqlindex = fullink.index('=')
    ContinueNewsLink = fullink[eqlindex+1:]
    ListofLinks.append(ContinueNewsLink)
    driver.implicitly_wait(15)
    time.sleep(0.1)

def GetandReadNews():
    print("Working")
    global ListofLinks
    data = []

    ContinueNewsLink = ListofLinks[0]
    print(ListofLinks[0])

    #Get news data
    global DataDict
    #Into News Page
    driver.get(ContinueNewsLink)
    driver.implicitly_wait(15)
    name = driver.find_element(By.XPATH,'/html/body/div[1]/div/div/div[2]/div/div/div/div[3]/div[2]/div[1]/span')
    head = driver.find_element(By.XPATH,'/html/body/div[1]/div/div/div[2]/div/div/div/div[3]/div[1]/h2')
    date = driver.find_element(By.XPATH,'/html/body/div[1]/div/div/div[2]/div/div/div/div[1]/div[1]/span/span')
    source = driver.find_element(By.XPATH,'/html/body/div[1]/div/div/div[2]/div/div/div/div[3]/div[2]/div[2]/span')
    data.append(name.text)
    data.append(head.text)
    data.append(date.text)
    data.append(source.text)
    driver.implicitly_wait(15)

    print("Company Name: " + name.text + " ")
    print("News Header: " + head.text + " ")
    log_to_gui(app.log_box, "Working on: ")
    log_to_gui(app.log_box, "Company Name: " + name.text + " ")
    log_to_gui(app.log_box, "News Header: " + head.text + " ")
    
    newsarticle = driver.find_element(By.XPATH,'/html/body/div[1]/div/div/div[2]/div/div/div/div[4]/div/div/div/pre/div')
    ReadNews = newsarticle.text
    ReadNews = ReadNews.replace('\n', ' ')
    data.append(ReadNews)
    driver.implicitly_wait(15)

    #Count Keywords
    for words in Keywords:
        amt = ReadNews.count(words)
        data.append(amt)

    for words in KeywordsEN:
        amt = ReadNews.count(words)
        data.append(amt)

    DataDict.append(data)
    ListofLinks.pop(0)
    time.sleep(0.2)



#Export to excel
def ExcelExport():
    global DataDict
    outdata=pd.DataFrame(DataDict,columns=['Company','news','date','source','body'] + Keywords + KeywordsEN)
    with pd.ExcelWriter(file_path) as writer:
        outdata.to_excel(writer, sheet_name='Company_Value', index=False)
    print('DataFrames are written to Excel File successfully.')
    DataDict = []

#Loop 
def recursiveLoop():
    if stopcommand == True:
        return
    
    global pageprg
    global companyprg
    global newsprg
    global CompanyList
    global CurrentLanguague

    if int(pageprg) > int(maxpageamt):
        ExcelExport()
        print('END')
        return()

    if int(companyprg) > int(maxcompanyamt):
        CompanyList = []
        pageprg = pageprg + 1
        companyprg = 1

        if int(pageprg) > int(maxpageamt):
            recursiveLoop()
            return()

        log_to_gui(app.log_box, 'Page: '+ str(pageprg))
        log_to_gui(app.log_box, 'Company Number: '+ str(companyprg))

        driver.get("https://www.set.or.th/th/market/get-quote/stock/")
        driver.implicitly_wait(15)
        SelectSubCompanyListPage()
        GetMaxCompanyAmt()
        GetPageCompany()
        if len(Keywords) > 0:
            CurrentLanguague = 'TH'
            EnterCompanyPage(companyprg)
            GetMoreNews()
            GetNewsCardsAmt()
            DirectDateSearch()

        if len(KeywordsEN) > 0:
            CurrentLanguague = 'EN'
            EnterCompanyPage(companyprg)
            GetMoreNews()
            GetNewsCardsAmt()
            DirectDateSearch()
        
        setnewscardamt()
        recursiveLoop()
        return()

    if int(newsprg) > int(newscardamt):
        companyprg = companyprg + 1
        newsprg = 1
        

        if int(companyprg) > int(maxcompanyamt):
            recursiveLoop()
            return()

        log_to_gui(app.log_box, 'Page: '+ str(pageprg))
        log_to_gui(app.log_box, 'Company Number: '+ str(companyprg))

        if len(Keywords) > 0:
            CurrentLanguague = 'TH'
            EnterCompanyPage(companyprg)
            GetMoreNews()
            GetNewsCardsAmt()
            DirectDateSearch()
        
        if len(KeywordsEN) > 0:
            CurrentLanguague = 'EN'
            EnterCompanyPage(companyprg)
            GetMoreNews()
            GetNewsCardsAmt()
            DirectDateSearch()
            
        setnewscardamt()
        recursiveLoop()
        return()


    GetandReadNews()
    driver.implicitly_wait(50)
    newsprg = newsprg + 1
    time.sleep(0.2)
    recursiveLoop()

#Start function and process
def StartScrapingProcess():
    global CurrentLanguague
    GetMaxPageAmt()
    GetMaxCompanyAmt()
    GetPageCompany()

    if len(Keywords) > 0:
        CurrentLanguague = 'TH'
        EnterCompanyPage(companyprg)
        GetMoreNews()
        GetNewsCardsAmt()
        DirectDateSearch()
        
    
    if len(KeywordsEN) > 0:
        CurrentLanguague = 'EN'
        EnterCompanyPage(companyprg)
        GetMoreNews()
        GetNewsCardsAmt()
        DirectDateSearch()
        
    setnewscardamt()
    recursiveLoop()
    #For Debug 
    #input("Press Enter to exit and close browser...")  # Keeps window open until you press Enter
    #driver.quit()

class App:
    # ========== Tkinter ==========

    def __init__(self, root):
        self.root = root
        self.root.title("SuperScraper Panel")
        self.running = False

        # Variables
        self.vcmd = (self.root.register(self.validate_int), '%P')
        self.keywords_var = tk.StringVar()
        self.keywords_varEN = tk.StringVar()
        self.keywords_list = []

        self.date_var = tk.BooleanVar()
        self.end_date_var = tk.BooleanVar()  # Added for End Date
        self.amount_var = tk.BooleanVar()

        self.selected_day = tk.StringVar()
        self.selected_month = tk.StringVar()
        self.selected_year = tk.StringVar()

        self.selected_end_day = tk.StringVar()
        self.selected_end_month = tk.StringVar()
        self.selected_end_year = tk.StringVar()

        self.amount = tk.StringVar()

        self.build_ui()
    

    def validate_int(self, value_if_allowed):
        return value_if_allowed == "" or value_if_allowed.isdigit()
    


    def build_ui(self):
       # ========== Start Date Selection ==========
        tk.Checkbutton(self.root, text="Enable Date Selection", variable=self.date_var, command=self.toggle_date_widgets).pack(anchor='w')
        self.date_frame = tk.Frame(self.root)
        self.day_cb = ttk.Combobox(self.date_frame, textvariable=self.selected_day, values=[], width=5, state='disabled')
        self.month_cb = ttk.Combobox(self.date_frame, textvariable=self.selected_month, values=[str(i) for i in range(1, 13)], width=5, state='disabled')
        current_year = datetime.now().year
        self.year_cb = ttk.Combobox(self.date_frame, textvariable=self.selected_year, values=[str(current_year - i) for i in range(6)], width=7, state='disabled')

        self.month_cb.bind('<<ComboboxSelected>>', self.update_days)
        self.year_cb.bind('<<ComboboxSelected>>', self.update_months)

        self.year_cb.pack(side='left')
        self.month_cb.pack(side='left', padx=5)
        self.day_cb.pack(side='left')
        self.date_frame.pack(pady=5)

        # ========== End Date Selection ==========
        tk.Checkbutton(self.root, text="Enable End Date Selection", variable=self.end_date_var, command=self.toggle_end_date_widgets).pack(anchor='w')
        self.end_date_frame = tk.Frame(self.root)
        self.end_day_cb = ttk.Combobox(self.end_date_frame, textvariable=self.selected_end_day, values=[], width=5, state='disabled')
        self.end_month_cb = ttk.Combobox(self.end_date_frame, textvariable=self.selected_end_month, values=[str(i) for i in range(1, 13)], width=5, state='disabled')
        self.end_year_cb = ttk.Combobox(self.end_date_frame, textvariable=self.selected_end_year, values=[str(current_year - i) for i in range(6)], width=7, state='disabled')

        self.end_month_cb.bind('<<ComboboxSelected>>', self.update_end_days)
        self.end_year_cb.bind('<<ComboboxSelected>>', self.update_end_months)

        self.end_year_cb.pack(side='left')
        self.end_month_cb.pack(side='left', padx=5)
        self.end_day_cb.pack(side='left')
        self.end_date_frame.pack(pady=5)

        """
        # ========== Amount Entry ==========
        tk.Checkbutton(self.root, text="Enable Card Amount Set", variable=self.amount_var, command=self.toggle_amount_widget).pack(anchor='w')
        self.amount_frame = tk.Frame(self.root)
        tk.Label(self.amount_frame, text="Card Amount:").pack(side='left')
        self.amount_entry = tk.Entry(self.amount_frame, textvariable=self.amount, width=10, state='disabled', validate="key", validatecommand=self.vcmd)
        self.amount_entry.pack(side='left', padx=5)
        self.amount_frame.pack(pady=5)
        """

        # ========== Keyword Entry ==========
        keyword_frame = tk.Frame(self.root)
        tk.Label(keyword_frame, text="Keywords (comma separated):").pack(side='left')
        self.keyword_entry = tk.Entry(keyword_frame, width=40)
        self.keyword_entry.pack(side='left', padx=5)
        keyword_frame.pack(pady=5)

        # ========== Keyword Entry EN ==========
        keyword_frameEN = tk.Frame(self.root)
        tk.Label(keyword_frameEN, text="KeywordsEN (comma separated):").pack(side='left')
        self.keyword_entryEN = tk.Entry(keyword_frameEN, width=40)
        self.keyword_entryEN.pack(side='left', padx=5)
        keyword_frameEN.pack(pady=5)

        # ========== Buttons ==========
        button_frame = tk.Frame(self.root)
        self.start_button = tk.Button(button_frame, text="Start", command=self.start_task)
        self.stop_button = tk.Button(button_frame, text="Stop", command=self.stop_task, state='disabled')
        self.start_button.pack(side='left', padx=5)
        self.stop_button.pack(side='left', padx=5)
        button_frame.pack(pady=10)

        # ========== Log Box ==========
        self.log_box = tk.Text(self.root, height=12, width=60)
        self.log_box.pack(pady=5)

    def update_days(self, *args):
        try:
            year = int(self.selected_year.get())
            month = int(self.selected_month.get())
            today = datetime.today()
            min_date = datetime(today.year - 6 + 1, today.month, today.day)

            days = []
            for day in range(1, calendar.monthrange(year, month)[1] + 1):
                try:
                    candidate = datetime(year, month, day)
                    if min_date <= candidate <= today:
                        days.append(str(day))
                except:
                    continue

            self.day_cb['values'] = days

            if self.selected_day.get() not in days:
                self.selected_day.set(days[0] if days else "")
        except:
            pass


    def update_end_days(self, *args):
        try:
            year = int(self.selected_end_year.get())
            month = int(self.selected_end_month.get())
            today = datetime.today()
            min_date = datetime(today.year - 6 + 1, today.month, today.day)

            days = []
            for day in range(1, calendar.monthrange(year, month)[1] + 1):
                try:
                    candidate = datetime(year, month, day)
                    if min_date <= candidate <= today:
                        days.append(str(day))
                except:
                    continue

            self.end_day_cb['values'] = days

            if self.selected_end_day.get() not in days:
                self.selected_end_day.set(days[0] if days else "")
        except:
            pass
    
    def update_months(self, *args):
        try:
            year = int(self.selected_year.get())
            today = datetime.today()

            if year == today.year:
                months = [str(m) for m in range(1, today.month + 1)]
            else:
                months = [str(m) for m in range(1, 13)]

            self.month_cb['values'] = months

            if self.selected_month.get() not in months:
                self.selected_month.set(months[0] if months else "")

            self.update_days()
        except:
            pass

    def update_end_months(self, *args):
        try:
            year = int(self.selected_end_year.get())
            today = datetime.today()

            if year == today.year:
                months = [str(m) for m in range(1, today.month + 1)]
            else:
                months = [str(m) for m in range(1, 13)]

            self.end_month_cb['values'] = months

            if self.selected_end_month.get() not in months:
                self.selected_end_month.set(months[0] if months else "")

            self.update_end_days()
        except:
            pass

    def toggle_date_widgets(self):
        state = "normal" if self.date_var.get() else "disabled"
        self.day_cb.config(state=state)
        self.month_cb.config(state=state)
        self.year_cb.config(state=state)
    
    def toggle_end_date_widgets(self):
        state = "normal" if self.end_date_var.get() else "disabled"
        self.end_day_cb.config(state=state)
        self.end_month_cb.config(state=state)
        self.end_year_cb.config(state=state)

    def toggle_amount_widget(self):
        state = "normal" if self.amount_var.get() else "disabled"
        self.amount_entry.config(state=state)

    def log(self, message):
        self.log_box.insert(tk.END, message + "\n")
        self.log_box.see(tk.END)

    def start_task(self):
        if not self.running:
            self.running = True
            self.thread = threading.Thread(target=self.run_task)
            self.thread.start()


    def stop_task(self):
        global stopcommand
        stopcommand = True
        self.running = False
        self.log("Stopping task...")
        self.log("Exporting to excel")
        ExcelExport()
        self.log("---Ended---")
        self.start_button.config(state='normal')
        self.stop_button.config(state='disabled')


    def run_task(self):
        while self.running:
            global dateselectionmode
            global Enddateselectionmode
            #global newscardamtoveride
            global stopcommand

            stopcommand = False

            if self.date_var.get():
                dateselectionmode = True
                day = self.selected_day.get()
                month = self.selected_month.get()
                year = self.selected_year.get()
                self.log(f"Selected Date: {day}-{month}-{year}")
                setdate(day, month, year)
            else:
                dateselectionmode = False
            
            if self.end_date_var.get():
                end_day = self.selected_end_day.get()
                end_month = self.selected_end_month.get()
                end_year = self.selected_end_year.get()
                self.log(f"Selected End Date: {end_day}-{end_month}-{end_year}")
                setEnddate(end_day, end_month, end_year)
            else:
                Enddateselectionmode = False

            """
            if self.amount_var.get():
                card_amt = self.amount.get()
                self.log(f"Card Amount: {card_amt}")
                global newscardamtoveride
                if card_amt == '':
                    newscardamtoveride = 2
                else:
                    newscardamtoveride = int(card_amt)
            else:
                newscardamtoveride = 0
            """

            keyword_input = self.keyword_entry.get()
            self.keywords = [kw.strip() for kw in keyword_input.split(',') if kw.strip()]
            self.log(f"Keywords: {self.keywords}")

            keyword_inputEN = self.keyword_entryEN.get()
            self.keywordsEN = [kw.strip() for kw in keyword_inputEN.split(',') if kw.strip()]
            self.log(f"KeywordsEN: {self.keywordsEN}")

            global Keywords
            global KeywordsEN
            Keywords = self.keywords
            KeywordsEN = self.keywordsEN

            if len(Keywords) <= 0 and len(KeywordsEN) <= 0:
                self.log("Please Enter Keywords")
                self.running = False
                break
                
            self.start_button.config(state='disabled')
            self.stop_button.config(state='normal')
            self.log("Running task step...")
            self.log("Please Check cmd for running steps")
            StartScrapingProcess()
            self.log("Excel Exported")
            self.log("Process - Ended")
            self.start_button.config(state='normal')
            self.stop_button.config(state='disabled')
            break  # remove or adjust this if task is continuous

# ========== Run App ==========
if __name__ == "__main__":
    root = tk.Tk()
    
    app = App(root)
    root.mainloop()



