# -*- coding: utf-8 -*-
"""
Created on Wed Jun 30 23:34:32 2021

@author: trim2
"""
from pandas import DataFrame
import pandas as pd
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
import pyautogui
import re
import time
import random
from selenium.webdriver.chrome.options import Options


q=[]
p=[]
count = 1
dff = DataFrame()
dfff = DataFrame()
df = pd.read_excel(r'C:\Users\User\OneDrive\研所檔案\Green Energy_安辰_7月.xlsx' ,  sheet_name= '7', usecols='A')
opts = Options()
opts.add_argument("--incognito")
driver = webdriver.Chrome(ChromeDriverManager().install(), chrome_options = opts)
a = df[0:5]
def search(data):
    try:
         global p
         global q
         global driver
         global coname
         global sconame
         global finame
         pyautogui.hotkey('ctrl', 't') 
         driver.switch_to.window(driver.window_handles[1])
         pyautogui.typewrite(data ,interval = 0.05)
         pyautogui.press('enter')
         time.sleep(random.randint(3, 6))      
         coname = driver.find_element_by_class_name('qrShPb').text
         sconame = str(coname)
         finame=re.sub(r'[^A-Za-z0-9 ]+', '', sconame)
         datacoll = []
         try:
            cname = len(driver.find_elements_by_class_name('GRkHZd'))
            for t in cname:
              
                  
                coname1 = driver.find_elements_by_class_name('LrzXr')[t].text
                sconame1 = str(coname1)
                finame1=re.sub(r'[^A-Za-z0-9 ]+', '', sconame1)
                classtype = driver.find_elements_by_class_name('GRkHZd')[t].text
                datacoll.append(classtype)
                if finame1 != '' and classtype =='上級機構：':
                  finalname = finame1 + ' '+finame
                  q.append(finalname)
                  time.sleep(random.randint(3, 6))
                else:
                  pass
            time.sleep(random.randint(3, 6))
            driver.close()
            datacoll.index('上級機構：') 


         except:
          if finame == '' :
             q.append('')
             p.append(data[0:-9])
             time.sleep(random.randint(3, 6))
             driver.close()
             time.sleep(random.randint(3, 6))
           
          else:
           q.append(finame)
           time.sleep(random.randint(3, 6))
           driver.close()
           time.sleep(random.randint(3, 6))
         
    except :
           q.append('')
           p.append(data[0:-9])
           time.sleep(random.randint(3, 6))
           driver.close()
           time.sleep(random.randint(3, 6))

for i in a.values:
     for c in i:
       c1 = str(c)
       if count%50 == 0:
          time.sleep(random.randint(700, 750))                                
          search(c1)
          count+=1
       else:
          search(c1)
          count+=1
dff = dff.append(q, ignore_index = True)
dff.to_excel(r'C:\Users\User\Desktop\fie\123.xlsx', index = False)     
dfff = dfff.append(p, ignore_index = True)
dfff.to_excel(r'C:\Users\User\Desktop\fie\321.xlsx', index = False) 

