import time
import sys,os, requests
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import selenium.webdriver.chrome.service as service
from time import gmtime, strftime
from datetime import datetime, timedelta, date
from six.moves.html_parser import HTMLParser
from lxml.html import fromstring
from selenium.webdriver.chrome.options import Options
from time import sleep
import xlrd
import re
import html


chrome_options = Options()
chrome_options.add_extension('Ultrasurf.crx')
driver = webdriver.Chrome(options=chrome_options, executable_path='chromedriver.exe')
driver.delete_all_cookies()

def junck(value):
    value = re.sub(r"\[","",str(value))
    value = re.sub(r"\]","",str(value))
    value = re.sub(r"\"","",str(value))
    value = re.sub(r"\'","",str(value))
    value = re.sub(r"\(","",str(value))
    value = re.sub(r"\)","",str(value))
    # value = re.sub(r"\\\\","",str(value))
    # value = re.sub(r"\@","",str(value))
    value = re.sub(r"\   ","",str(value))
    return value
	
excel_sheet = xlrd.open_workbook("Input.xlsx")
sheet1= excel_sheet.sheet_by_name('Sheet1')



for i in range(0, sheet1.nrows):        
    row = sheet1.row_slice(i)        
    Id = row[0].value 
    Sku = row[1].value    
    Link = row[2].value
    print(Sku)
    # time.sleep(5)

    try:    
        response = driver.get(str(Link))
        content1 = driver.page_source
        content = content1.encode('utf-8')
        content = (content1,'utf-8')
        content = re.sub(r"\\n", "",str(content))
        content = re.sub(r"\\t", "",str(content))
        content=re.sub(r'&amp;','&',str(content)) 

        file_name=str(Id)+".html"
        print(file_name)
        driver.delete_all_cookies()	
        # time.sleep(2)
        # a =open(file_name,"w")
        # a.write(content)
        # a.close()

        Address=re.findall('\<meta\ content\=\"\s*([^>]*?)\s*\"\ itemprop\=\"description\"\>', str(content), re.I)
        Address = junck(Address)
        print(Address, "\n")
        	
        f=open("Output.txt", 'a', encoding="utf-8")
        f.write(str(Id)+"\t"+str(Sku)+"\t"+str(Link)+"\t"+str(Address)+"\n")
        f.close()
        time.sleep(2)
    except:
        f=open("No_Data.txt", 'a', encoding="utf-8")
        f.write(str(Id)+"\t"+str(Sku)+"\t"+str(Link)+"\n")
        f.close()
        print("No Data")