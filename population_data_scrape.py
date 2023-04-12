# -*- coding: utf-8 -*-
"""
Created on Wed Apr  5 19:16:16 2023

@author: Dom
"""

# import webdriver and relevant webdriver features for automated web browsing
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait as Wait
from selenium.webdriver.support import expected_conditions as EC

# import the Excel driver
import win32com.client as win32

# initialize Excel driver and data workbook
# workbook is invisible to the user
excel = win32.gencache.EnsureDispatch('Excel.Application')
excel.Visible = False
wb = excel.Workbooks.Open('2020-21 state population data.xlsx')
ws = wb.Worksheets('State population data')

# make the automated browser invisible to the user
chrome_options = Options()
chrome_options.add_argument('--headless')

# open invisible browser window and go to state population data source website
driver = webdriver.Chrome(options=chrome_options)
driver.get('https://data.ers.usda.gov/reports.aspx?ID=17827')

# wait for website to finish loading
# test this by constantly checking if a link with the text 'Puerto Rico' exists
ready = False
while ready == False:
    try:
        Wait(driver, 0.00001).until(
                EC.presence_of_element_located((By.LINK_TEXT, 'Puerto Rico')))
        ready = True
    except:
        pass

# all state names and the names 'United States', 'District of Columbia', and 'Puerto Rico',
# as well as the 2020/2021 population and % change numbers, are found in the text of 'td'
# HTML tags

# initialize variable to iterate over the list of all 'td' HTML tags on the page, stopping
# at the first tag with the text 'United States'
#
# this is the first element we will scrape from the first data row we scrape
iter_num = 0
while driver.find_elements(By.TAG_NAME, 'td')[iter_num].text != 'United States':
    iter_num += 1

# initialize variable for Excel spreadsheet row to copy data into
excel_row = 2

# scrape the state/territory name, 2020 population, 2021 population, and 2020-21 % change in
# population, then move on to the next data row on the web page and in the spreadsheet
#
# repeat until Puerto Rico (final state/territory to be scraped) is the most recently
# scraped territory
while driver.find_elements(By.TAG_NAME, 'td')[iter_num-7].text != 'Puerto Rico':
    ws.Cells(excel_row, 1).Value = driver.find_elements(By.TAG_NAME, 'td')[iter_num].text
    ws.Cells(excel_row, 2).Value = driver.find_elements(By.TAG_NAME, 'td')[iter_num+4].text
    ws.Cells(excel_row, 3).Value = driver.find_elements(By.TAG_NAME, 'td')[iter_num+5].text
    ws.Cells(excel_row, 4).Value = driver.find_elements(By.TAG_NAME, 'td')[iter_num+6].text
    iter_num += 7
    excel_row += 1

# test that the sum of the 50 states' populations and D.C.'s population in each year is
# equal to the total U.S. population scraped for that year
#
# otherwise, raise an exception
#
# close browser and save and close workbook in either case
try:
    ws.Range('E2').Formula = '=SUM(B3:B53)'
    ws.Range('F2').Formula = '=SUM(C3:C53)'
    if ws.Range('B2').Value != ws.Range('E2').Value:
        raise Exception("The total U.S. 2020 population in the workbook is not equal to the sum of the states' and D.C.'s 2020 populations.")
    if ws.Range('C2').Value != ws.Range('F2').Value:
        raise Exception("The total U.S. 2021 population in the workbook is not equal to the sum of the states' and D.C.'s 2021 populations.")
    ws.Range('E2').Formula = ''
    ws.Range('F2').Formula = ''
    driver.close()
    wb.Close(True)
except:
    ws.Range('E2').Formula = ''
    ws.Range('F2').Formula = ''
    driver.close()
    wb.Close(True)
