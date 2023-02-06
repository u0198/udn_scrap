# %%
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
import selenium
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

s=Service('C:/Users/chromedriver.exe')
driver = webdriver.Chrome(service=s)
driver.get("https://covid-19.nchc.org.tw/")
pause_time = 0.5

element = WebDriverWait(driver, 8000).until(
    EC.presence_of_element_located((By.XPATH, '/html/body/div[2]/div/div[3]/p/span/small'))
)

month = datetime.now().strftime("%m")
date = datetime.now().strftime("%d")

if int(month) < 10:
    month = month[1:]

if int(date) < 10:
    date = date[1:]

cases = element.text.split(" ")[1]

covid_data = f'{month}/{date},{cases}'

covid_data

# %%
with open('covid_case.csv', mode='a') as file:
    file.writelines(covid_data + '\n')

# %%
#credentials to the account
cred = ServiceAccountCredentials.from_json_keyfile_name(r'C:\Users\u0198\Desktop\AutoDev\udn_scrap\cred.json') ;
# authorize the clientsheet 
client = gspread.authorize(cred)

# %%
sh = client.open('coviddata')
worksheet = sh.worksheet('2022')

# worksheet.insert_row([month/date, cases], index=3)
worksheet.append_row([f'{month}/{date}', cases], table_range="A:A")


