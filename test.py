import selenium
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
import warnings, time
warnings.filterwarnings('ignore')

driver = webdriver.Chrome(executable_path='C:/Users/chromedriver.exe')
driver.get("https://covid-19.nchc.org.tw/")
pause_time = 0.5

# <span class="country_confirmed_percent"><small>本土病例 64972</small></span>
# /html/body/div[2]/div/div[3]/p/span/small

element = WebDriverWait(driver, 3000).until(
    EC.presence_of_element_located((By.XPATH, '/html/body/div[2]/div/div[3]/p/span/small'))
)

cases = element.text.split(" ")[1]

# covid_data = f'{month}/{date},{cases}'

cases