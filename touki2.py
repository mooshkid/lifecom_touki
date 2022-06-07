from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import TimeoutException
import pandas as pd
import time
import os
import sys
import logging

# logging config
path = os.getcwd()
logPath = os.path.join(path, "Logs/touki.log")

logging.basicConfig(
    level=logging.INFO,
    format=u'%(asctime)s [%(levelname)s] %(message)s',
    handlers=[
        logging.FileHandler(logPath, encoding='utf8', mode='a'),
        logging.StreamHandler()
    ]
)
class LoggerWriter:
    def __init__(self, level):
        self.level = level
    def write(self, message):
        if message != '\n':
            self.level(message)
    def flush(self): pass
log = logging.getLogger(__name__)
sys.stdout = LoggerWriter(log.debug)
sys.stderr = LoggerWriter(log.error)

# start time
start = time.time()
# counter
count = 0


##### CHANGE EXCEL PATH HERE #####
excelPath = os.path.join(path, 'touki.xlsx')
# dataframes
df=pd.read_excel(excelPath, 'script2 ', names=['pre', 'text','#'])

#error list
error_list = []
#error .txt
with open(os.path.join(path, "Logs/touki_error.txt"), 'a', encoding="utf-8") as f:
    f.write("\n" + time.strftime("%B %d, %Y %H:%M:%S"))

# chrome options !Remember to set options to YOUR profile (chrome://version)
options = webdriver.ChromeOptions()
options.add_argument('--user-data-dir=C:\\Users\\kokoku\\AppData\\Local\\Google\\Chrome\\User Data')
options.add_argument('--profile-directory=Default')
# options.add_argument('--user-data-dir=C:\\Users\\yamanaka\\AppData\\Local\\Google\\Chrome\\User Data')
# options.add_argument('--profile-directory=Profile 8')
options.add_argument("start-maximized")
driver = webdriver.Chrome(options=options)


### START MAIN SCRIPT ###
# launch URL
driver.get("https://www.touki.or.jp/TeikyoUketsuke/")

# login
loginButton = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="wrapper"]/div[3]/div[3]/div[1]/form/div[3]/button')))
loginButton.click()

# if already logged in on somewhere else
try:
   forceLogin = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="mainBox"]/div[2]/form/button[2]')))
   forceLogin.click()
except TimeoutException:
    pass

# real estate page
realEstate = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="uketsukeMenuFrom"]/div/h5[2]/span[1]/a')))
realEstate.click()


### BILLING PAGE ###
for i, row in df.iterrows():
    r = row['pre']
    t = row['text']
    n = row['#']
    tn = t + " " + n
    count += 1
    
    log.info('Starting(' + str(count) + '): ' + tn)

#########################################
##### SELECT STATE HERE ~/option[x] #####
#########################################
    selectState = driver.find_element(By.XPATH, '//*[@id="fuTodofukenShozai"]/optgroup/*[contains(text(), "{}")]'.format(r))
    selectState.click()
    # check input manually
    checkManual = driver.find_element(By.XPATH, '//*[@id="fuShozaiChokusetuNyuryoku"]').click()
    # address text box
    addressTextBox = driver.find_element(By.XPATH, '//*[@id="fuChibanKuiki"]')
    addressTextBox.clear()
    addressTextBox.send_keys(t)

    # address number
    addressNumberBox = driver.find_element(By.XPATH, '//*[@id="fuChibanKaoku"]')
    addressNumberBox.clear()
    addressNumberBox.send_keys(n)

    # select options
    optionOne = driver.find_element(By.XPATH, '//*[@id="fuAll"]').click()
    optionTwo = driver.find_element(By.XPATH, '//*[@id="fuShoyusya"]').click()
    # confirm button
    confirmButton = driver.find_element(By.XPATH, '//*[@id="tabsFudosan"]/div[5]/button[2]')
    confirmButton.click()

    ### LIST PAGE ###
    try:
        checkListItem = WebDriverWait(driver, 1).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="sentaku_1"]')))
        checkListItem.click()

        billingButton = driver.find_element(By.XPATH, '//*[@id="wrapper"]/div[4]/div[6]/button[2]')
        billingButton.click()
        log.info('Billing Complete')
        # click ok
        okButton = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="top"]/div[6]/div[3]/button[2]')))
        okButton.click()
        
        ### PDF PAGE ###
        # refresh page
        refreshbutton = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="myReloadButton"]')))
        refreshbutton.click()
        time.sleep(2)
        
        try:
            refreshbutton2 = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/div[4]/div[5]/div[4]/form/div[3]/div[2]/div/div[2]/table/tbody/tr[1]/td[6]/span/button')))
            refreshbutton2.click()
            time.sleep(2)
        except TimeoutException:
            pass

        # check PDF Item
        checkPdf = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="myPageTable"]/tbody/tr[1]/td[1]/input'))).click()
        # download to folder
        downloadButton = driver.find_element(By.XPATH, '//*[@id="downloadButton"]/button').click()
        log.info('PDF Downloaded')

        ## return to real estate tab
        time.sleep(2)
        realEstateTab = driver.find_element(By.XPATH, '//*[@id="wrapper"]/div[4]/ul/li[2]/a')
        realEstateTab.click()
        

    except TimeoutException:
        log.warning('ADDRESS NOT FOUND: ' + tn)
        # append tn to error.txt
        with open(os.path.join(path, "Logs/touki_error.txt"), 'a', encoding="utf-8") as f:
            f.write("\n" + tn)
        
        # add failed address to list
        error_list.append(tn)
        # click clear button
        driver.find_element(By.XPATH, '//*[@id="tabsFudosan"]/div[5]/button[1]').click()
        pass

# log list of addresses not found
log.info(str(len(error_list)) + ' Addresses NOT found: ' + str(error_list))

# nf = pd.DataFrame(error_list)
# excelErrorPath = os.path.join(path, 'touki/touki_error.xlsx')
# nf.to_excel(excelErrorPath)
# log.info(nf)

# elasped time 
end = time.time()
elapsed = end - start
log.info('All ' + str(count) + ' Tasks Completed in: ' + time.strftime('%H:%M:%S', time.gmtime(elapsed)))

driver.close()