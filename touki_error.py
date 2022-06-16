# This is for TESTING PURPOSES ONLY!!! 
# Confirm payment and downloading pdfs are disabled 
# Logging is removed
# Name the excel book 'test.xlsx'

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import pandas as pd
import time
import os


# change cwd to the script directory 
os.chdir(os.path.dirname(__file__))
path = os.getcwd()
print(path)

# counter
count = 0

# dataframes
df=pd.read_excel('test.xlsx', 'script', names=['pre', 'text','#'])
# create empty dataframe 
df3 = pd.DataFrame()

# chrome options !Remember to set options to YOUR profile (chrome://version)
options = webdriver.ChromeOptions()
# options.add_argument('--user-data-dir=C:\\Users\\kokoku\\AppData\\Local\\Google\\Chrome\\User Data')
# options.add_argument('--profile-directory=Default')
options.add_argument('--user-data-dir=C:\\Users\\yamanaka\\AppData\\Local\\Google\\Chrome\\User Data')
options.add_argument('--profile-directory=Profile 8')
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
    rtn = r + t + n
    count += 1
    
    print('Starting(' + str(count) + '): ' + rtn)

    # select prefecture 
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

    # # select options
    # optionOne = driver.find_element(By.XPATH, '//*[@id="fuAll"]').click()
    # optionTwo = driver.find_element(By.XPATH, '//*[@id="fuShoyusya"]').click()
    # confirm button
    confirmButton = driver.find_element(By.XPATH, '//*[@id="tabsFudosan"]/div[5]/button[2]')
    confirmButton.click()
    time.sleep(1)

# ### LIST PAGE ###
#     try:
#         checkListItem = WebDriverWait(driver, 1).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="sentaku_1"]')))
#         checkListItem.click()

#         billingButton = driver.find_element(By.XPATH, '//*[@id="wrapper"]/div[4]/div[6]/button[2]')
#         billingButton.click()
#         print('Billing Complete')
#         # click ok
#         okButton = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="top"]/div[6]/div[3]/button[2]')))
#         okButton.click()
        
#         ### PDF PAGE ###
#         # refresh page
#         refreshbutton = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="myReloadButton"]')))
#         refreshbutton.click()
#         time.sleep(2)
        
#         try:
#             refreshbutton2 = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[2]/div[4]/div[5]/div[4]/form/div[3]/div[2]/div/div[2]/table/tbody/tr[1]/td[6]/span/button')))
#             refreshbutton2.click()
#             time.sleep(2)
#         except TimeoutException:
#             pass

#         # check PDF Item
#         checkPdf = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="myPageTable"]/tbody/tr[1]/td[1]/input'))).click()
#         # download to folder
#         downloadButton = driver.find_element(By.XPATH, '//*[@id="downloadButton"]/button').click()
#         print('PDF Downloaded')

#         ## return to real estate tab
#         time.sleep(2)
#         realEstateTab = driver.find_element(By.XPATH, '//*[@id="wrapper"]/div[4]/ul/li[2]/a')
#         realEstateTab.click()

#         ### FOR TESTING PURPOSES ONLY | RETURNS TO LIST PAGE ###
#         time.sleep(1)
#         returnButton = driver.find_element(By.XPATH, '//*[@id="wrapper"]/div[4]/div[6]/button[1]').click()


#     except TimeoutException:
#         print('ADDRESS NOT FOUND: ' + rtn)

#         # create an error dictionary 
#         error_list = {'pre': r, 'text': t, '#': n}
#         # convert dictionary to dataframe 
#         df2 = pd.DataFrame(error_list, index=[0])

#         # append df2 to df3 
#         df3.append(df2, ignore_index=True)
#         # write to excel 
#         df3.to_excel('output.xlsx')

    if driver.find_element(By.XPATH, '//*[@id="fuErrMsgArea"]/ul/li/span[contains(text(), "請求できない所在です")]'):
        print('ADDRESS NOT FOUND: ' + rtn)

        # create an error dictionary 
        error_list = {'pre': r, 'text': t, '#': n}
        # convert dictionary to dataframe 
        df2 = pd.DataFrame(error_list, index=[0])

        # append df2 to df3 
        df3.append(df2, ignore_index=True)
        # write to excel 
        df3.to_excel('output.xlsx')

        # click clear button
        driver.find_element(By.XPATH, '//*[@id="tabsFudosan"]/div[5]/button[1]').click()

    else:
        # return to search 
        returnButton = driver.find_element(By.XPATH, '//*[@id="wrapper"]/div[4]/div[6]/button[1]').click()


print(df)

driver.close()