from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import ElementNotInteractableException
from selenium.common.exceptions import TimeoutException
import pandas as pd
import time


#dataframes
df = pd.read_excel('02_data.xlsx')

# address list 
address_list = df['所有者住所'].tolist()

print(address_list)

# zip code list 
zip_list = []

# driver 
options = webdriver.ChromeOptions()
options.add_argument('--user-data-dir=C:\\Users\\yamanaka\\AppData\\Local\\Google\\Chrome\\User Data')
options.add_argument('--profile-directory=Profile 7')
options.add_argument("start-maximized")
driver = webdriver.Chrome(options=options)

# launch URL
driver.get("https://www.google.com/maps")


# the loop 
for i in address_list:
    # search box 
    searchBox = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="searchboxinput"]')))
    searchBox.clear()
    searchBox.send_keys(i)
    # enter 
    findButton = driver.find_element(By.XPATH, '//*[@id="searchbox-searchbutton"]')
    findButton.click()

    try:
        # zip code 
        zipCode = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[3]/div[9]/div[9]/div/div/div[1]/div[2]/div/div[1]/div/div/div[2]/div[1]/div[1]/div[1]/h1/span[1]')))
        zipCode = str(zipCode.text)[:9]

        print(zipCode)
        zip_list.append(zipCode)

        with open('zip_list.txt', 'a', encoding="utf-8") as f:
            f.write("\n" + zipCode)


    except TimeoutException:
        print('blank')
        zip_list.append('blank')

        with open('zip_list.txt', 'a', encoding="utf-8") as f:
            f.write("\n" + "blank")

        
df['zip'] = zip_list
df.to_excel('02_data.xlsx')
print(df)
