from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import pandas as pd
import time
import re

#dataframes
df = pd.read_excel('data.xlsx')

# address list 
address_list = df['所有者住所'].tolist()

print(address_list)

# zip code list 
zip_list = []


# driver 
options = webdriver.ChromeOptions()
options.add_argument('--user-data-dir=C:\\Users\\kokoku\\AppData\\Local\\Google\\Chrome\\User Data')
options.add_argument('--profile-directory=Default')
# options.add_argument('--user-data-dir=C:\\Users\\yamanaka\\AppData\\Local\\Google\\Chrome\\User Data')
# options.add_argument('--profile-directory=Profile 7')
# options.add_argument("start-maximized")
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

        print(str(zipCode.text))

        # japan zip code regex
        pattern = re.compile(r"〒[0-9]{3}-[0-9]{4}")

        match = pattern.match(zipCode.text)

        if match:
            print(match.group())
            zip_list.append(match.group())

            with open('zip_list.txt', 'a', encoding="utf-8") as f:
                f.write("\n" + match.group())

        else:
            print('blank')
            zip_list.append('blank')

            with open('zip_list.txt', 'a', encoding="utf-8") as f:
                f.write("\n" + "blank")

    except TimeoutException:
        print('blank')
        zip_list.append('blank')

        with open('zip_list.txt', 'a', encoding="utf-8") as f:
            f.write("\n" + "blank")

        
df['zip'] = zip_list
df.to_excel('zip.xlsx', index=False)
print('Completed')

driver.close()