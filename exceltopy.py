import openpyxl
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

# Load the Excel file
excelfile = openpyxl.load_workbook('E-commerce URL.xlsx')
sheet = excelfile.active

with open('scraped_data.txt', 'w',encoding='utf-8') as f:
# Set up the web driver
    driver = webdriver.Chrome('C:/Drivers/chromedriver.exe')

    # Loop through each row in the sheet and scrape the URL
    for row in sheet.iter_rows(min_row=1, max_row=10, min_col=1, max_col=26, values_only=True):
        try:
            url = row[0]
            driver.get(url)
            # print(url)
            c = webdriver.ChromeOptions()
            c.add_argument("--headless")
            c.add_argument("--disable-gpu")
            c.add_argument("--disable-software-rasterizer")
            c.add_argument("--ignore-certificate-error")
            c.add_argument("--ignore-ssl-errors")
            c.add_experimental_option('excludeSwitches', ['enable-logging'])
            # driver = webdriver.Chrome(chrome_options = c, executable_path= "C:/drivers/chromedriver.exe")
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            price = (WebDriverWait(driver,5).until(EC.visibility_of_element_located((By.CLASS_NAME, 'pqTWkA'))).get_attribute("innerText"))
            print(price)
            f.write(price + '\n')
        except Exception as e:
            print("The Product Dosen't Exits ")
            f.write("The Product Dosen't Exits" +'\n')
            continue

driver.quit()