from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
import pandas as pd
import time
import os

download_dir = os.path.join(os.path.expanduser("~"), "Downloads")

chrome_options = Options()
prefs = {
    "download.default_directory": download_dir,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True
}
chrome_options.add_experimental_option("prefs", prefs)


PATH = r"C:\Windows\chromedriver.exe" 
service = Service(executable_path=PATH)
driver = webdriver.Chrome(service=service, options=chrome_options)
driver.get("https://www.amazon.in")
driver.maximize_window()


search_box = driver.find_element(By.ID, "twotabsearchtextbox")
search_box.clear()
search_box.send_keys("dell laptops")
driver.find_element(By.ID, "nav-search-submit-button").click()
time.sleep(2)


try:
    dell_filter = driver.find_element(By.XPATH, "//span[text()='Dell']")
    dell_filter.click()
    time.sleep(2)
except:
    print("Dell brand filter not found.")


laptops = driver.find_elements(By.XPATH, '//div[@data-component-type="s-search-result"]')
laptop_name = []
laptop_price = []
no_reviews = []

for laptop in laptops:
    
    try:
        name = laptop.find_element(By.XPATH, ".//h2/a/span").text
    except:
        name = "N/A"
    laptop_name.append(name)

    try:
        price = laptop.find_element(By.XPATH, ".//span[@class='a-price-whole']").text
    except:
        price = "0"
    laptop_price.append(price)
 
    try:
        reviews = laptop.find_element(By.XPATH, ".//span[@class='a-size-base s-underline-text']").text
    except:
        reviews = "0"
    no_reviews.append(reviews)


df = pd.DataFrame(zip(laptop_name, laptop_price, no_reviews),
                  columns=["laptop_name", "laptop_price", "no_reviews"])

excel_path = os.path.join(download_dir, "dell_laptops.xlsx")
df.to_excel(excel_path, index=False)

print(f"Excel file saved to: {excel_path}")
print(f"Total laptops scraped: {len(df)}")
driver.quit()
