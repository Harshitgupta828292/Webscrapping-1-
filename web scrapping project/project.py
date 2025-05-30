
import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys

# Initialize lists to store laptop data
laptop_names = []
laptop_prices = []
laptop_reviews = []

# Start Selenium WebDriver
driver = webdriver.Chrome()
driver.get("https://www.amazon.in")
driver.maximize_window()

# Search for Dell laptops
search_box = driver.find_element(By.ID, "twotabsearchtextbox")
search_box.clear()
search_box.send_keys("dell laptop")
driver.find_element(By.ID, "nav-search-submit-button").click()

# Apply Dell brand filter
time.sleep(3)  # Wait for filters to load
try:
    driver.find_element(By.XPATH, "//span[text()='Dell']").click()
    time.sleep(4)  # Wait for page to update
except:
    print("Dell filter not found, proceeding without it.")

# Scrape all product listings
all_products = driver.find_elements(By.XPATH, "//div[@data-component-type='s-search-result']")

for product in all_products:
    # Extract laptop name
    try:
        name = product.find_element(By.XPATH, ".//span[contains(text(),'Dell')]")
        laptop_names.append(name.text)
    except:
        laptop_names.append("N/A")

    # Extract laptop price
    try:
        price = product.find_element(By.XPATH, ".//span[@class='a-price-whole']")
        laptop_prices.append(price.text)
    except:
        laptop_prices.append("N/A")

    # Extract laptop review count
    try:
        review = product.find_element(By.XPATH, ".//span[@class='a-size-base s-underline-text']")
        laptop_reviews.append(review.text)
    except:
        laptop_reviews.append("0")

# Print collected data counts
print(f"Total Laptops Found: {len(laptop_names)}")
print(f"Total Prices Found: {len(laptop_prices)}")
print(f"Total Reviews Found: {len(laptop_reviews)}")

# in these i use the python library called pandas to save the file to the excel sheet 
df=pd.DataFrame(zip(laptop_names,laptop_prices,laptop_reviews),columns=['laptop_names','laptop_prices','laptop_reviews'])
save_path=(r"C:\Users\Harshitramji11\Downloads\Learnerea\live_laptop.xlsx")
df.to_excel(save_path, index=False, engine='openpyxl')
print(f"Excel file saved successfully at: {save_path}")
driver.quit()