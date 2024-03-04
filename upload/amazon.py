from concurrent.futures import ThreadPoolExecutor, as_completed
import pandas as pd
from selenium import webdriver
import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
import time
from bs4 import BeautifulSoup
import random
def init_driver():
    # Initialize a Chrome driver instance
    options = uc.ChromeOptions()
    # options.add_argument("--headless")  # Set to True for headless mode
    options.add_argument("--disable-gpu")
    driver = uc.Chrome(options=options)
    driver.get("https://www.amazon.ca/")
    return driver

def solve_captcha(driver, index):
    # Display message to solve CAPTCHA, specific to each driver instance
    input(f"Please solve the CAPTCHA for driver {index + 1} and press Enter here when done...")

def extract_data(driver, sku):
    try:
        # Navigate to Amazon and search for the SKU
        
        search_box = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "twotabsearchtextbox")))
        search_box.send_keys(Keys.CONTROL + "a")
        search_box.send_keys(Keys.BACKSPACE)
        time.sleep(random.randint(1, 5))
        search_box.send_keys(sku + Keys.ENTER)
        
        # Wait for the search results to load
        WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.CSS_SELECTOR, ".s-main-slot.s-result-list")))
        
        html_content = driver.page_source
        soup = BeautifulSoup(html_content, "html.parser")

        title_element = soup.find('div', attrs={"data-cy": "title-recipe"})
        if title_element and sku in title_element.text:
            price_container = soup.find('span', class_='a-price')
            if price_container:
                price_offscreen = price_container.find('span', class_='a-offscreen').text
                return (sku, price_offscreen)
            else:
                return (sku, "Price not found")
        else:
            return (sku, "Part number not found in title")
        
    except Exception as e:
        return (sku, "Error occurred")

def process_sku(driver, sku):
    # Wrapper function to use in threading
    return extract_data(driver, sku)

# Load SKUs from Excel
df = pd.read_excel("Affan.xlsx")
skus = df['Variant SKU'].tolist()

# Initialize a pool of Chrome drivers
num_drivers = 5
drivers = [init_driver() for _ in range(num_drivers)]

# Manually solve CAPTCHA for each driver
for index, driver in enumerate(drivers):
    solve_captcha(driver, index)

# Use ThreadPoolExecutor to process each SKU concurrently across the initialized drivers
results = []
with ThreadPoolExecutor(max_workers=num_drivers) as executor:
    # Create a future for each SKU
    futures = [executor.submit(process_sku, drivers[i % num_drivers], sku) for i, sku in enumerate(skus)]
    
    # Wait for all futures to complete
    for future in as_completed(futures):
        results.append(future.result())

# Close all drivers
for driver in drivers:
    driver.quit()

# Save the results to a DataFrame and then to an Excel file
results_df = pd.DataFrame(results, columns=['Variant SKU', 'Scraped Price'])
results_df.to_excel("final_results.xlsx", index=False)

print("Data exported to final_results.xlsx successfully.")
