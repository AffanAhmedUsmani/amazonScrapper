from selenium import webdriver
import pandas as pd
import concurrent.futures
import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
import time
from bs4 import BeautifulSoup
import random

def init_driver():
    options = uc.ChromeOptions()
    options.add_argument("--disable-gpu")
    driver = uc.Chrome(options=options)
    driver.get("https://www.amazon.com/")
    return driver

def solve_captcha(drivers):
    for i, driver in enumerate(drivers, start=1):
        input(f"Solve CAPTCHA for window {i} and press Enter here when done...")

def extract_data(driver, url):
    try:
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="twotabsearchtextbox"]'))
        ).click()

        search_box = driver.find_element(By.XPATH, '//*[@id="twotabsearchtextbox"]')
        search_box.send_keys(Keys.CONTROL + "a")
        search_box.send_keys(Keys.BACKSPACE)
        time.sleep(random.randint(1, 5))
        search_box.send_keys(url + Keys.ENTER)

        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, '//div[contains(@class, "s-main-slot s-result-list s-search-results sg-row")]'))
        )

        html_content = driver.page_source
        soup = BeautifulSoup(html_content, "html.parser")

        title_element = soup.find('div', attrs={"data-cy": "title-recipe"})
        if title_element and url in title_element.text:
            price_container = soup.find('span', class_='a-price')
            if price_container:
                price_offscreen = price_container.find('span', class_='a-offscreen').text
                return (url, price_offscreen)
            else:
                return (url, "Price not found")
        else:
            return (url, "Part number not found in title")
    except Exception as e:
        return (url, f"Error: {e}")

# Initialize 5 Chrome instances
drivers = [init_driver() for _ in range(5)]

# Solve CAPTCHA on each
solve_captcha(drivers)

# Load SKUs from Excel
df = pd.read_excel("Affan.xlsx")
SKUs = df['Variant SKU'].tolist()

# Ensure the task list does not exceed the number of drivers
tasks = SKUs[:len(drivers)]

# Use ThreadPoolExecutor to process tasks concurrently
results = []
with concurrent.futures.ThreadPoolExecutor(max_workers=len(drivers)) as executor:
    future_to_sku = {executor.submit(extract_data, drivers[i], sku): sku for i, sku in enumerate(tasks)}
    for future in concurrent.futures.as_completed(future_to_sku):
        results.append(future.result())

# Close drivers after completion
for driver in drivers:
    driver.quit()

# Process and export results
results_df = pd.DataFrame(results, columns=['Variant SKU', 'Scraped Price'])
results_df.to_excel("final_captcha_solved.xlsx", index=False)
print("Data exported to final_captcha_solved.xlsx successfully.")
