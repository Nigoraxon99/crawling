import csv
import time
import random
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options

# Set user-agent
user_agent = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/97.0.4692.99 Safari/537.36'

# Configure Chrome webdriver options with user-agent
chrome_options = Options()
chrome_options.add_argument(f'user-agent={user_agent}')


# Function to extract product information
def extract_product_info(url, driver):
    driver.get(url)
    WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.CLASS_NAME, 'prod-buy-header__title')))
    WebDriverWait(driver, 25).until(
        EC.presence_of_all_elements_located((By.CSS_SELECTOR, '#breadcrumb.new-breadcrumbs-style .breadcrumb-link')))
    WebDriverWait(driver, 25).until(
        EC.presence_of_all_elements_located((By.CSS_SELECTOR, '.subType-IMAGE img[src^="//thumbnail"]')))

    product_name = driver.find_element(By.CSS_SELECTOR, '.prod-buy-header__title').text.strip()
    brand_name = driver.find_element(By.CSS_SELECTOR, '.prod-brand-name.brandshop-link').text.strip()
    product_price = driver.find_element(By.CSS_SELECTOR, '.total-price').text.strip()

    category_elements = driver.find_elements(By.CSS_SELECTOR, '#breadcrumb.new-breadcrumbs-style .breadcrumb-link')
    categories = [category.text.strip() for category in category_elements[1:]]

    thumbnail_urls = [img.get_attribute('src') for img in driver.find_elements(By.CSS_SELECTOR, '.lazy-load-img')]

    product_description_elements = driver.find_elements(By.CSS_SELECTOR, '.prod-description-attribute .prod-attr-item')
    product_description = '\n'.join([item.text.strip() for item in product_description_elements])

    images = [img.get_attribute('src') for img in
              driver.find_elements(By.CSS_SELECTOR, '.subType-IMAGE img[src^="//thumbnail"]')]

    return {
        'Product Name': product_name,
        'Product Brand Name': brand_name,
        'Product Price': product_price,
        'Category': '|'.join(categories),
        'Thumbnail URLs': '|'.join(thumbnail_urls),
        'Product Description': product_description,
        'Detailed Image URLs': '|'.join(images)
    }


# Main function
def main():
    # Read URLs from the Excel file
    excel_file = 'test.xlsx'
    df = pd.read_excel(excel_file, header=None)
    urls = df[0].tolist()

    # Initialize Chrome webdriver
    driver = webdriver.Chrome(options=chrome_options)

    # Create a CSV file to store the extracted data
    with open('28.csv', 'w', newline='', encoding='utf-8') as csvfile:
        fieldnames = ['Product Name', 'Product Brand Name', 'Product Price', 'Category', 'Thumbnail URLs',
                      'Product Description', 'Detailed Image URLs']
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        writer.writeheader()

        for url in urls:
            try:
                product_info = extract_product_info(url, driver)
                writer.writerow(product_info)
                print(f"Information extracted for: {product_info['Product Name']}")
            except Exception as e:
                print(f"Error extracting information for {url}: {str(e)}")

            # Add random wait time between 2 to 5 seconds
            wait_time = random.randint(2, 5)
            print(f"Waiting for {wait_time} seconds before next request...")
            time.sleep(wait_time)

    # Quit WebDriver
    driver.quit()


if __name__ == "__main__":
    main()
