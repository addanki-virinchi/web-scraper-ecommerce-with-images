import time
import requests
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from openpyxl import Workbook
from openpyxl.drawing.image import Image as ExcelImage
from io import BytesIO
from PIL import Image as PILImage

# Setup the Selenium WebDriver
service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service)

driver.get("https://www.myer.com.au/c/women/clothing/women-clothing-tops?")
driver.implicitly_wait(5)

# Function to find the next button
def find_next_button():
    try:
        next_button = driver.find_element(By.CLASS_NAME, 'item-number')
        return next_button
    except:
        return None

# Initialize page number and lists to store data
page_number = 1
Brand_name = []
prod = []
product_price = []
product_images = []

# Loop to scrape data across all pages
while True:
    print(f"Processing page {page_number}")
    
    # Scrape brand names
    hh = driver.find_elements(By.CSS_SELECTOR, "p.css-1ps1gwj")
    Brand_name.extend([brand.text for brand in hh])
    
    # Scrape product names
    product_name = driver.find_elements(By.CSS_SELECTOR, '[data-automation="product-name"]')
    prod.extend([name.text for name in product_name])
    
    # Scrape product prices
    price_element = driver.find_elements(By.CSS_SELECTOR, '[data-automation="product-price-was"]')
    product_price.extend([price.text for price in price_element])
    
    # Scrape product images
    image_elements = driver.find_elements(By.CSS_SELECTOR, '[data-automation="product-image"]')
    product_images.extend([img.get_attribute("src") for img in image_elements])
    
    # Find and click the next button
    next_button = find_next_button()
    if next_button and page_number < 30:
        next_button.click()
        time.sleep(2)
        page_number += 1
    else:
        break

# Close the WebDriver
driver.quit()

# Check lengths of lists
print(f"Lengths -> Brand_name: {len(Brand_name)}, prod: {len(prod)}, product_price: {len(product_price)}, product_images: {len(product_images)}")

# Write the data to an Excel file
wb = Workbook()
ws = wb.active
ws.title = "Scraped Data"

# Add headers
ws.append(["Brand Name", "Product Name", "Price", "Image URL", "Image"])

# Set column width for the image column
ws.column_dimensions['E'].width = 30  # Adjust as needed

# Insert the data into the Excel file
for i in range(min(len(Brand_name), len(prod), len(product_price), len(product_images))):
    ws.append([Brand_name[i], prod[i], product_price[i], product_images[i]])
    
    # Download and insert the image
    try:
        response = requests.get(product_images[i])
        img_data = BytesIO(response.content)
        img = PILImage.open(img_data)
        
        # Resize the image (optional)
        img.thumbnail((100, 100), PILImage.ANTIALIAS)
        
        # Save the image temporarily in memory for adding to Excel
        temp_img = BytesIO()
        img.save(temp_img, format="PNG")
        temp_img.seek(0)
        
        # Insert the image into the Excel sheet
        excel_img = ExcelImage(temp_img)
        ws.add_image(excel_img, f"E{i + 2}")

        # Optionally, adjust row height
        ws.row_dimensions[i + 2].height = 80  # Adjust as needed
    except Exception as e:
        print(f"Error with image at index {i}: {e}")

# Save the workbook
wb.save("scraped_data_with_images.xlsx")

print("Data scraping completed and saved to Excel.")
