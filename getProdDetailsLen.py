# import requests
# from bs4 import BeautifulSoup
# import pandas as pd

# # Function to look up product details based on serial number
# def get_product_details(serial_number):
#     url = f"https://pcsupport.lenovo.com/us/en/warrantylookup?serialNumber={serial_number}"
    
#     # Send a GET request to the Lenovo warranty page
#     response = requests.get(url)
    
#     if response.status_code == 200:
#         # Parse the page content
#         soup = BeautifulSoup(response.content, 'html.parser')
        
#         # Look for the manufacturing date or other relevant information
#         # You will need to inspect the HTML structure to find the correct tag
#         # For illustration, let's assume manufacturing date is in a tag with id 'manufacture-date'
#         manufacture_date = soup.find(id="manufacture-date")
        
#         if manufacture_date:
#             return manufacture_date.text.strip()  # Extract text and remove extra spaces
#         else:
#             return "Manufacture date not found"
#     else:
#         return "Error: Unable to fetch details"

# # Read serial numbers from Excel file (assuming serial numbers are in the first column)
# df = pd.read_excel('H:\Reference\Serial_Numbers.xlsx')

# # Create a new column to store manufacturing dates
# df['Manufacturing Date'] = df['Serial Number'].apply(get_product_details)

# # Save the results back to a new Excel file
# df.to_excel('output_with_manufacture_dates.xlsx', index=False)

# print("Manufacturing dates have been added to the Excel file.")


import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Load the Excel file
input_file_path = r"H:\Reference\Serial_Numbers.xlsx"
df = pd.read_excel(input_file_path)

# Set up the WebDriver
driver = webdriver.Chrome()  # Make sure the ChromeDriver is correctly installed and in PATH

# Open the website
driver.get("https://pcsupport.lenovo.com/us/en/warranty-lookup#/")
wait = WebDriverWait(driver, 3)

# Iterate through each serial number
for index, row in df.iterrows():
    serial_number = row['Serial Number']

    if pd.isna(serial_number) or str(serial_number).strip() == "":
        print(f"Skipping empty or invalid serial number at index {index}")
        continue

    try:
        print(f"Processing serial number: {serial_number}")

        # Locate input field and enter serial number
        input_field = wait.until(EC.presence_of_element_located((By.CLASS_NAME, "button-placeholder__input")))
        input_field.clear()
        input_field.send_keys(serial_number)

        # Locate and click the submit button
        input_field.send_keys(Keys.RETURN)

        # Wait for 'Start Date' element
        date_element = wait.until(
            EC.presence_of_element_located(
                (By.XPATH, "//div[@class='detail-property'][.//span[@class='property-title' and text()='Start Date']]//span[@class='property-value']")
            )
        )
        start_date = date_element.text
        print(f"Start Date for {serial_number}: {start_date}")

        # Store the result in the DataFrame
        df.at[index, 'Start Date'] = start_date

    except Exception as e:
        print(f"Error processing serial number {serial_number}: {e}")
        df.at[index, 'Start Date'] = "Error"

    finally:
        # Return to the initial page
        driver.get("https://pcsupport.lenovo.com/us/en/warranty-lookup#/")
        time.sleep(2)  # Short delay to ensure the page resets

# Save the updated DataFrame to a new Excel file
output_file_path = r"H:\Reference\SerialNumber\output_with_start_dates.xlsx"
df.to_excel(output_file_path, index=False)

# Close the WebDriver
driver.quit()

print(f"Process completed. Results saved to {output_file_path}.")