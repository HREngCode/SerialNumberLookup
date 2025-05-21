import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Load the Excel file
input_file_path = r"H:\Reference\HP_Serial_Numbers.xlsx"
df = pd.read_excel(input_file_path)

# Set up the WebDriver
driver = webdriver.Chrome()  # Make sure the ChromeDriver is correctly installed and in PATH

# Open the website
driver.get("https://support.hp.com/us-en/check-warranty")
wait = WebDriverWait(driver, 10)

# Click the "Accept" button
accept_button = wait.until(EC.element_to_be_clickable((By.ID, "onetrust-accept-btn-handler")))
accept_button.click()

# Iterate through each serial number
for index, row in df.iterrows():
    serial_number = row['Serial Number']

    if pd.isna(serial_number) or str(serial_number).strip() == "":
        print(f"Skipping empty or invalid serial number at index {index}")
        continue

    try:
        print(f"Processing serial number: {serial_number}")

        # Locate input field and enter serial number
        input_field = wait.until(EC.presence_of_element_located((By.XPATH, "//input[@id='inputtextpfinder']")))
        input_field.clear()
        input_field.send_keys(serial_number)

        # Locate and click the submit button
        input_field.send_keys(Keys.RETURN)

        # Wait for 'Start Date' element
        date_element = wait.until(
            EC.presence_of_element_located(
                (By.XPATH, "//div[contains(@class, 'info-item')]//div[contains(@class, 'label') and contains(text(), 'Start date')]/following-sibling::div[contains(@class, 'text')]")
            )
        )
        print(f"Element found: {date_element.text}")
        start_date = date_element.text
        print(f"Start Date for {serial_number}: {start_date}")

        # Store the result in the DataFrame
        df.at[index, 'Start Date'] = start_date

    except Exception as e:
        print(f"Error processing serial number {serial_number}: {e}")
        df.at[index, 'Start Date'] = "Error"

    finally:
        # Return to the initial page
        driver.get("https://support.hp.com/us-en/check-warranty")
        time.sleep(2)  # Short delay to ensure the page resets

# Save the updated DataFrame to a new Excel file
output_file_path = r"H:\Reference\SerialNumber\hp_with_start_dates.xlsx"
df.to_excel(output_file_path, index=False)

# Close the WebDriver
driver.quit()

print(f"Process completed. Results saved to {output_file_path}.")