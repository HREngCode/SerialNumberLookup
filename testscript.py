from selenium import webdriver

# For Chrome
driver = webdriver.Chrome()  # Ensure 'chromedriver' is in PATH

# For Firefox (if you're using Firefox)
# driver = webdriver.Firefox()  # Ensure 'geckodriver' is in PATH

# Open a webpage
driver.get("https://www.google.com")

# Print the title of the webpage
print(driver.title)

# Close the browser
driver.quit()
