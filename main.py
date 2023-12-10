from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import ElementClickInterceptedException, NoSuchElementException, \
    ElementNotInteractableException
from selenium.webdriver import ActionChains
from selenium.webdriver.support.select import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from fill_forms import FillForms
import time


driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()))
driver.get("https://www.rightmove.co.uk/")
driver.maximize_window()
driver.implicitly_wait(10)


# Deal with cookies
try:
    accept_cookies_button = driver.find_element(By.ID, "onetrust-accept-btn-handler")
    accept_cookies_button.click()
except NoSuchElementException as e:
    print("Cookies options not present")
except ElementNotInteractableException as e:
    print("Cookies options not present")

actions = ActionChains(driver)

# Select "Chesterfield, Derbyshire"
search_input = driver.find_element(By.NAME, "typeAheadInputField")
search_input.clear()
search_input.send_keys("Chesterfield, Derbyshire")
driver.find_element(By.XPATH, "//button[normalize-space()='For Sale']").click()

# Select radius 10 miles
drp_radius_el = driver.find_element(By.ID, "radius")
drp_radius = Select(drp_radius_el)
drp_radius.select_by_value("10.0")

# Select minimum 3 bedrooms
drp_min_bedroom_el = driver.find_element(By.ID, "minBedrooms")
drp_min_bedroom = Select(drp_min_bedroom_el)
drp_min_bedroom.select_by_value("3")

# Select Property Type: Houses
drp_property_type_el = driver.find_element(By.ID, "displayPropertyType")
drp_property_type = Select(drp_property_type_el)
drp_property_type.select_by_value("houses")

# Select Added to site in last 24h
drp_property_added_to_site_el = driver.find_element(By.ID, "maxDaysSinceAdded")
drp_property_added_to_site = Select(drp_property_added_to_site_el)
drp_property_added_to_site.select_by_value("1")

# Submit
driver.find_element(By.ID, "submit").click()

# Get the number of result pages
number_of_pages = int(driver.find_element(By.XPATH, "//span[@data-bind='text: total']").text)

# Scrape required data
all_locations = []
all_prices = []
all_links = []

for n in range(number_of_pages):

    # Find and append locations
    locations = driver.find_elements(By.XPATH, "//meta[@itemprop='streetAddress']")
    for location in locations:
        get_location = location.get_attribute("content")
        strip_get_location = " ".join(get_location.splitlines())    # Remove a potential 'Enter'(new line) from the location
        all_locations.append(strip_get_location)

    # Find and append prices
    prices = driver.find_elements(By.CLASS_NAME, "propertyCard-priceValue")
    for price in prices:
        price_text = price.text
        all_prices.append(price_text)

    # Find and append links
    links = driver.find_elements(By.XPATH,
                                 "//div[@class='l-searchResult is-list']//descendant::a[@class='propertyCard-link']")
    for link in links:
        link_href = link.get_attribute("href")
        all_links.append(link_href)

    # Deal with stale element exception after moving to the next page
    try:
        def find(driver):
            next_btn = driver.find_element(By.XPATH, "//button[@title='Next page']")
            if next_btn:
                return next_btn
            else:
                return False
        next_button = WebDriverWait(driver, 5).until(find)
        next_button.click()
    except ElementClickInterceptedException:
        break


# Remove duplicate properties
non_dup_links = []
n = 0
for link in all_links:
    if link not in non_dup_links:
        non_dup_links.append(link)
    else:
        all_locations.pop(n)
        all_prices.pop(n)
        all_links.pop(n)
    n += 1

# print(all_locations)
# print(all_prices)
# print(all_links)

fill_form = FillForms(all_locations, all_prices, all_links)

# Save all_locations, all_prices, all_link to Excel spreadsheet
fill_form.fill_excel()

# Save all_locations, all_prices, all_link to Google Drive Excel
fill_form.fill_google_spreadsheet()

# Save all_locations, all_prices, all_link to csv file
fill_form.fill_csv()


time.sleep(4)
