from seleniumbase import Driver
from selenium.webdriver.common.by import By
import time
d = Driver(uc=True, headless=True)
d.get('https://www.thermofisher.com/proteins/product/Human-Apo-SAA-Recombinant-Protein/300-13-50UG')
time.sleep(10)
els = d.find_elements(By.XPATH, "//*[contains(text(), 'SDS')]")
for e in els:
    print(e.tag_name, e.get_attribute('href'), e.text)

# Also check for document-connect links
links = d.find_elements(By.XPATH, "//a[contains(@href, 'documents.thermofisher.com') or contains(@href, 'document-connect')]")
for l in links:
    print('Doc link:', l.get_attribute('href'))
d.quit()
