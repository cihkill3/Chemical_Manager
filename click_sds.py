from seleniumbase import Driver
from selenium.webdriver.common.by import By
import time
d = Driver(uc=True, headless=True)
d.get('https://www.thermofisher.com/proteins/product/Human-Apo-SAA-Recombinant-Protein/300-13-50UG')
time.sleep(10)
try:
    sds_link = d.find_element(By.XPATH, "//*[contains(text(), 'Safety Data Sheets')]")
    d.execute_script("arguments[0].scrollIntoView();", sds_link)
    time.sleep(1)
    
    # Try to find the link containing SDS text next to the diamond icon
    sds_elem = d.find_element(By.XPATH, "//a[contains(., 'SDS')]")
    print("Found element:", sds_elem.text, sds_elem.get_attribute('href'))
    
    # Click it!
    d.execute_script("arguments[0].click();", sds_elem)
    time.sleep(5)
    print("Current URL:", d.current_url)
    print("Windows:", d.window_handles)
    for handle in d.window_handles:
        d.switch_to.window(handle)
        print("Window URL:", d.current_url)
except Exception as e:
    print("Error:", e)
d.quit()
