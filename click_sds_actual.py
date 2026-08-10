from seleniumbase import Driver
from selenium.webdriver.common.by import By
import time
d = Driver(uc=True, headless=True)
d.get('https://www.thermofisher.com/proteins/product/Human-Apo-SAA-Recombinant-Protein/300-13-50UG')
time.sleep(5)
for _ in range(5):
    d.execute_script("window.scrollBy(0, 500);")
    time.sleep(1)

try:
    sds_elem = d.find_element(By.CSS_SELECTOR, ".pdp-documents__asset--sds a")
    print("Found element:", sds_elem.text, sds_elem.get_attribute('href'))
    
    # Click it!
    d.execute_script("arguments[0].click();", sds_elem)
    time.sleep(5)
    print("Windows:", d.window_handles)
    for handle in d.window_handles:
        d.switch_to.window(handle)
        print("Window URL:", d.current_url)
except Exception as e:
    print("Error:", e)
d.quit()
