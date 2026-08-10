from seleniumbase import Driver
from selenium.webdriver.common.by import By
import time
d = Driver(uc=True, headless=True)
d.get('https://www.thermofisher.com/proteins/product/Human-Apo-SAA-Recombinant-Protein/300-13-50UG')
time.sleep(10)
els = d.find_elements(By.XPATH, "//a[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'sds') or contains(@href, 'SDS')]")
for e in els:
    try:
        print('TEXT:', e.text.strip(), 'HREF:', e.get_attribute('href'))
    except:
        pass
d.quit()
