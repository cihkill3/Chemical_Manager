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
    d.execute_script("arguments[0].click();", sds_elem)
    time.sleep(10)
    with open('tf_after_click.html', 'w', encoding='utf-8') as f:
        f.write(d.page_source)
except Exception as e:
    print("Error:", e)
d.quit()
