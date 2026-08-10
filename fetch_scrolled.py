from seleniumbase import Driver
from selenium.webdriver.common.by import By
import time
d = Driver(uc=True, headless=True)
d.get('https://www.thermofisher.com/proteins/product/Human-Apo-SAA-Recombinant-Protein/300-13-50UG')
time.sleep(5)
for _ in range(5):
    d.execute_script("window.scrollBy(0, 500);")
    time.sleep(1)
html = d.page_source
with open('tf_protein_page_scrolled.html', 'w', encoding='utf-8') as f:
    f.write(html)
d.quit()
