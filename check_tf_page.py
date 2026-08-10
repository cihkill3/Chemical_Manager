from seleniumbase import Driver
import time

driver = Driver(uc=True, headless=True)
driver.get("https://www.thermofisher.com/order/catalog/product/704060F")
time.sleep(5)
html = driver.page_source
with open("tf_704060F_page.html", "w", encoding="utf-8") as f:
    f.write(html)
driver.quit()
