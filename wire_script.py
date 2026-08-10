from seleniumwire import webdriver
import time

options = webdriver.ChromeOptions()
options.add_argument('--headless')
options.add_argument('--disable-gpu')
driver = webdriver.Chrome(options=options)

driver.get('https://www.thermofisher.com/proteins/product/Human-Apo-SAA-Recombinant-Protein/300-13-50UG')
time.sleep(10)

for request in driver.requests:
    if request.response:
        url = request.url
        if '300-13' in url or 'sds' in url.lower() or 'document' in url.lower():
            print(request.url, request.response.status_code)

driver.quit()
