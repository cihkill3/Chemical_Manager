from seleniumbase import Driver
import time
d = Driver(uc=True, headless=True)
d.get('https://www.thermofisher.com/search/browse/results?keyword=300-13&searchType=SDS&reqType=json')
time.sleep(5)
html = d.page_source
with open('tf_sds_api_2.html', 'w', encoding='utf-8') as f:
    f.write(html)
d.quit()
