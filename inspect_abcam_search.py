from seleniumbase import Driver
import time
from bs4 import BeautifulSoup

d = Driver(uc=True, headless=True)
try:
    d.get('https://www.abcam.com/ko/search?keywords=ab76003')
    time.sleep(10)
    html = d.page_source
    soup = BeautifulSoup(html, 'html.parser')
    
    # print all product links
    for a in soup.find_all('a'):
        href = a.get('href', '')
        if 'ab76003' in href:
            print("Found:", href)
except Exception as e:
    print("Err:", e)
finally:
    d.quit()
