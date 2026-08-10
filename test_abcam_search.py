from seleniumbase import Driver
import time

d = Driver(uc=True, headless=True)
try:
    d.get('https://www.abcam.com/ko/search?keywords=ab76003')
    time.sleep(5)
    print('Current URL:', d.current_url)
    print('Page Title:', d.title)
    
    links = d.find_elements("css selector", "a[href*='/products/']")
    print(f"Found {len(links)} links")
    for l in links[:5]:
        print(l.get_attribute('href'))
except Exception as e:
    print('Err:', e)
finally:
    d.quit()
