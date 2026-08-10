from seleniumbase import Driver
import time
d = Driver(uc=True, headless=True)
try:
    d.get('https://www.abcam.com/ko/search?keywords=ab76003')
    time.sleep(10)
    links = d.find_elements('css selector', "a[href*='/products/']")
    for l in links:
        h = l.get_attribute('href')
        if not h: continue
        if 'ab76003' in h.lower():
            print('MATCH:', h)
        elif len(h.split('-')) > 2:
            print('POTENTIAL:', h)
except Exception as e:
    print('Err:', e)
finally:
    d.quit()
