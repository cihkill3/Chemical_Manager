from seleniumbase import Driver
import time
d = Driver(uc=True, headless=True)
d.get('https://www.thermofisher.com/proteins/product/Human-Apo-SAA-Recombinant-Protein/300-13-50UG')
time.sleep(10)
resources = d.execute_script('return window.performance.getEntriesByType("resource").map(e => e.name);')
for r in resources:
    if 'api' in r.lower() or 'document' in r.lower():
        print(r)
d.quit()
