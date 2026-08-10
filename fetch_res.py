from seleniumbase import Driver
import time
import json
d = Driver(uc=True, headless=True)
d.get('https://www.thermofisher.com/proteins/product/Human-Apo-SAA-Recombinant-Protein/300-13-50UG')
time.sleep(10)
resources = d.execute_script('return window.performance.getEntriesByType("resource").map(e => e.name);')
res = [r for r in resources if '300' in r or 'sds' in r.lower()]
with open('tf_res.json', 'w', encoding='utf-8') as f:
    json.dump(res, f, indent=2)
d.quit()
