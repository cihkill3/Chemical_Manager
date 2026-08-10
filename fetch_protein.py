from seleniumbase import Driver
import time
d = Driver(uc=True, headless=True)
d.get('https://www.thermofisher.com/proteins/product/Human-Apo-SAA-Recombinant-Protein/300-13-50UG')
time.sleep(10)
html = d.page_source
with open('tf_protein_page.html', 'w', encoding='utf-8') as f:
    f.write(html)
import re
links = re.findall(r'href=[\"\']([^\"\']*?\.pdf[^\"\']*?)[\"\']', html, flags=re.IGNORECASE)
for l in links:
    print('PDF LINK:', l)
d.quit()
