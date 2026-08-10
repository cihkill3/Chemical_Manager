from seleniumbase import Driver
import time
import re

d = Driver(uc=True, headless=True)
d.get('https://www.sigmaaldrich.com/KR/ko/sds/aldrich/803200')
time.sleep(5)
html = d.page_source

# Try to find a PDF link
pdf_links = re.findall(r'href=[\'\"]([^\'\"]+\.pdf[^\'\"]*)[\'\"]', html)
print('PDF HREFS:', pdf_links)

d.quit()
