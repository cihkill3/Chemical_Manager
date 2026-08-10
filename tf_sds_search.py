from seleniumbase import Driver
import time
from bs4 import BeautifulSoup
import re

def search_tf_sds(product_num):
    driver = Driver(uc=True, headless=True)
    try:
        url = f"https://www.thermofisher.com/search/browse/category/us/en/90226/Safety+Data+Sheets+(SDS)?query={product_num}"
        print(f"Visiting {url}")
        driver.get(url)
        time.sleep(5) # wait for angular/react to load
        html = driver.page_source
        soup = BeautifulSoup(html, 'html.parser')
        links = soup.find_all('a', href=True)
        pdf_links = []
        for a in links:
            href = a['href']
            if 'SDS' in href or 'sds' in href or 'document-connect' in href or '.pdf' in href:
                pdf_links.append(href)
                print("Found SDS related link:", href)
        return pdf_links
    finally:
        driver.quit()

if __name__ == '__main__':
    search_tf_sds('704060F')
