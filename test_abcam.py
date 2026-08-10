from seleniumbase import Driver
import time
from bs4 import BeautifulSoup
import re

def scrape_abcam_sds():
    driver = Driver(uc=True, headless=True)
    try:
        url = "https://www.abcam.com/ko/products/primary-antibodies/mmp9-antibody-ep1254-ab76003"
        print("Visiting:", url)
        driver.get(url)
        time.sleep(5)  # Wait for JS to load
        
        html = driver.page_source
        soup = BeautifulSoup(html, 'html.parser')
        
        # 1. Search for links containing 'SDS' in text
        links = soup.find_all('a')
        sds_links = []
        for a in links:
            text = a.get_text(strip=True).upper()
            href = a.get('href', '')
            if 'SDS' in text or 'SAFETY DATA SHEET' in text:
                print(f"Found SDS text link: {text} -> {href}")
                if 'amazonaws.com' in href or 'sds' in href.lower() or 'pdf' in href.lower():
                    sds_links.append(href)
        
        # 2. Also search for any aws s3 link matching SDS pattern
        if not sds_links:
            for a in links:
                href = a.get('href', '')
                if 'wercs-api-prod-bucket' in href or 'SDS' in href:
                    print(f"Found SDS href pattern: {href}")
                    sds_links.append(href)
                    
        return sds_links

    except Exception as e:
        print("Error:", e)
    finally:
        driver.quit()

if __name__ == '__main__':
    res = scrape_abcam_sds()
    print("Final result:", res)
