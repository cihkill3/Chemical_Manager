from seleniumbase import Driver
import time
import os

def fetch_html():
    driver = Driver(uc=True, headless=True)
    
    urls = {
        "thermo": "https://chemicals.thermofisher.kr/apac/product/L09319",
        "tci": "https://www.tcichemicals.com/KR/en/p/C0119",
        "aldrich": "https://www.sigmaaldrich.com/KR/en/product/aldrich/803200"
    }
    
    os.makedirs("html_dumps", exist_ok=True)
    
    for name, url in urls.items():
        print(f"Fetching {name}...")
        try:
            driver.get(url)
            time.sleep(5) # wait for JS
            
            with open(f"html_dumps/{name}.html", "w", encoding="utf-8") as f:
                f.write(driver.page_source)
            print(f"Saved {name}")
        except Exception as e:
            print(f"Error {name}: {e}")
            
    driver.quit()

if __name__ == "__main__":
    fetch_html()
