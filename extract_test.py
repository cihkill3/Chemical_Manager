import time
from seleniumbase import Driver
import re

products = [
    ("thermo fisher", "L09319"),
    ("thermo fisher", "L16400"),
    ("TCI", "C0119"),
    ("TCI", "T0751"),
    ("aldrich", "803200"),
    ("aldrich", "909270-10G")
]

def parse_all():
    driver = Driver(uc=True, headless=True)
    
    for vendor, p in products:
        print(f"\n================ {vendor} - {p} ================")
        try:
            if vendor == "thermo fisher":
                url = f"https://www.thermofisher.com/order/catalog/product/{p}"
            elif vendor == "TCI":
                url = f"https://www.tcichemicals.com/KR/en/p/{p}"
            elif vendor == "aldrich":
                # Need just the product number, so remove the '-10G' if present
                clean_p = p.split('-')[0]
                url = f"https://www.sigmaaldrich.com/KR/en/product/aldrich/{clean_p}"
                
            driver.get(url)
            time.sleep(6)
            
            print(f"URL: {url}")
            print(f"Title: {driver.get_title()}")
            
            try:
                if vendor == "thermo fisher":
                    name = driver.get_text("h1")
                elif vendor == "TCI":
                    name = driver.get_text("h1.product-name")
                elif vendor == "aldrich":
                    name = driver.get_text("h1")
                print(f"Name: {name}")
            except Exception as e:
                print(f"Name error: {e}")
                
            html = driver.page_source
            
            # Use regex to find CAS, temp, hazard in raw html
            if "CAS" in html:
                print("CAS string found in HTML.")
            
            # Print all TRs text
            print("--- TRs ---")
            for tr in driver.find_elements("tr"):
                if "cas" in tr.text.lower() or "temperature" in tr.text.lower() or "hazard" in tr.text.lower() or "signal word" in tr.text.lower():
                    print(f"TR: {tr.text.strip().replace(chr(10), ' ')}")
                    
            # Print all elements that might have Hazard statements
            print("--- Hazard Texts ---")
            # For TCI and Aldrich they usually have specific sections
            if vendor == "TCI":
                for el in driver.find_elements(".product-hazard"):
                    print(f"Hazard: {el.text.replace(chr(10), ' ')}")
            if vendor == "aldrich":
                for el in driver.find_elements("td"):
                    if "H" in el.text and ":" in el.text:
                        print(f"TD with H: {el.text.replace(chr(10), ' ')}")

            print("--- Links ---")
            sds_links = []
            for a in driver.find_elements("a"):
                txt = a.text.lower()
                href = a.get_attribute("href")
                if "sds" in txt or "safety data" in txt or (href and "sds" in href.lower()):
                    sds_links.append((txt, href))
            for sl in set(sds_links):
                print(f"SDS Link: {sl}")

        except Exception as e:
            print(f"Error parsing: {e}")
            
    driver.quit()

if __name__ == "__main__":
    parse_all()
