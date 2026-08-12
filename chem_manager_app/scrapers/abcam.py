import time
import urllib.parse
from bs4 import BeautifulSoup
from .base_scraper import BaseScraper
import logging

logger = logging.getLogger(__name__)

class AbcamScraper(BaseScraper):
    def scrape(self, product_number):
        product_number = product_number.strip().lower()
        url = f"https://www.abcam.com/ko-kr/products/search?keywords={product_number}"
        short_url = f"https://www.abcam.com/{product_number}"
        logger.info(f"[Abcam] Visiting shortlink for {product_number}: {short_url}")
        
        try:
            self.context.get(short_url)
            time.sleep(5) # Reverted to explicit sleep
            
            current_url = self.context.get_current_url()
            logger.info(f"[Abcam] Redirected to: {current_url}")
            
            soup = BeautifulSoup(self.context.get_page_source(), "html.parser")

            
            title = soup.find('title')
            title_text = title.text.strip() if title else "Unknown Product"
            
            # Remove " | Abcam" from title if present
            if " | Abcam" in title_text:
                title_text = title_text.replace(" | Abcam", "").strip()
            
            # CAS Number usually not present for antibodies, but we'll try to find it just in case
            cas = "N/A"
            for td in soup.find_all('td'):
                if 'CAS' in td.text:
                    nxt = td.find_next_sibling('td')
                    if nxt:
                        cas = nxt.text.strip()
                        
            data = {
                "Manufacturer": "Abcam", "Catalog No.": product_number.upper(), "Product Name": title_text,
                "CAS No.": cas if cas != "정보 없음" else "", "Storage Temp.": "R", "Signal Word": "", "Key Hazards": "",
                "Detailed Hazard Classification": "", "Sensitivity": "", "Detail_Link": current_url, "SDS_Link": "", "SDS_Local_Path": ""
            }
            
            # Extract storage temperature by parsing the raw text
            text_lower = soup.text.lower()
            import re
            
            modern_match = re.search(r'long-term storage conditions[\s\n]*([+\-0-9]{1,3})\s*°?\s*[cf]', text_lower)
            if modern_match:
                best_temp = modern_match.group(1)
            else:
                old_match = re.search(r'store at\s*([+\-0-9]{1,3})\s*°?\s*[cf](?:(?!store at).){0,50}long\s*term', text_lower)
                if old_match:
                    best_temp = old_match.group(1)
                else:
                    all_temps = re.findall(r'store at\s*([+\-0-9]{1,3})\s*°?\s*[cf]', text_lower)
                    best_temp = all_temps[-1] if all_temps else None
                    
            if best_temp:
                val = best_temp.replace('+', '').strip()
                norm_t = "R"
                if val in ['4', '+4']:
                    norm_t = "R"
                elif val in ['-20', '-80']:
                    norm_t = "F" if val == '-20' else "DF"
                else:
                    norm_t = f"{val}°C"
                data["Storage Temp."] = norm_t
                
            return data
            
        except Exception as e:
            logger.error(f"[Abcam] Scraping error: {e}")
            return {
                "Manufacturer": "Abcam", "Catalog No.": product_number, "Product Name": "",
                "CAS No.": "Search Failed", "Storage Temp.": "-", "Signal Word": "-", "Key Hazards": "-",
                "Detailed Hazard Classification": "-", "Sensitivity": "-", "Detail_Link": "Product Not Found", "SDS_Link": "-", "SDS_Local_Path": "-"
            }
