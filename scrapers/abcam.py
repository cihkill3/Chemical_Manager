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
            time.sleep(5)
            
            current_url = self.context.current_url
            logger.info(f"[Abcam] Redirected to: {current_url}")
            
            soup = BeautifulSoup(self.context.page_source, "html.parser")

            
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
                        
            # Store info
            data = {
                "제품번호": product_number.upper(),
                "시약명": title_text,
                "CAS Number": cas,
                "보관온도": "refrigerator (4 도)", # Default for most antibodies if not found
                "위험분류": "정보 없음",
                "민감성": "정보 없음",
                "상세정보_링크": current_url,
                "SDS_Link": "수동 다운로드 필요 (보안)", # Since it's protected by Cloudflare/AWS
                "SDS_Local_Path": "정보 없음"
            }
            
            # Extract storage temperature by parsing the raw text
            data["보관온도"] = "refrigerator (4 도)" # Default fallback
            
            text_lower = soup.text.lower()
            import re
            
            # Check for the modern Abcam layout format first
            modern_match = re.search(r'long-term storage conditions[\s\n]*([+\-0-9]{1,3})\s*°?\s*[cf]', text_lower)
            if modern_match:
                best_temp = modern_match.group(1)
            else:
                # Check for older "store at -20 long term" format
                old_match = re.search(r'store at\s*([+\-0-9]{1,3})\s*°?\s*[cf](?:(?!store at).){0,50}long\s*term', text_lower)
                if old_match:
                    best_temp = old_match.group(1)
                else:
                    # Fallback to the last mentioned "store at" temperature
                    all_temps = re.findall(r'store at\s*([+\-0-9]{1,3})\s*°?\s*[cf]', text_lower)
                    best_temp = all_temps[-1] if all_temps else None
                    
            if best_temp:
                val = best_temp.replace('+', '').strip()
                if val == '4':
                    data["보관온도"] = "refrigerator (4 도)"
                elif val == '-20':
                    data["보관온도"] = "freezer (-20도)"
                elif val == '-80':
                    data["보관온도"] = "deep freezer (-80도)"
                else:
                    data["보관온도"] = f"{val}도"
                
            return data
            
        except Exception as e:
            logger.error(f"[Abcam] Scraping error: {e}")
            return None
