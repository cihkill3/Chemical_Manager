import os
import requests
from scrapers.base_scraper import BaseScraper
from db_manager import DBManager

class ThermofisherScraper(BaseScraper):
    def scrape(self, product_number):
        url = f"https://chemicals.thermofisher.kr/apac/product/{product_number}"
        
        result = {
            "제조사": "Thermo Fisher",
            "제품번호": product_number,
            "시약명": "제조사 홈페이지에서 검색 실패",
            "CAS Number": "정보 없음",
            "보관온도": "정보 없음",
            "위험분류": "정보 없음",
            "민감성": "정보 없음",
            "상세정보_링크": "정보 없음",
            "SDS_Link": "정보 없음",
            "SDS_Local_Path": "정보 없음"
        }
        
        try:
            suffixes = ["", ".MD", ".06", ".14", ".36", ".18", ".22", ".03"]
            found = False
            for suffix in suffixes:
                current_url = f"https://chemicals.thermofisher.kr/apac/product/{product_number}{suffix}"
                try:
                    self.context.set_page_load_timeout(10)
                    self.context.get(current_url)
                    self.context.sleep(2)
                except Exception as e:
                    print(f"    Timeout or error on {current_url}: {e}")
                    continue
                
                page_title = self.context.get_title().lower()
                if "not found" in page_title or "error" in page_title:
                    continue
                    
                try:
                    name = self.context.get_text("h1").strip()
                    if name:
                        result["시약명"] = name
                        result["상세정보_링크"] = current_url
                        found = True
                        break
                except:
                    pass
            
            if not found:
                # Fallback to global thermofisher.com
                global_url = f"https://www.thermofisher.com/order/catalog/product/{product_number}"
                try:
                    self.context.set_page_load_timeout(15)
                    self.context.get(global_url)
                    self.context.sleep(3)
                except Exception as e:
                    print(f"    Timeout or error on global {global_url}: {e}")
                
                try:
                    name = self.context.get_text("h1").strip()
                    if name:
                        result["시약명"] = name
                        result["상세정보_링크"] = global_url
                        found = True
                except:
                    pass
            
            if not found:
                return result
                
            # Properties parsing
            try:
                html = self.context.page_source
                from bs4 import BeautifulSoup
                import re
                soup = BeautifulSoup(html, 'html.parser')
                
                sensitivities = []
                hazard_statements = []
                
                for element in soup.find_all(['tr', 'p', 'div', 'li', 'span']):
                    text = element.get_text(separator=' ', strip=True).lower()
                    if not text:
                        continue
                        
                    if "cas" in text and ("number" in text or "번호" in text):
                        m = re.search(r'\d{2,7}-\d{2}-\d', text)
                        if m and result["CAS Number"] == "정보 없음":
                            result["CAS Number"] = m.group(0)
                    elif ("temperature" in text or "보관온도" in text) and "rt" not in str(result["보관온도"]).lower():
                        result["보관온도"] = DBManager.normalize_temperature(text)
                            
                    if "sensitive" in text or "hygroscopic" in text or "sensitive to" in text:
                        if "light sensitive" in text or "빛에 민감" in text: sensitivities.append("Light sensitive")
                        if "moisture sensitive" in text or "수분에 민감" in text: sensitivities.append("Moisture sensitive")
                        if "air sensitive" in text or "공기에 민감" in text: sensitivities.append("Air sensitive")
                        if "heat sensitive" in text or "열에 민감" in text: sensitivities.append("Heat sensitive")
                        if "hygroscopic" in text or "흡습성" in text: sensitivities.append("Hygroscopic")
                        
                if sensitivities:
                    result["민감성"] = ", ".join(list(dict.fromkeys(sensitivities)))
                
                # Extract H-codes using regex
                from scrapers.hcodes_dict import get_h_statement, GHS_H_CODES
                h_codes_raw = re.findall(r'\bh[234]\d{2}[a-z]*\b', html.lower())
                h_codes = []
                for code in h_codes_raw:
                    if code in GHS_H_CODES:
                        stmt = get_h_statement(code)
                        if stmt not in h_codes:
                            h_codes.append(stmt)
                
                if h_codes:
                    result["위험분류"] = " / ".join(h_codes)
            except Exception as e:
                print(f"  TF properties parsing error: {e}")

            # SDS 링크 설정
            result["SDS_Link"] = f"https://www.thermofisher.com/search/browse/results?term={product_number}"
        except Exception as e:
            print(f"  ThermoFisher scraping error: {e}")
            
        return result
