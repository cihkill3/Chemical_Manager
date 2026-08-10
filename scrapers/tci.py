import os
import requests
import re
from bs4 import BeautifulSoup
from scrapers.base_scraper import BaseScraper
from db_manager import DBManager

class TciScraper(BaseScraper):
    def scrape(self, product_number):
        url = f"https://www.tcichemicals.com/KR/en/p/{product_number}"
        
        result = {
            "제조사": "TCI",
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
            self.context.get(url)
            self.context.sleep(3) # Cloudflare 대기
            
            page_title = self.context.get_title()
            if "not found" in page_title.lower() or "error" in page_title.lower() or "access denied" in page_title.lower():
                return result

            try:
                # TCI 타이틀 형식: 1,1'-Carbonyldiimidazole | 530-62-1 | Tokyo Chemical Industry
                parts = page_title.split('|')
                if len(parts) >= 2:
                    result["시약명"] = parts[0].strip()
                    result["CAS Number"] = parts[1].strip()
                    result["상세정보_링크"] = url
                else:
                    name = self.context.get_text("h1").strip()
                    if name: 
                        result["시약명"] = name
                        result["상세정보_링크"] = url
            except:
                pass
                
            # Properties parsing
            try:
                # ----------------------------------------------------
                # Extract hazard and storage info using BeautifulSoup
                # ----------------------------------------------------
                soup = BeautifulSoup(self.context.page_source, 'html.parser')
                
                hazard_statements = ""
                signal_word = ""
                sensitivities = []
                
                for el in soup.find_all(['tr', 'li', 'div', 'p']):
                    row_text_raw = el.get_text(separator=" ", strip=True)
                    text = row_text_raw.lower()
                    
                    if not text:
                        continue
                        
                    if "sensitive" in text or "hygroscopic" in text or "sensitive to" in text:
                        if "light sensitive" in text or "빛에 민감함" in text: sensitivities.append("Light sensitive")
                        if "moisture sensitive" in text or "수분에 민감" in text: sensitivities.append("Moisture sensitive")
                        if "air sensitive" in text or "공기에 민감" in text: sensitivities.append("Air sensitive")
                        if "heat sensitive" in text or "열에 민감" in text: sensitivities.append("Heat sensitive")
                        if "hygroscopic" in text or "흡습성" in text: sensitivities.append("Hygroscopic")
                        
                    if "storage condition" in text or "storage temperature" in text or "보관 온도" in text:
                        if "RT" not in str(result["보관온도"]):
                            result["보관온도"] = DBManager.normalize_temperature(row_text_raw)
                    elif "cas rn" in text or "cas number" in text or "cas 번호" in text:
                        cas_candidate = text.replace("cas rn", "").replace("cas number", "").replace("cas 번호", "").strip()
                        if cas_candidate and result["CAS Number"] == "정보 없음":
                            result["CAS Number"] = cas_candidate
                            
                    if "signal word" in text or "신호어" in text:
                        signal_word = re.sub(r'(?i)signal\s*word|신호어', '', row_text_raw).strip()
                    elif "hazard statements" in text or "hazard statement" in text or "유해·위험 문구" in text or "유해성 분류" in text:
                        hazard_statements = re.sub(r'(?i)hazard\s*statements?|유해·위험 문구|유해성 분류', '', row_text_raw).strip()
                        
                if sensitivities:
                    result["민감성"] = ", ".join(list(dict.fromkeys(sensitivities)))
                        
                if hazard_statements or signal_word:
                    combined = ""
                    if signal_word:
                        combined += f"[{signal_word}] "
                    if hazard_statements:
                        combined += hazard_statements
                    result["위험분류"] = combined.strip()
            except Exception as e:
                print(f"  TCI properties parsing error: {e}")

            # SDS 추출
            try:
                links = self.context.find_elements("a")
                sds_url = None
                for a in links:
                    txt = a.text.upper()
                    href = a.get_attribute("href")
                    if "SDS" in txt or "SAFETY DATA SHEET" in txt or (href and "sds" in href.lower()):
                        sds_url = href
                        break
                
                if sds_url:
                    if sds_url.startswith("/"):
                        sds_url = "https://www.tcichemicals.com" + sds_url
                    result["SDS_Link"] = sds_url
                    
                    filename = DBManager.clean_filename(result["시약명"])
                    if filename == "unknown":
                        filename = f"TCI_{product_number}"
                        
                    sds_path = os.path.join(os.getcwd(), "sds", f"{filename}.pdf")
                    os.makedirs(os.path.join(os.getcwd(), "sds"), exist_ok=True)
                    
                    cookies = {c['name']: c['value'] for c in self.context.get_cookies()}
                    
                    res = requests.get(sds_url, cookies=cookies, timeout=15)
                    if res.status_code == 200:
                        with open(sds_path, 'wb') as f:
                            f.write(res.content)
                        result["SDS_Local_Path"] = sds_path
            except Exception as e:
                print(f"  SDS Download Error: {e}")

        except Exception as e:
            print(f"  TCI scraping error: {e}")
            
        return result
