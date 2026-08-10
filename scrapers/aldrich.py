import os
import requests
import re
from scrapers.base_scraper import BaseScraper
from db_manager import DBManager

class AldrichScraper(BaseScraper):
    def scrape(self, product_number):
        # Aldrich 제품 번호는 뒤에 -10G 등 패키지 크기가 붙는 경우가 많으므로 제거합니다.
        clean_product_number = product_number.split('-')[0]
        url = f"https://www.sigmaaldrich.com/KR/en/product/aldrich/{clean_product_number}"
        
        result = {
            "제조사": "Sigma-Aldrich",
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
            self.context.sleep(3) # Cloudflare 대기 및 렌더링
            
            page_title = self.context.get_title()
            if "not found" in page_title.lower() or "error" in page_title.lower():
                return result

            try:
                # Aldrich는 h1 태그나 title 등에서 시약명을 가져옵니다.
                name = self.context.get_text("h1").strip()
                if name:
                    result["시약명"] = name
                    result["상세정보_링크"] = url
                elif "=" in page_title:
                    result["시약명"] = page_title.split("=")[0].strip()
                    result["상세정보_링크"] = url
            except:
                pass
                
            # HTML 전문을 BeautifulSoup으로 파싱하여 프로퍼티 추출 (div/span 등 구조가 불규칙함)
            try:
                html = self.context.page_source
                from bs4 import BeautifulSoup
                soup = BeautifulSoup(html, 'html.parser')
                
                # CAS Number
                for el in soup.find_all(['div', 'span']):
                    text = el.get_text(separator=' ', strip=True).lower()
                    if "cas 번호" in text or "cas number" in text:
                        import re
                        m = re.search(r'\d{2,7}-\d{2}-\d', text)
                        if m:
                            result["CAS Number"] = m.group(0)
                            break
                            
                # 보관온도
                for el in soup.find_all(['div', 'span', 'tr']):
                    text = el.get_text(separator=' ', strip=True).lower()
                    if "storage temp" in text or "보관온도" in text:
                        result["보관온도"] = DBManager.normalize_temperature(text.replace("storage temp.", "").replace("storage temp", "").strip())
                        break
                        
                result["SDS_Link"] = f"https://www.sigmaaldrich.com/KR/ko/sds/aldrich/{product_number}"

                # 민감성 (Sensitivity)
                sensitivities = []
                for el in soup.find_all(['div', 'span', 'tr', 'p', 'li']):
                    text = el.get_text(separator=' ', strip=True).lower()
                    if "sensitive" in text or "hygroscopic" in text or "sensitive to" in text:
                        if "light sensitive" in text: sensitivities.append("Light sensitive")
                        if "moisture sensitive" in text: sensitivities.append("Moisture sensitive")
                        if "air sensitive" in text: sensitivities.append("Air sensitive")
                        if "heat sensitive" in text: sensitivities.append("Heat sensitive")
                        if "hygroscopic" in text: sensitivities.append("Hygroscopic")
                if sensitivities:
                    result["민감성"] = ", ".join(list(dict.fromkeys(sensitivities)))

                # 위험 분류 (GHS Hazard) - H-codes
                from scrapers.hcodes_dict import get_h_statement, GHS_H_CODES
                h_codes_raw = re.findall(r'\bh[234]\d{2}[a-z]*\b', html.lower())
                h_codes = []
                for code in h_codes_raw:
                    if code in GHS_H_CODES:
                        stmt = get_h_statement(code)
                        if stmt not in h_codes:
                            h_codes.append(stmt)
                
                hazard_str = " / ".join(h_codes)
                if hazard_str:
                    result["위험분류"] = hazard_str
                else:
                    result["위험분류"] = "위험분류 정보 없음"
            except Exception as e:
                print(f"  Aldrich properties parsing error: {e}")

            # SDS 추출
            try:
                links = self.context.find_elements("a")
                sds_url = None
                for a in links:
                    txt = a.text.upper()
                    if "SDS" in txt or "SAFETY DATA SHEET" in txt:
                        sds_url = a.get_attribute("href")
                        break
                
                if sds_url:
                    if sds_url.startswith("/"):
                        sds_url = "https://www.sigmaaldrich.com" + sds_url
                    result["SDS_Link"] = sds_url
                    
                    filename = DBManager.clean_filename(result["시약명"])
                    if filename == "unknown":
                        filename = f"Aldrich_{product_number}"
                        
                    sds_path = os.path.join(os.getcwd(), "sds", f"{filename}.pdf")
                    os.makedirs(os.path.join(os.getcwd(), "sds"), exist_ok=True)
                    
                    cookies = {c['name']: c['value'] for c in self.context.get_cookies()}
                    # headers = {"User-Agent": self.context.get_user_agent()} # SeleniumBase는 기본적으로 자동 세팅됨
                    
                    res = requests.get(sds_url, cookies=cookies, timeout=15)
                    if res.status_code == 200:
                        with open(sds_path, 'wb') as f:
                            f.write(res.content)
                        result["SDS_Local_Path"] = sds_path
            except Exception as e:
                print(f"  SDS Download Error: {e}")

        except Exception as e:
            print(f"  Aldrich scraping error: {e}")
            
        return result
