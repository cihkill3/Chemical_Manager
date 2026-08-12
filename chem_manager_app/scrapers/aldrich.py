import os
import requests
import re
from scrapers.base_scraper import BaseScraper
from core.db_manager import DBManager

class AldrichScraper(BaseScraper):
    def scrape(self, product_number):
        # Aldrich 제품 번호의 뒤에 -10G 등 패키지 크기가 붙는 경우가 많으므로 제거합니다.
        clean_product_number = product_number.split('-')[0]
        
        brands_to_try = [
            f"aldrich/{clean_product_number}",
            f"mm/{clean_product_number.replace('.', '')}",
            f"mm/{clean_product_number}",
            f"sial/{clean_product_number}",
            f"sigma/{clean_product_number}",
            f"supelco/{clean_product_number}"
        ]
        
        current_url = ""
        title_text = "제조사 홈페이지에서 검색 실패"
        found = False
        matched_brand = "aldrich"
        
        for brand_path in brands_to_try:
            url = f"https://www.sigmaaldrich.com/KR/en/product/{brand_path}"
            try:
                self.context.get(url)
                self.context.sleep(3)
                
                title_text = self.context.get_title().lower()
                
                # Check for 404 or missing product
                if "404" in title_text or "not found" in title_text or "error" in title_text or "sigma-aldrich" == title_text.strip():
                    continue
                else:
                    found = True
                    current_url = url
                    matched_brand = brand_path.split('/')[0]
                    break
            except Exception as e:
                print(f"    Timeout or error on {url}: {e}")
                continue
                
        if not found:
            return {
                "Manufacturer": "Aldrich", "Catalog No.": product_number, "Product Name": "",
                "CAS No.": "Search Failed", "Storage Temp.": "-", "Signal Word": "-", "Key Hazards": "-",
                "Detailed Hazard Classification": "-", "Sensitivity": "-", "Detail_Link": "Product Not Found", "SDS_Link": "-", "SDS_Local_Path": "-"
            }
            
        result = {
            "Manufacturer": "Aldrich",
            "Catalog No.": product_number,
            "Product Name": title_text.title(),
            "CAS No.": "",
            "Storage Temp.": "",
            "Signal Word": "",
            "Key Hazards": "",
            "Detailed Hazard Classification": "",
            "Sensitivity": "",
            "Detail_Link": current_url,
            "SDS_Link": "",
            "SDS_Local_Path": ""
        }
        
        try:
            html = self.context.get_page_source()
            from bs4 import BeautifulSoup
            soup = BeautifulSoup(html, 'html.parser')
            
            try:
                h1 = soup.find('h1')
                if h1:
                    name_text = h1.text.strip()
                    result["Product Name"] = name_text
            except:
                pass

            sensitivities = []
            hazard_statements = []
            
            for tr in soup.find_all('tr'):
                text = tr.text.lower()
                if "cas number" in text or "cas-no." in text or "cas 번호" in text or "cas번호" in text:
                    m = re.search(r'\d{2,7}-\d{2}-\d', text)
                    if m:
                        cas_val = m.group(0)
                        result["CAS No."] = cas_val
                elif ("storage temp" in text or "보관 온도" in text or "보관온도" in text) and "rt" not in str(result["Storage Temp."]).lower():
                    temp_val = DBManager.normalize_temperature(text)
                    result["Storage Temp."] = temp_val
                
                if "sensitive" in text or "hygroscopic" in text:
                    if "light sensitive" in text: sensitivities.append("Light")
                    if "moisture sensitive" in text: sensitivities.append("Moisture")
                    if "air sensitive" in text: sensitivities.append("Air")
                    if "heat sensitive" in text: sensitivities.append("Heat")
                    if "hygroscopic" in text: sensitivities.append("Hygroscopic")
                    
            # Hazard parsing (GraphQL or DOM)
            for div in soup.find_all('div'):
                text = div.text.lower()
                from scrapers.hcodes_dict import get_h_statement, GHS_H_CODES
                h_codes_raw = re.findall(r'\bh[234]\d{2}[a-z]*\b', text)
                for code in h_codes_raw:
                    if code in GHS_H_CODES:
                        stmt = get_h_statement(code)
                        if stmt not in hazard_statements:
                            hazard_statements.append(stmt)
                            
                if "signal word" in text or "신호어" in text:
                    if "danger" in text or "위험" in text:
                        result["Signal Word"] = "● Danger"
                    elif "warning" in text or "경고" in text:
                        result["Signal Word"] = "▲ Warning"

            # Fallback for CAS and Storage Temp using full page text
            full_text = soup.text.lower()
            if result["CAS No."] in ["", "-", None, "Search Failed"]:
                m_cas = re.search(r'cas\s*(?:number|-no\.|번호)?[\s:]*(\d{2,7}-\d{2}-\d)', full_text)
                if m_cas:
                    cas_val = m_cas.group(1)
                    result["CAS No."] = cas_val
            
            if result["Storage Temp."] in ["", "-", None]:
                m_temp = re.search(r'(?:storage temp|보관 온도|보관온도)[^\d\-]*([\-\d]+(?:\.\d+)?\s*°c)', full_text)
                if m_temp:
                    temp_val = DBManager.normalize_temperature(m_temp.group(1))
                    result["Storage Temp."] = temp_val

            if sensitivities:
                sens_val = ", ".join(list(dict.fromkeys(sensitivities)))
                result["Sensitivity"] = sens_val
                
            if hazard_statements:
                hazard_str = " / ".join(hazard_statements)
                from core.hazard_parser import parse_hazard
                sw, kh, _ = parse_hazard(hazard_str, result["Signal Word"])
                if not result["Signal Word"]:
                    result["Signal Word"] = sw
                result["Key Hazards"] = kh
                result["Detailed Hazard Classification"] = hazard_str
        except Exception as e:
            print(f"  [Aldrich] parsing error: {e}")

        # --- SDS DOWNLOAD (from old sds_downloader.py) ---
        try:
            import io
            
            pdf_url = f"https://www.sigmaaldrich.com/KR/ko/sds/{matched_brand}/{clean_product_number}"
            result["SDS_Link"] = pdf_url
            
            print(f"  [Aldrich] SDS Fetch from: {pdf_url}")
            
            # Use SeleniumBase execute_async_script to bypass Cloudflare
            script = f"""
                var done = arguments[arguments.length - 1];
                fetch('{pdf_url}', {{
                    headers: {{ 'Accept': 'application/pdf,*/*', 'Referer': 'https://www.sigmaaldrich.com/KR/ko' }},
                    credentials: 'include'
                }})
                .then(r => {{
                    if (!r.ok) return done({{error: 'HTTP ' + r.status}});
                    return r.arrayBuffer();
                }})
                .then(buf => {{
                    var arr = Array.from(new Uint8Array(buf));
                    done({{data: arr}});
                }})
                .catch(e => done({{error: e.toString()}}));
            """
            
            res = self.context.execute_async_script(script)
            
            if res and 'data' in res:
                content = bytes(res['data'])
                if len(content) > 1000:
                    # Validate PDF using PyMuPDF (pymupdf)
                    is_valid = False
                    try:
                        import pymupdf as fitz
                        doc = fitz.open(stream=content, filetype="pdf")
                        if len(doc) > 0:
                            text = doc[0].get_text() or ""
                            if len(text) > 10 or len(doc) >= 1:
                                is_valid = True
                        doc.close()
                    except:
                        is_valid = True # fallback if parser fails on valid pdf
                        
                    if is_valid:
                        p_name = result.get("Product Name", result.get("시약명", ""))
                        filename = DBManager.format_sds_filename(p_name, "Aldrich", clean_product_number)
                        
                        sds_dir = os.path.join(self.base_dir, "sds")
                        os.makedirs(sds_dir, exist_ok=True)
                        sds_path = os.path.join(sds_dir, f"{filename}.pdf")
                        
                        print(f"  [Aldrich] Saving SDS PDF to: {sds_path} (bytes: {len(content)})")
                        with open(sds_path, 'wb') as f:
                            f.write(content)
                        result["SDS_Local_Path"] = sds_path
                    else:
                        print(f"  [Aldrich] SDS PDF validation failed (len: {len(content)})")
            elif res and 'error' in res:
                print(f"  [Aldrich] SDS Fetch Error: {res['error']}")
                
        except Exception as e:
            print(f"  [Aldrich] SDS Error: {e}")
            
        return result
