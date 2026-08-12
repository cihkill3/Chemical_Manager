import os
import requests
from scrapers.base_scraper import BaseScraper
from core.db_manager import DBManager

class ThermofisherScraper(BaseScraper):
    def scrape(self, product_number):
        url = f"https://chemicals.thermofisher.kr/apac/product/{product_number}"
        
        result = {
            "Manufacturer": "ThermoFisher", "Catalog No.": product_number, "Product Name": "",
            "CAS No.": "Search Failed", "Storage Temp.": "-", "Signal Word": "-", "Key Hazards": "-",
            "Detailed Hazard Classification": "-", "Sensitivity": "-", "Detail_Link": "Product Not Found", "SDS_Link": "-", "SDS_Local_Path": "-"
        }
        
        try:
            suffixes = ["", ".MD", ".06", ".14", ".36", ".18", ".22", ".03"]
            found = False
            for suffix in suffixes:
                current_url = f"https://chemicals.thermofisher.kr/apac/product/{product_number}{suffix}"
                try:
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
                        result["Product Name"] = name
                        result["Detail_Link"] = current_url
                        found = True
                        break
                except:
                    pass
            
            if not found:
                # Fallback to global thermofisher.com
                global_url = f"https://www.thermofisher.com/order/catalog/product/{product_number}"
                try:
                    self.context.get(global_url)
                    self.context.sleep(3)
                except Exception as e:
                    print(f"    Timeout or error on global {global_url}: {e}")
                
                try:
                    name = self.context.get_text("h1").strip()
                    if name:
                        result["Product Name"] = name
                        result["Detail_Link"] = global_url
                        found = True
                except:
                    pass
            
            if not found:
                return result
                
            # Properties parsing
            try:
                html = self.context.get_page_source()
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
                        if m and result["CAS No."] in ["Search Failed", "", "-", None]:
                            result["CAS No."] = m.group(0)
                    elif ("temperature" in text or "보관온도" in text) and "rt" not in str(result["Storage Temp."]).lower():
                        temp_val = DBManager.normalize_temperature(text)
                        result["Storage Temp."] = temp_val
                            
                    if "sensitive" in text or "hygroscopic" in text or "sensitive to" in text:
                        if "light sensitive" in text or "빛에 민감" in text: sensitivities.append("Light")
                        if "moisture sensitive" in text or "수분에 민감" in text or "흡습성" in text: sensitivities.append("Moisture")
                        if "air sensitive" in text or "공기에 민감" in text: sensitivities.append("Air")
                        if "heat sensitive" in text or "열에 민감" in text: sensitivities.append("Heat")
                        if "hygroscopic" in text: sensitivities.append("Hygroscopic")
                        
                    from scrapers.hcodes_dict import get_h_statement, GHS_H_CODES
                    # Extract H-codes
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
                print(f"  TF properties parsing error: {e}")

            # --- SDS DOWNLOAD (from old sds_downloader.py) ---
            try:
                import io
                
                sds_url_found = None
                child_skus = f"{product_number}.MF,{product_number}.03,{product_number}.MD,{product_number}.06,{product_number}.14,{product_number}"
                headers = {
                    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
                    "Accept": "application/json"
                }
                
                for lang in ['ko', 'en']:
                    api_url = f"https://chemicals.thermofisher.kr/apac/api/document/search/sds?childSkus={child_skus}&language={lang}"
                    try:
                        api_resp = requests.get(api_url, headers=headers, timeout=10)
                        if api_resp.status_code == 200:
                            data = api_resp.json()
                            pdf_url = data.get("data")
                            if pdf_url and isinstance(pdf_url, str) and pdf_url.startswith("http"):
                                sds_url_found = pdf_url
                                break
                    except:
                        pass
                
                if not sds_url_found:
                    base_sku = product_number
                    if "-" in product_number:
                        parts = product_number.split("-")
                        if len(parts) >= 2:
                            base_sku = f"{parts[0]}-{parts[1]}"
                    fallback_urls = [
                        f"https://documents.thermofisher.com/TFS-Assets/LSG/SDS/{base_sku}_MTR-NALT_EN.pdf",
                        f"https://documents.thermofisher.com/TFS-Assets/LSG/SDS/{product_number}_MTR-NALT_EN.pdf"
                    ]
                    for f_url in fallback_urls:
                        try:
                            pdf_resp = requests.get(f_url, headers=headers, timeout=10, allow_redirects=True)
                            if pdf_resp.status_code == 200 and 'application/pdf' in pdf_resp.headers.get('content-type', '').lower():
                                sds_url_found = f_url
                                break
                        except:
                            pass
                
                if sds_url_found:
                    result["SDS_Link"] = sds_url_found
                    
                    p_name = result.get("Product Name", result.get("시약명", ""))
                    filename = DBManager.format_sds_filename(p_name, "ThermoFisher", product_number)
                        
                    sds_dir = os.path.join(self.base_dir, "sds")
                    os.makedirs(sds_dir, exist_ok=True)
                    sds_path = os.path.join(sds_dir, f"{filename}.pdf")
                    
                    pdf_resp = requests.get(sds_url_found, headers=headers, timeout=30, allow_redirects=True)
                    if pdf_resp.status_code == 200:
                        content = pdf_resp.content
                        
                        # Validate PDF using PyMuPDF (pymupdf)
                        is_valid = False
                        try:
                            import pymupdf as fitz
                            doc = fitz.open(stream=content, filetype="pdf")
                            if len(doc) > 0:
                                is_valid = True
                            doc.close()
                        except:
                            is_valid = True
                            
                        if is_valid:
                            with open(sds_path, 'wb') as f:
                                f.write(content)
                            result["SDS_Local_Path"] = sds_path
                else:
                    result["SDS_Link"] = f"https://www.thermofisher.com/search/browse/results?term={product_number}"
            except Exception as e:
                print(f"  TF SDS Error: {e}")
                
        except Exception as e:
            print(f"  ThermoFisher scraping error: {e}")
            
        return result
