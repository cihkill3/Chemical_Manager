import os
import requests
import re
from bs4 import BeautifulSoup
from scrapers.base_scraper import BaseScraper
from core.db_manager import DBManager

class TciScraper(BaseScraper):
    def scrape(self, product_number):
        url = f"https://www.tcichemicals.com/KR/en/p/{product_number}"
        
        result = {
            "Manufacturer": "TCI", "Catalog No.": product_number, "Product Name": "",
            "CAS No.": "Search Failed", "Storage Temp.": "-", "Signal Word": "-", "Key Hazards": "-",
            "Detailed Hazard Classification": "-", "Sensitivity": "-", "Detail_Link": "Product Not Found", "SDS_Link": "-", "SDS_Local_Path": "-"
        }
        
        try:
            self.context.get(url)
            self.context.sleep(3) # Cloudflare wait
            page_title = self.context.get_title()
            if "not found" in page_title.lower() or "error" in page_title.lower() or "access denied" in page_title.lower():
                return result

            try:
                # TCI Title Format: 1,1'-Carbonyldiimidazole | 530-62-1 | Tokyo Chemical Industry
                parts = page_title.split('|')
                if len(parts) >= 2:
                    result["Product Name"] = parts[0].strip()
                    result["CAS No."] = parts[1].strip()
                    result["Detail_Link"] = url
                else:
                    name = self.context.get_text("h1").strip()
                    if name: 
                        result["Product Name"] = name
                        result["Detail_Link"] = url
            except:
                pass
                
            # Properties parsing
            try:
                soup = BeautifulSoup(self.context.get_page_source(), 'html.parser')
                
                hazard_statements = ""
                signal_word = ""
                sensitivities = []
                
                for el in soup.find_all(['tr', 'li', 'div', 'p']):
                    row_text_raw = el.get_text(separator=" ", strip=True)
                    text = row_text_raw.lower()
                    
                    if not text:
                        continue
                        
                    if "sensitive" in text or "hygroscopic" in text or "sensitive to" in text:
                        if "light sensitive" in text or "빛에 민감" in text: sensitivities.append("Light")
                        if "moisture sensitive" in text or "수분에 민감" in text: sensitivities.append("Moisture")
                        if "air sensitive" in text or "공기에 민감" in text: sensitivities.append("Air")
                        if "heat sensitive" in text or "열에 민감" in text: sensitivities.append("Heat")
                        if "hygroscopic" in text or "흡습성" in text: sensitivities.append("Hygroscopic")
                        
                    if "storage condition" in text or "storage temperature" in text or "보관 온도" in text:
                        if "RT" not in str(result["Storage Temp."]):
                            temp_val = DBManager.normalize_temperature(row_text_raw)
                            result["Storage Temp."] = temp_val
                    elif "cas rn" in text or "cas number" in text or "cas 번호" in text or "cas" in text:
                        m_cas = re.search(r'\d{2,7}-\d{2}-\d', text)
                        if m_cas and result["CAS No."] in ["Search Failed", "", "-", None]:
                            result["CAS No."] = m_cas.group(0)
                            
                    if "signal word" in text or "신호어" in text:
                        signal_word = re.sub(r'(?i)signal\s*word|신호어', '', row_text_raw).strip()
                    elif "hazard statements" in text or "hazard statement" in text or "유해·위험 문구" in text or "유해성 분류" in text:
                        hazard_statements = re.sub(r'(?i)hazard\s*statements?|유해·위험 문구|유해성 분류', '', row_text_raw).strip()
                        
                if sensitivities:
                    sens_val = ", ".join(list(dict.fromkeys(sensitivities)))
                    result["Sensitivity"] = sens_val
                        
                if hazard_statements or signal_word:
                    combined = ""
                    if signal_word:
                        combined += f"[{signal_word}] "
                    if hazard_statements:
                        combined += hazard_statements
                        
                    if "위험" in combined or "danger" in combined.lower():
                        result["Signal Word"] = "● Danger"
                    elif "경고" in combined or "warning" in combined.lower():
                        result["Signal Word"] = "▲ Warning"
                        
                    from core.hazard_parser import parse_hazard
                    sw, kh, _ = parse_hazard(hazard_statements, result["Signal Word"])
                    if not result["Signal Word"]:
                        result["Signal Word"] = sw
                    result["Key Hazards"] = kh
                    result["Detailed Hazard Classification"] = hazard_statements
            except Exception as e:
                print(f"  TCI properties parsing error: {e}")

            # SDS 추출
            try:
                import base64
                p_name = result.get("Product Name", result.get("시약명", ""))
                filename = DBManager.format_sds_filename(p_name, "TCI", product_number)
                sds_dir = os.path.join(self.base_dir, "sds")
                os.makedirs(sds_dir, exist_ok=True)
                sds_path = os.path.join(sds_dir, f"{filename}.pdf")

                intercept_and_fetch_js = f"""
                var done = arguments[arguments.length - 1];
                var pCode = "{product_number}";
                
                window._pdfBlob = null;
                const origOpen = XMLHttpRequest.prototype.open;
                const origSend = XMLHttpRequest.prototype.send;
                XMLHttpRequest.prototype.open = function(method, url) {{
                    this._reqUrl = url;
                    return origOpen.apply(this, arguments);
                }};
                XMLHttpRequest.prototype.send = function(body) {{
                    this.addEventListener('load', function() {{
                        if (this.responseType === 'blob' && this.response) {{
                            window._pdfBlob = this.response;
                        }}
                    }});
                    return origSend.apply(this, arguments);
                }};
                
                var btn = document.getElementById('sdsSearchButton') || document.querySelector('.ViewDocument');
                var lang = document.getElementById('langSelector');
                if (lang) lang.value = 'ko';
                if (btn) {{
                    btn.click();
                }} else {{
                    var contextPath = (typeof ACC !== 'undefined' && ACC.config && ACC.config.encodedContextPath) ? ACC.config.encodedContextPath : '/KR/en';
                    var postUrl = contextPath + "/documentSearch/productSDSSearchDoc";
                    
                    var xhr = new XMLHttpRequest();
                    xhr.open("POST", postUrl, true);
                    xhr.setRequestHeader("Content-Type", "application/x-www-form-urlencoded; charset=UTF-8");
                    xhr.setRequestHeader("X-Requested-With", "XMLHttpRequest");
                    xhr.responseType = "blob";
                    xhr.onload = function() {{
                        if (xhr.status === 200 && xhr.response) {{
                            window._pdfBlob = xhr.response;
                        }}
                    }};
                    xhr.send("brandCode=TCI&productCode=" + encodeURIComponent(pCode) + "&langSelector=ko&selectedCountry=KR");
                }}
                
                var count = 0;
                var interval = setInterval(function() {{
                    count++;
                    if (window._pdfBlob) {{
                        clearInterval(interval);
                        var reader = new FileReader();
                        reader.readAsDataURL(window._pdfBlob);
                        reader.onloadend = function() {{
                            done(reader.result);
                        }};
                    }} else if (count >= 10) {{
                        clearInterval(interval);
                        done(null);
                    }}
                }}, 500);
                """

                b64_res = self.context.execute_async_script(intercept_and_fetch_js)
                
                if not b64_res or not isinstance(b64_res, str) or ',' not in b64_res:
                    doc_search_url = "https://www.tcichemicals.com/KR/en/documentSearch"
                    self.context.get(doc_search_url)
                    self.context.sleep(2)
                    
                    doc_search_js = f"""
                    var done = arguments[arguments.length - 1];
                    var pCode = "{product_number}";
                    window._pdfBlob = null;
                    const origOpen = XMLHttpRequest.prototype.open;
                    const origSend = XMLHttpRequest.prototype.send;
                    XMLHttpRequest.prototype.open = function(method, url) {{
                        this._reqUrl = url;
                        return origOpen.apply(this, arguments);
                    }};
                    XMLHttpRequest.prototype.send = function(body) {{
                        this.addEventListener('load', function() {{
                            if (this.responseType === 'blob' && this.response) {{
                                window._pdfBlob = this.response;
                            }}
                        }});
                        return origSend.apply(this, arguments);
                    }};
                    
                    var pInput = document.getElementById('sdsProductCode');
                    var lSelect = document.getElementById('langSelector');
                    var btn = document.getElementById('sdsSearchButton');
                    
                    if (pInput) pInput.value = pCode;
                    if (lSelect) lSelect.value = 'ko';
                    if (btn) btn.click();
                    
                    var count = 0;
                    var interval = setInterval(function() {{
                        count++;
                        if (window._pdfBlob) {{
                            clearInterval(interval);
                            var reader = new FileReader();
                            reader.readAsDataURL(window._pdfBlob);
                            reader.onloadend = function() {{
                                done(reader.result);
                            }};
                        }} else if (count >= 10) {{
                            clearInterval(interval);
                            done(null);
                        }}
                    }}, 500);
                    """
                    b64_res = self.context.execute_async_script(doc_search_js)

                if b64_res and isinstance(b64_res, str) and ',' in b64_res:
                    b64_data = b64_res.split(',')[1]
                    content = base64.b64decode(b64_data)
                    
                    is_valid = False
                    try:
                        import pymupdf as fitz
                        doc = fitz.open(stream=content, filetype="pdf")
                        if len(doc) > 0:
                            is_valid = True
                        doc.close()
                    except:
                        if content[:4] == b'%PDF':
                            is_valid = True

                    if is_valid:
                        with open(sds_path, 'wb') as f:
                            f.write(content)
                        result["SDS_Local_Path"] = sds_path
                        result["SDS_Link"] = "https://www.tcichemicals.com/KR/en/documentSearch"
            except Exception as e:
                print(f"  SDS Download Error: {e}")

        except Exception as e:
            print(f"  TCI scraping error: {e}")
            
        return result
