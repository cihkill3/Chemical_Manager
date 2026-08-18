import time
import os
import re
import json
import urllib.parse
from bs4 import BeautifulSoup
from .base_scraper import BaseScraper
import logging

logger = logging.getLogger(__name__)

class AbcamScraper(BaseScraper):
    coa_vendor = "Abcam"

    @staticmethod
    def _sds_url_from_html(html, current_url):
        soup = BeautifulSoup(html or "", "html.parser")
        for link in soup.find_all("a", href=True):
            href = link.get("href", "").strip()
            label = link.get_text(" ", strip=True).casefold()
            lowered = href.casefold()
            if (
                "wercs-api-prod-bucket" in lowered
                or ("sds" in label and (".pdf" in lowered or "sds" in lowered))
            ):
                return urllib.parse.urljoin(current_url, href)

        match = re.search(
            r'https?://[^"\'<>\\\s]*wercs-api-prod-bucket[^"\'<>\\\s]*\.pdf(?:\?[^"\'<>\\\s]*)?',
            html or "",
            flags=re.IGNORECASE,
        )
        return match.group(0).replace("&amp;", "&") if match else ""

    def _download_sds(self, result, product_number, html, current_url):
        from core.db_manager import DBManager

        fresh_path = self.find_fresh_sds("Abcam", product_number)
        if fresh_path:
            result["SDS_Local_Path"] = fresh_path
            return

        sds_url = self._sds_url_from_html(html, current_url)
        if not sds_url:
            try:
                product_code_json = json.dumps(product_number.lower())
                api_script = """
                    const productCode = __PRODUCT_CODE__;
                    const done = arguments[arguments.length - 1];
                    const query = `query EDS_SDS {
                      document(filter: { productCode: "${productCode}", countryCode: "KR" }) {
                        sds { name displayName url countryCode languageCode }
                      }
                    }`;
                    fetch('https://proxy-gateway.abcam.com/product', {
                      method: 'POST',
                      headers: {
                        'Content-Type': 'application/json',
                        'X-Abcam-App-Id': 'b2c-public-website'
                      },
                      body: JSON.stringify({query})
                    }).then(response => {
                      if (!response.ok) throw new Error(`HTTP ${response.status}`);
                      return response.json();
                    }).then(payload => {
                      const documents = payload?.data?.document?.sds || [];
                      done({documents});
                    }).catch(error => done({error: error.toString()}));
                """.replace("__PRODUCT_CODE__", product_code_json)
                api_result = self.execute_async_script(api_script)
                documents = api_result.get("documents", []) if isinstance(api_result, dict) else []
                preferred = next(
                    (doc for doc in documents if str(doc.get("languageCode", "")).lower() == "en"),
                    documents[0] if documents else None,
                )
                sds_url = preferred.get("url", "") if preferred else ""
            except Exception as error:
                logger.warning("[Abcam] SDS API lookup failed for %s: %s", product_number, error)
                sds_url = ""
        if not sds_url:
            logger.warning("[Abcam] SDS link not found for %s", product_number)
            return

        result["SDS_Link"] = sds_url
        try:
            response = self.http_get(
                sds_url,
                headers={"User-Agent": "Mozilla/5.0"},
                timeout=30,
                allow_redirects=True,
            )
            response.raise_for_status()
            content = response.content
            if len(content) < 100 or not content.startswith(b"%PDF"):
                raise ValueError("downloaded content is not a PDF")

            product_name = result.get("Product Name", "")
            filename = DBManager.format_sds_filename(product_name, "Abcam", product_number)
            sds_dir = os.path.join(self.base_dir, "sds")
            os.makedirs(sds_dir, exist_ok=True)
            sds_path = os.path.abspath(os.path.join(sds_dir, f"{filename}.pdf"))
            with open(sds_path, "wb") as output:
                output.write(content)
            result["SDS_Local_Path"] = sds_path
            logger.info("[Abcam] SDS saved: %s", sds_path)
        except Exception as error:
            logger.warning("[Abcam] SDS download failed for %s: %s", product_number, error)

    def scrape(self, product_number):
        product_number = product_number.strip().lower()
        url = f"https://www.abcam.com/ko-kr/products/search?keywords={product_number}"
        short_url = f"https://www.abcam.com/{product_number}"
        logger.info(f"[Abcam] Visiting shortlink for {product_number}: {short_url}")
        
        try:
            self.context.get(short_url)
            self.wait_for_page(timeout=5, reject_titles=("just a moment", "checking your browser"))
            
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

            self._download_sds(data, product_number, self.context.get_page_source(), current_url)
                
            return data
            
        except Exception as e:
            logger.error(f"[Abcam] Scraping error: {e}")
            return {
                "Manufacturer": "Abcam", "Catalog No.": product_number, "Product Name": "",
                "CAS No.": "Search Failed", "Storage Temp.": "-", "Signal Word": "-", "Key Hazards": "-",
                "Detailed Hazard Classification": "-", "Sensitivity": "-", "Detail_Link": "Product Not Found", "SDS_Link": "-", "SDS_Local_Path": "-"
            }
