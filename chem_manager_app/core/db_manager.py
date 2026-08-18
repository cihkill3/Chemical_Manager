import datetime
import re

class DBManager:
    COLUMN_MAP = {
        "제조사": "Manufacturer",
        "제품번호": "Catalog No.",
        "시약명": "Product Name",
        "CAS Number": "CAS No.",
        "CAS 번호": "CAS No.",
        "보관온도": "Storage Temp.",
        "신호어": "Signal Word",
        "주요위험": "Key Hazards",
        "상세 위험분류": "Detailed Hazard Classification",
        "민감성": "Sensitivity",
        "상세정보_링크": "Detail_Link",
        "갱신일": "Revision Date"
    }

    COLUMNS = [
        "Key", "Manufacturer", "Catalog No.", "Product Name", "CAS No.", 
        "Storage Temp.", "Sensitivity", "Signal Word", "Key Hazards", "Detailed Hazard Classification", "Detail_Link", "SDS_Link", 
        "SDS_Local_Path", "Revision Date"
    ]

    @staticmethod
    def normalize_manufacturer(m_name):
        from scrapers.registry import normalize_manufacturer
        return normalize_manufacturer(m_name)

    @staticmethod
    def crawl_key(manufacturer, catalog_no):
        """Stable identity used to deduplicate crawl work in one run."""
        from scrapers.registry import product_key
        return product_key(manufacturer, catalog_no)

    @staticmethod
    def needs_recrawl(record):
        """Detect clearly corrupted/incomplete crawl records without refreshing healthy rows."""
        product_name = str(record.get("Product Name", "") or "").strip()
        detail_link = str(record.get("Detail_Link", "") or "").strip()
        sds_link = str(record.get("SDS_Link", "") or "").strip()
        catalog = str(record.get("Catalog No.", "") or "").strip()
        corrupted_markers = ("@charset", "header.min.css", "body.overflow-hidden", "/* hash:")
        if any(marker in product_name.casefold() for marker in corrupted_markers):
            return True
        if "thermofisher.com/search/browse/results" in sds_link.casefold():
            return True
        return False

    @staticmethod
    def clean_filename(filename):
        """파일명으로 사용할 수 없는 특수문자를 제거합니다."""
        if not filename or not isinstance(filename, str) or filename == "제조사 홈페이지에서 검색 실패":
            return "unknown"
        cleaned = re.sub(r'[\\/*?:"<>|\r\n]', "", str(filename))
        cleaned = cleaned.strip()
        if not cleaned:
            return "unknown"
        return cleaned

    @staticmethod
    def format_sds_filename(product_name, manufacturer, catalog_no):
        """Formats SDS filename as: Product Name (Manufacturer, Catalog No).pdf"""
        p_name = DBManager.clean_filename(product_name)
        mfr = DBManager.clean_filename(manufacturer)
        cat = DBManager.clean_filename(catalog_no)
        
        if p_name == "unknown" or not p_name:
            p_name = "Product"
        if mfr == "unknown" or not mfr:
            mfr = "Manufacturer"
        if cat == "unknown" or not cat:
            cat = "CatalogNo"
            
        return f"{p_name} ({mfr}, {cat})"

    @staticmethod
    def is_sds_fresh(sds_path, max_days=180):
        """
        Checks if sds_path exists, is non-empty, and was modified within max_days (default 180 days = 6 months).
        """
        if not sds_path or not isinstance(sds_path, str) or sds_path in ["-", "", "nan", "None"]:
            return False
        import os, datetime
        if not os.path.exists(sds_path) or os.path.getsize(sds_path) == 0:
            return False
        try:
            mtime = datetime.datetime.fromtimestamp(os.path.getmtime(sds_path))
            age_days = (datetime.datetime.now() - mtime).days
            return age_days < max_days
        except Exception:
            return False

    @staticmethod
    def normalize_revision_date(value):
        """Return an Excel/date/datetime/string revision value as YYYY-MM-DD."""
        if value in (None, "", "-", "None", "nan"):
            return "" if value in (None, "", "None", "nan") else "-"
        if hasattr(value, "strftime"):
            return value.strftime("%Y-%m-%d")
        if isinstance(value, (int, float)) and not isinstance(value, bool):
            # Excel's 1900 date system (including its historical leap-year bug).
            date_value = datetime.datetime(1899, 12, 30) + datetime.timedelta(days=float(value))
            return date_value.strftime("%Y-%m-%d")
        text = str(value).strip()
        match = re.match(r"^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})", text)
        if match:
            year, month, day = (int(part) for part in match.groups())
            return f"{year:04d}-{month:02d}-{day:02d}"
        return text.split()[0]

    @staticmethod
    def normalize_temperature(temp_text):
        if not temp_text:
            return "-"
            
        t_str = str(temp_text).strip()
        t_lower = t_str.lower()

        if t_lower in ["none", "nan", "n/a", "정보 없음", "검색 실패", "search failed", "", "unknown", "-"]:
            return "-"

        if t_str in ["RT", "R", "F", "DF"]:
            return t_str

        if any(kw in t_lower for kw in ["deep freezer", "ultracold", "deep-freezer", "초저온"]):
            return "DF"
        if any(kw in t_lower for kw in ["frozen", "freezer", "freeze", "냉동"]):
            return "F"
        if any(kw in t_lower for kw in ["refrigerat", "refrigerator", "cold", "chilled", "cool", "냉장"]):
            return "R"

        if any(kw in t_lower for kw in ["room temp", "ambient", "실온", "상온"]) or re.search(r'\brt\b', t_lower):
            return "RT"

        normalized = re.sub(r'(\d+)\s*[\-~～to]+\s*(\d+)', r'\1 to \2', t_str)

        numbers = [float(n) for n in re.findall(r'[-+]?\d+(?:\.\d+)?', normalized)]
        
        if numbers:
            max_temp = max(numbers)
            if max_temp >= 21:
                return "RT"
            elif max_temp >= 0:
                return "R"
            elif max_temp >= -30:
                return "F"
            else:
                return "DF"

        return "-"

    @staticmethod
    def extract_sensitivity(text):
        if not text: return "-"
        t = str(text).lower()
        if t in ["none", "nan", "n/a", "정보 없음", "검색 실패", "search failed", "", "-"]:
            return "-"
            
        sensitivities = []
        
        if "moisture" in t or "hygroscopic" in t or "흡습" in t or "수분" in t:
            sensitivities.append("Moisture")
        if "light" in t or "광" in t or "차광" in t:
            sensitivities.append("Light")
        if "air" in t or "oxygen" in t or "공기" in t or "산소" in t:
            sensitivities.append("Air")
        if "heat" in t or "열" in t:
            sensitivities.append("Heat")
        if "argon" in t or "nitrogen" in t or "불활성" in t:
            sensitivities.append("Ar/N2 Required")
            
        if not sensitivities:
            return "-"
        return ", ".join(sensitivities)
