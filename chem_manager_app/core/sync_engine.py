import os
import time
import shutil
import datetime
import win32com.client
from utils.color_utils import hex_to_bgr
from utils.excel_utils import get_col_letter

class SyncEngine:
    @staticmethod
    def get_col_idx(names, col_map):
        for n in names:
            if n in col_map and col_map[n] > 0:
                return col_map[n]
        return 0

    def __init__(self, config_data, callback_progress=None):
        self.config = config_data
        self.callback = callback_progress
        self.excel = None
        self.source_wb = None
        self.target_wb = None

    def log(self, message):
        if self.callback:
            self.callback(message)
        else:
            print(message)

    def safe_set_val(self, cell_or_range, val, retries=5):
        for attempt in range(retries):
            try:
                cell_or_range.Value = val
                return
            except Exception as e:
                if attempt == retries - 1:
                    raise e
                time.sleep(0.1)

    def run_sync(self):
        import urllib.parse
        src_path_raw = self.config.get("source_file", "")
        if src_path_raw.startswith("file:///"):
            src_path_raw = urllib.parse.unquote(src_path_raw[8:])
        src_path = os.path.abspath(src_path_raw)
        
        src_sheet_name = self.config.get("source_sheet", "")
        header_row = int(self.config.get("header_row", 1))

        if not os.path.exists(src_path):
            raise Exception("원본 파일을 찾을 수 없습니다.")

        src_folder = os.path.dirname(src_path)
        
        target_path = os.path.join(src_folder, "ChemicalList.xlsx")
        self.target_file = target_path
        
        # 백업 로직은 실제 저장 시점으로 이동했습니다.

        try:
            self.excel = win32com.client.DispatchEx("Excel.Application")
            try:
                import win32process
                _, self.excel_pid = win32process.GetWindowThreadProcessId(self.excel.Hwnd)
            except:
                self.excel_pid = None
                
            self.excel.Visible = False
            self.excel.ScreenUpdating = False
            self.excel.DisplayAlerts = False
            
            self.log("원본 파일 여는 중...")
            self.source_wb = self.excel.Workbooks.Open(src_path, ReadOnly=False, UpdateLinks=False)
            
            self.source_is_readonly = self.source_wb.ReadOnly
            if self.source_is_readonly:
                self.log("원본 파일이 현재 열려있어 읽기 전용으로 접근합니다. (초록색 완료 표시는 생략됨)")
            
            src_ws = None
            for sheet in self.source_wb.Worksheets:
                if sheet.Name == src_sheet_name:
                    src_ws = sheet
                    break
            
            if not src_ws:
                raise Exception(f"'{src_sheet_name}' 시트를 찾을 수 없습니다.")

            # Target Setup
            self.log("최신 ChemicalList 파일 확인 중...")
            if target_path and os.path.exists(target_path):
                self.log(f"타겟 파일 열기: {os.path.basename(target_path)}")
                if os.path.exists(src_path) and os.path.samefile(src_path, target_path):
                    self.target_wb = self.source_wb
                else:
                    self.target_wb = self.excel.Workbooks.Open(target_path)
                if self.target_wb.ReadOnly:
                    self.log("최신 ChemicalList 파일이 현재 사용 중(열림)입니다. (읽기 전용으로 작업 후 새 파일로 저장됩니다)")
                tgt_ws = self.target_wb.Sheets(1)
            else:
                self.log("ChemicalList 파일이 없어 새로 생성합니다.")
                self.target_wb = self.excel.Workbooks.Add()
                tgt_ws = self.target_wb.Sheets(1)
                tgt_ws.Name = "ChemicalList"
                # Write Headers
                for idx, h in enumerate(self.config["target_headers"]):
                    tgt_ws.Cells(1, idx + 1).Value = h
                
                hrange = tgt_ws.Range(tgt_ws.Cells(1, 1), tgt_ws.Cells(1, len(self.config["target_headers"])))
                c_dict = self.config.get("colors", {})
                hrange.Font.Bold = True
                hrange.Font.Color = hex_to_bgr(c_dict.get("header_font", "#ffffff"))
                hrange.Interior.Color = hex_to_bgr(c_dict.get("header_bg", "#464646"))
                hrange.HorizontalAlignment = -4108 # xlCenter

            # Build Target Indices dynamically from actual sheet row 1
            tgt_headers = self.config["target_headers"]
            tgt_col_map = {}
            tgt_last_col = tgt_ws.Cells(1, tgt_ws.Columns.Count).End(-4159).Column
            for c in range(1, max(tgt_last_col + 1, 30)):
                val = tgt_ws.Cells(1, c).Value
                if val:
                    tgt_col_map[str(val).strip()] = c
                    
            for idx, name in enumerate(tgt_headers):
                if name not in tgt_col_map:
                    tgt_col_map[name] = idx + 1
            
            # Ensure hidden Original Product Name column exists
            orig_pn_col = tgt_col_map.get("Original Product Name", 0)
            if orig_pn_col == 0:
                orig_pn_col = max(tgt_col_map.values()) + 1 if tgt_col_map else 16
                tgt_ws.Cells(1, orig_pn_col).Value = "Original Product Name"
                tgt_col_map["Original Product Name"] = orig_pn_col
            try:
                tgt_ws.Columns(orig_pn_col).Hidden = True
            except: pass
            
            tgt_order_num_col = tgt_col_map.get("Order No.", 0)
            if tgt_order_num_col == 0:
                raise Exception("Target 헤더에 'Order No.' 가 없습니다.")

            # Find Source Columns dynamically based on row
            src_last_col = src_ws.Cells(header_row, src_ws.Columns.Count).End(-4159).Column # xlToLeft
            src_col_map = {}
            for c in range(1, src_last_col + 1):
                val = str(src_ws.Cells(header_row, c).Value).strip()
                if val:
                    src_col_map[val] = c
            
            src_order_num_col = src_col_map.get(self.config["mapping"].get("Order No.", ""), 0)
            if src_order_num_col == 0:
                raise Exception("원본 파일에서 주문 번호 매핑 열을 찾을 수 없습니다.")

            src_last_row = src_ws.Cells(src_ws.Rows.Count, src_order_num_col).End(-4162).Row # xlUp
            abs_src_last_row = src_ws.Cells.Find("*", SearchOrder=1, SearchDirection=2)
            if abs_src_last_row:
                src_last_row = max(src_last_row, abs_src_last_row.Row)
                
            if src_last_row <= header_row:
                raise Exception("원본 시트에 데이터가 없습니다.")

            # Read Source Data
            self.log("원본 데이터 읽기 중...")
            src_data = src_ws.Range(src_ws.Cells(header_row + 1, 1), src_ws.Cells(src_last_row, src_last_col)).Value
            
            if src_data is None:
                src_data = []
            elif not isinstance(src_data[0], (tuple, list)):
                src_data = [src_data]
            
            # Read Target Data
            tgt_prod_col = tgt_col_map.get("Product Name", 0)
            if tgt_prod_col > 0:
                tgt_last_row = tgt_ws.Cells(tgt_ws.Rows.Count, tgt_prod_col).End(-4162).Row
            else:
                tgt_last_row = tgt_ws.Cells(tgt_ws.Rows.Count, tgt_order_num_col).End(-4162).Row

            tgt_dict = {}
            if tgt_last_row >= 2:
                for r in range(2, tgt_last_row + 1):
                    val = str(tgt_ws.Cells(r, tgt_order_num_col).Value).strip()
                    if val and val != "None":
                        if val not in tgt_dict:
                            tgt_dict[val] = r

            mapping = self.config["mapping"]
            
            cnt_new = 0
            cnt_upd = 0
            cnt_dup = 0
            cnt_db_upd = 0
            src_done_rows = []

            from core.db_manager import DBManager
            total_updates = 0
            for row_data in src_data:
                if not row_data: continue
                order_num_str = str(row_data[src_order_num_col - 1]).strip() if src_order_num_col - 1 < len(row_data) else ""
                if not order_num_str or order_num_str == "None": continue
                if order_num_str not in tgt_dict:
                    total_updates += 1
            fast_mode = total_updates <= 5

            from scrapers.thermofisher import ThermofisherScraper
            from scrapers.tci import TciScraper
            from scrapers.aldrich import AldrichScraper
            from scrapers.abcam import AbcamScraper

            # db_manager는 target_wb(즉, ChemicalList.xlsx)에 쓰거나 읽을 때 열어둔 파일과 충돌날 수 있습니다.
            # 하지만 openpyxl은 별도로 열고 닫으므로, 엑셀 앱에서 열기 전에 미리 처리하거나 
            # 일단 여기서는 pandas 읽기(load_db)만 하고 저장은 다른 곳에 위임하는 식으로 해야합니다.
            # win32com에서 이미 열었기 때문에 openpyxl에서 쓰기 권한이 막힐 수 있습니다.
            # 이를 해결하기 위해 win32com 엑셀 파일은 임시로 생성하거나 다루고, 
            # DB 시트는 win32com 객체를 이용해 읽고 쓰는 방식으로 해야 충돌이 나지 않습니다.
            
            # TODO: win32com 내에서 DB 시트를 조작하도록 DB 처리 방식을 보강했습니다.
            self.log("데이터 동기화 진행 중...")
            
            # Find specific cols for logic
            qty_col = tgt_col_map.get("Quantity", 0)
            used_col = tgt_col_map.get("Used", 0)
            stat_col = tgt_col_map.get("Status", 0)
            
            receipt_src_col = src_col_map.get("수령확인", 0)
            chem_src_col = src_col_map.get("품목명", src_col_map.get("시약", 0))

            for i, row_data in enumerate(src_data):
                if not row_data:
                    continue
                
                # src_data is 0-indexed tuple of tuples
                order_num_val = row_data[src_order_num_col - 1] if src_order_num_col - 1 < len(row_data) else None
                if not order_num_val:
                    continue
                order_num_str = str(order_num_val).strip()
                if not order_num_str or order_num_str == "None":
                    continue

                # Check if this row meets inclusion criteria
                if receipt_src_col == 0:
                    continue  # "수령확인" 열을 찾을 수 없으면 해당 행 스킵
                receipt_val = row_data[receipt_src_col - 1]
                if str(receipt_val).strip().upper() != "O" and str(receipt_val).strip() != "ㅇ":
                    continue # Not received yet
                
                # Prepare identifier for logging
                chem_name = str(row_data[chem_src_col - 1]).strip() if chem_src_col > 0 and row_data[chem_src_col - 1] is not None else "Unknown"
                identifier = f"[{chem_name}] (Order: {order_num_str})"

                exists = order_num_str in tgt_dict
                
                if exists:
                    tgt_r = tgt_dict[order_num_str]
                    changed_cols = []
                    for t_header, s_header in mapping.items():
                        if not s_header: continue
                        if t_header == "Status": continue # Status is handled by formula
                        t_c = tgt_col_map.get(t_header, 0)
                        s_c = src_col_map.get(s_header, 0)
                        if t_c > 0 and s_c > 0 and s_c - 1 < len(row_data):
                            new_val = row_data[s_c - 1]
                            if new_val is None: new_val = ""
                            old_val = tgt_ws.Cells(tgt_r, t_c).Value
                            if old_val is None: old_val = ""
                            
                            # 기존 값이 이미 적혀있다면 덮어쓰지 않음
                            if str(old_val).strip() != "" and str(old_val).strip() != "None":
                                continue
                                
                            if str(old_val).strip() != str(new_val).strip() and str(new_val).strip() != "":
                                changed_cols.append(f"{t_header}: 빈칸 -> '{str(new_val).strip()}'")
                                tgt_ws.Cells(tgt_r, t_c).Value = new_val
                                
                    if changed_cols:
                        self.log(f"  [업데이트] {identifier} - 변경항목: " + ", ".join(changed_cols))
                        cnt_upd += 1
                    else:
                        cnt_dup += 1
                else:
                    self.log(f"  [신규추가] {identifier}")
                    tgt_last_row += 1
                    tgt_r = tgt_last_row
                    tgt_dict[order_num_str] = tgt_r
                    
                    for t_header, s_header in mapping.items():
                        if not s_header: continue
                        if t_header == "Status": continue
                        t_c = tgt_col_map.get(t_header, 0)
                        s_c = src_col_map.get(s_header, 0)
                        if t_c > 0 and s_c > 0 and s_c - 1 < len(row_data):
                            new_val = row_data[s_c - 1]
                            if new_val is not None:
                                tgt_ws.Cells(tgt_r, t_c).Value = new_val
                                
                tgt_ws.Range(tgt_ws.Cells(tgt_r, 1), tgt_ws.Cells(tgt_r, len(tgt_headers))).HorizontalAlignment = -4108
                if not exists:
                    cnt_new += 1

                src_done_rows.append(header_row + 1 + i)

            self.log("Chemical List 내 전체 시약 정보 크롤링 및 DB 동기화 검토 중...")
            
            db_sheet_name = "DB"
            db_ws = None
            for sheet in self.target_wb.Worksheets:
                if sheet.Name == db_sheet_name:
                    db_ws = sheet
                    break
            
            if not db_ws:
                db_ws = self.target_wb.Worksheets.Add(After=self.target_wb.Worksheets(self.target_wb.Worksheets.Count))
                db_ws.Name = db_sheet_name
                db_cols = ["Key", "Manufacturer", "Catalog No.", "Product Name", "CAS No.", "Storage Temp.", "Sensitivity", "Signal Word", "Key Hazards", "Detailed Hazard Classification", "Detail_Link", "SDS_Link", "SDS_Local_Path", "Revision Date"]
                for c_idx, c_name in enumerate(db_cols, 1):
                    db_ws.Cells(1, c_idx).Value = c_name
            
            db_last_row = db_ws.Cells(db_ws.Rows.Count, 1).End(-4162).Row
            
            db_cols_idx = {}
            for c in range(1, 20):
                val = db_ws.Cells(1, c).Value
                if val: db_cols_idx[str(val).strip()] = c
                
            from core.db_manager import DBManager
            
            def get_tc(names):
                return SyncEngine.get_col_idx(names, tgt_col_map)

            def get_db_col(names):
                return SyncEngine.get_col_idx(names, db_cols_idx)

            if "Key" not in db_cols_idx:
                db_ws.Columns(1).Insert()
                db_ws.Cells(1, 1).Value = "Key"
                db_cols_idx = {}
                for c in range(1, 20):
                    val = db_ws.Cells(1, c).Value
                    if val: db_cols_idx[str(val).strip()] = c

            m_c = get_db_col(["Manufacturer", "제조사"])
            cat_c = get_db_col(["Catalog No.", "제품번호"])
            rev_c = get_db_col(["Revision Date", "갱신일"])

            for r in range(2, db_last_row + 1):
                db_ws.Cells(r, 1).Formula = f'=B{r}&"|"&C{r}'
                db_ws.Cells(r, 3).NumberFormat = "@"
                if rev_c > 0:
                    r_val = str(db_ws.Cells(r, rev_c).Value or "").strip()
                    if " " in r_val:
                        db_ws.Cells(r, rev_c).Value = r_val.split(" ")[0]

            for tgt_r in range(2, tgt_last_row + 1):
                mfr_col = get_tc(["Manufacturer", "제조사", "회사"])
                cat_col = get_tc(["Catalog No.", "제품번호", "품번", "카탈로그 번호"])
                name_col = get_tc(["Product Name", "시약명", "품목명", "제품명"])
                
                raw_m = tgt_ws.Cells(tgt_r, mfr_col).Value if mfr_col > 0 else ""
                product_num = tgt_ws.Cells(tgt_r, cat_col).Value if cat_col > 0 else ""
                
                manufacturer = DBManager.normalize_manufacturer(raw_m)
                if manufacturer and mfr_col > 0:
                    self.safe_set_val(tgt_ws.Cells(tgt_r, mfr_col), manufacturer)
                    
                product_num = str(product_num).strip() if product_num and product_num != "None" else ""
                if product_num.endswith(".0"):
                    product_num = product_num[:-2]
                    
                if cat_col > 0 and product_num:
                    tgt_ws.Cells(tgt_r, cat_col).NumberFormat = "@"
                    self.safe_set_val(tgt_ws.Cells(tgt_r, cat_col), product_num)
                    
                chem_name_fallback = tgt_ws.Cells(tgt_r, name_col).Value if name_col > 0 else "Unknown"
                chem_name_fallback = str(chem_name_fallback).strip() if chem_name_fallback and chem_name_fallback != "None" else "Unknown"
                
                if manufacturer and product_num:
                    db_result = None
                    existing_db_row = 0
                    k_db_col = get_db_col(["Key"])
                    m_db_col = get_db_col(["Manufacturer", "제조사"])
                    c_db_col = get_db_col(["Catalog No.", "제품번호"])
                    
                    target_key = f"{manufacturer}|{product_num}"
                    
                    for r in range(2, db_last_row + 1):
                        r_key = str(db_ws.Cells(r, k_db_col if k_db_col > 0 else 1).Value).strip()
                        r_man = DBManager.normalize_manufacturer(db_ws.Cells(r, m_db_col if m_db_col > 0 else 2).Value)
                        r_num = str(db_ws.Cells(r, c_db_col if c_db_col > 0 else 3).Value).strip()
                        if r_num.endswith(".0"): r_num = r_num[:-2]
                        
                        if (r_key and r_key.lower() == target_key.lower()) or (r_man.lower() == manufacturer.lower() and r_num == product_num):
                            existing_db_row = r
                            db_result = {}
                            for k, idx in db_cols_idx.items():
                                db_result[k] = db_ws.Cells(r, idx).Value
                            break
                            
                    if not db_result:
                        man_lower = manufacturer.lower()
                        is_crawlable = any(kw in man_lower for kw in ["thermo", "alfa", "fisher", "invitrogen", "acros", "tci", "tokyo", "sigma", "aldrich", "millipore", "abcam", "앱캠", "티씨아이", "써모", "알드리치", "머크"])
                        
                        if not is_crawlable:
                            self.log(f"  [미지원 제조사 스킵] {manufacturer} - {product_num}")
                            crawled_data = {"error": "Manual Entry Required"}
                        else:
                            self.log(f"  [크롤링 시작] {manufacturer} - {product_num}")
                            if not hasattr(self, 'sb_context_manager') or not self.sb_context_manager:
                                from seleniumbase import SB
                                is_headless = self.config.get("headless", True)
                                self.sb_context_manager = SB(uc=True, headless=is_headless)
                                self.sb = self.sb_context_manager.__enter__()
                            
                            crawled_data = None
                            try:
                                tgt_folder = os.path.dirname(self.target_file)
                                if "thermo" in man_lower or "alfa" in man_lower or "fisher" in man_lower or "invitrogen" in man_lower or "acros" in man_lower:
                                    crawled_data = ThermofisherScraper(browser_context=self.sb, fast_mode=fast_mode, base_dir=tgt_folder).scrape(product_num)
                                elif "tci" in man_lower or "tokyo" in man_lower:
                                    crawled_data = TciScraper(browser_context=self.sb, fast_mode=fast_mode, base_dir=tgt_folder).scrape(product_num)
                                elif "sigma" in man_lower or "aldrich" in man_lower or "millipore" in man_lower:
                                    crawled_data = AldrichScraper(browser_context=self.sb, fast_mode=fast_mode, base_dir=tgt_folder).scrape(product_num)
                                elif "abcam" in man_lower:
                                    crawled_data = AbcamScraper(browser_context=self.sb, fast_mode=fast_mode, base_dir=tgt_folder).scrape(product_num)
                                else:
                                    crawled_data = {"error": "Manual Entry Required"}
                            except Exception as e:
                                crawled_data = {"error": f"Scraping error: {e}"}
                        
                        if crawled_data:
                            if "error" in crawled_data:
                                is_man_req = crawled_data["error"] in ["Manual Entry Required", "Manual Input Required"]
                                db_result = {
                                    "Manufacturer": manufacturer,
                                    "Catalog No.": product_num,
                                    "Product Name": chem_name_fallback,
                                    "CAS No.": "Manual Input Required" if is_man_req else "Search Failed",
                                    "Storage Temp.": "-",
                                    "Signal Word": "-",
                                    "Key Hazards": "-",
                                    "Detailed Hazard Classification": "-",
                                    "Sensitivity": "-",
                                    "Detail_Link": "-" if is_man_req else "Product Not Found",
                                    "SDS_Link": "-",
                                    "SDS_Local_Path": "-",
                                    "Revision Date": "-"
                                }
                            else:
                                db_result = crawled_data
                                db_result["Manufacturer"] = manufacturer
                                db_result["Catalog No."] = product_num
                                
                                p_name = db_result.get("Product Name", db_result.get("시약명", ""))
                                if p_name in ["제조사 홈페이지에서 검색 실패", "검색 실패", "Search Failed", "Product Not Found", "", None] or str(p_name).strip() == "":
                                    db_result["Product Name"] = chem_name_fallback
                                    db_result["CAS No."] = "Search Failed"
                                    db_result["Detail_Link"] = "Product Not Found"
                                    db_result["Storage Temp."] = "-"
                                    db_result["Sensitivity"] = "-"
                                    db_result["Signal Word"] = "-"
                                    db_result["Key Hazards"] = "-"
                                    db_result["Detailed Hazard Classification"] = "-"
                                    db_result["SDS_Link"] = "-"
                                    db_result["SDS_Local_Path"] = "-"
                                    db_result["Revision Date"] = "-"
                                else:
                                    if db_result.get("CAS No.", db_result.get("CAS Number", "")) in ["정보 없음", "", "nan", "None", "N/A", "-"]:
                                        db_result["CAS No."] = "N/A"
                                        
                                    norm_temp = DBManager.normalize_temperature(db_result.get("Storage Temp.", db_result.get("보관온도", "")))
                                    db_result["Storage Temp."] = norm_temp
                                    
                                    det_haz = db_result.get("Detailed Hazard Classification", db_result.get("상세 위험분류", ""))
                                    sens = DBManager.extract_sensitivity(det_haz)
                                    db_result["Sensitivity"] = sens if sens else db_result.get("Sensitivity", "-")
                                    if not db_result.get("Signal Word"): db_result["Signal Word"] = "-"
                                    if not db_result.get("Key Hazards"): db_result["Key Hazards"] = "-"
                                    if not db_result.get("Detailed Hazard Classification"): db_result["Detailed Hazard Classification"] = "-"
                                    if not db_result.get("Detail_Link"): db_result["Detail_Link"] = "-"
                                    if not db_result.get("SDS_Link"): db_result["SDS_Link"] = "-"
                                    if not db_result.get("SDS_Local_Path"): db_result["SDS_Local_Path"] = "-"
                                    db_result["Revision Date"] = datetime.datetime.now().strftime("%Y-%m-%d")
                             
                            # Deduplicated DB writing
                            if existing_db_row > 0:
                                target_db_row = existing_db_row
                            else:
                                db_last_row += 1
                                target_db_row = db_last_row
                                
                            max_c = max(db_cols_idx.values()) if db_cols_idx else 14
                            
                            # Write Catalog No. as text format
                            c_cat_idx = db_cols_idx.get("Catalog No.", 3)
                            db_ws.Cells(target_db_row, c_cat_idx).NumberFormat = "@"
                            
                            for k, v in db_result.items():
                                if k == "Key": continue
                                c_idx = db_cols_idx.get(k)
                                if c_idx and c_idx <= max_c:
                                    self.safe_set_val(db_ws.Cells(target_db_row, c_idx), v)
                                    
                            # Inject Key Excel Formula in Col A
                            db_ws.Cells(target_db_row, 1).Formula = f'=B{target_db_row}&"|"&C{target_db_row}'
                            cnt_db_upd += 1
                        else:
                            db_result = {
                                "Manufacturer": manufacturer, "Catalog No.": product_num, "Product Name": chem_name_fallback,
                                "CAS No.": "Search Failed", "Detail_Link": "Product Not Found",
                                "Storage Temp.": "-", "Signal Word": "-", "Key Hazards": "-", "Detailed Hazard Classification": "-",
                                "Sensitivity": "-", "SDS_Link": "-", "SDS_Local_Path": "-", "Revision Date": "-"
                            }
                            if existing_db_row > 0:
                                target_db_row = existing_db_row
                            else:
                                db_last_row += 1
                                target_db_row = db_last_row
                                
                            max_c = max(db_cols_idx.values()) if db_cols_idx else 14
                            c_cat_idx = db_cols_idx.get("Catalog No.", 3)
                            db_ws.Cells(target_db_row, c_cat_idx).NumberFormat = "@"
                            
                            for k, v in db_result.items():
                                if k == "Key": continue
                                c_idx = db_cols_idx.get(k)
                                if c_idx and c_idx <= max_c:
                                    self.safe_set_val(db_ws.Cells(target_db_row, c_idx), v)
                                    
                            db_ws.Cells(target_db_row, 1).Formula = f'=B{target_db_row}&"|"&C{target_db_row}'
                            cnt_db_upd += 1
                            
                    if db_result:
                        from openpyxl.utils import get_column_letter
                        cat_col_let = get_column_letter(cat_col) if cat_col > 0 else "H"
                        mfr_col_let = get_column_letter(mfr_col) if mfr_col > 0 else "E"

                        lookup_map = [
                            # Storage Temp. is handled separately by apply_formatting_and_validation (VLOOKUP formula over whole column)
                            (["Sensitivity", "민감성"], ["Sensitivity", "민감성"]),
                            (["Signal Word", "신호어"], ["Signal Word", "신호어"]),
                            (["Key Hazards", "주요위험"], ["Key Hazards", "Key Hazards", "주요위험", "Major Hazard"]),
                            (["Detailed Hazard Classification", "상세 위험분류"], ["Detailed Hazard Classification", "Detailed Hazard Classification", "상세 위험분류", "Hazard Statement"]),
                            (["Detail_Link", "상세정보_링크"], ["Detail_Link", "Detail_Link", "상세정보_링크", "Product Link"]),
                            (["SDS_Link"], ["SDS_Link", "SDS Link"]),
                            (["SDS_Local_Path"], ["SDS_Local_Path", "SDS Path"])
                        ]
                        
                        for db_keys, header_names in lookup_map:
                            t_c = get_tc(header_names)
                            if t_c > 0:
                                db_c = 0
                                for k in db_keys:
                                    if k in db_cols_idx:
                                        db_c = db_cols_idx[k]
                                        break
                                if db_c >= 1:
                                    formula = f'=IF(OR({mfr_col_let}{tgt_r}="", {cat_col_let}{tgt_r}=""), "", IFERROR(IF(VLOOKUP({mfr_col_let}{tgt_r}&"|"&{cat_col_let}{tgt_r}&"", DB!$A:$N, {db_c}, FALSE)="","-",VLOOKUP({mfr_col_let}{tgt_r}&"|"&{cat_col_let}{tgt_r}&"", DB!$A:$N, {db_c}, FALSE)), "-"))'
                                    self.safe_set_val(tgt_ws.Cells(tgt_r, t_c), formula)

                        pname_tc = get_tc(["Product Name", "시약명", "품명", "제품명"])
                        orig_pn_tc = get_tc(["Original Product Name"])
                        if pname_tc > 0 and orig_pn_tc > 0:
                            orig_val = tgt_ws.Cells(tgt_r, orig_pn_tc).Value
                            if not orig_val or str(orig_val).strip() in ["", "None"]:
                                curr_p = tgt_ws.Cells(tgt_r, pname_tc).Value
                                fallback_str = str(curr_p).strip() if curr_p and not str(curr_p).startswith("=") else chem_name_fallback
                                tgt_ws.Cells(tgt_r, orig_pn_tc).Value = fallback_str
                            orig_let = get_column_letter(orig_pn_tc)
                            db_pn_c = db_cols_idx.get("Product Name", 4)
                            pn_formula = f'=IF(OR({mfr_col_let}{tgt_r}="", {cat_col_let}{tgt_r}=""), {orig_let}{tgt_r}, IFERROR(IF(VLOOKUP({mfr_col_let}{tgt_r}&"|"&{cat_col_let}{tgt_r}&"", DB!$A:$N, {db_pn_c}, FALSE)="", {orig_let}{tgt_r}, VLOOKUP({mfr_col_let}{tgt_r}&"|"&{cat_col_let}{tgt_r}&"", DB!$A:$N, {db_pn_c}, FALSE)), {orig_let}{tgt_r}))'
                            self.safe_set_val(tgt_ws.Cells(tgt_r, pname_tc), pn_formula)

                        cas_tc = get_tc(["CAS No.", "CAS Number", "CAS 번호"])
                        if cas_tc > 0:
                            curr = tgt_ws.Cells(tgt_r, cas_tc).Value
                            curr_str = str(curr).strip() if curr is not None else ""
                            if not curr_str or curr_str in ["None", "nan", "N/A", "Search Failed", "Manual Input Required", "Manual Entry Required"]:
                                db_cas_c = db_cols_idx.get("CAS No.", 5)
                                formula = f'=IF(OR({mfr_col_let}{tgt_r}="", {cat_col_let}{tgt_r}=""), "", IFERROR(VLOOKUP({mfr_col_let}{tgt_r}&"|"&{cat_col_let}{tgt_r}&"", DB!$A:$N, {db_cas_c}, FALSE), ""))'
                                self.safe_set_val(tgt_ws.Cells(tgt_r, cas_tc), formula)

            # If no new items were added and no items were updated, do NOT perform formatting, completion coloring, or saving
            if cnt_new == 0 and cnt_upd == 0 and cnt_db_upd == 0:
                self.log("동기화할 새로운 내용이 없습니다 (기존 파일 유지).")
                self.cleanup()
                return {
                    "success": True,
                    "new": 0,
                    "updated": 0,
                    "duplicate": cnt_dup,
                    "no_changes": True
                }

            # Color Source Rows
            if len(src_done_rows) > 0 and not self.source_is_readonly:
                self.log("원본 시트 완료 표시 중...")
                c_dict = self.config.get("colors", {})
                done_bg = hex_to_bgr(c_dict.get("done_bg", "#e6f7e6"))
                done_fg = hex_to_bgr(c_dict.get("done_font", "#000000"))
                for r in src_done_rows:
                    if receipt_src_col > 0:
                        r_cell = src_ws.Cells(r, receipt_src_col)
                        r_cell.Interior.Color = done_bg
                        r_cell.Font.Color = done_fg
                try:
                    self.source_wb.Save()
                except Exception as e:
                    self.log(f"원본 시트 저장 중 오류 (무시됨): {e}")

            # Apply Conditional Formatting & Validation on Target
            try:
                self.apply_formatting_and_validation(tgt_ws, tgt_col_map, max(tgt_last_row, 1000))
                sheet_names = [s.Name for s in self.target_wb.Worksheets]
                if "DB" in sheet_names:
                    db_ws = self.target_wb.Worksheets("DB")
                    db_last_r = db_ws.Cells(db_ws.Rows.Count, 1).End(-4162).Row
                    if db_last_r >= 2:
                        self.apply_db_formatting(db_ws, db_last_r)
            except Exception as e:
                self.log(f"서식 주입 중 오류 (무시됨): {e}")
            
            # --- Insert Guide Sheet ---
            try:
                guide_sheet_name = "Guide"
                sheet_names = [s.Name for s in self.target_wb.Worksheets]
                if "가이드(Guide)" in sheet_names and "Guide" not in sheet_names:
                    try:
                        self.target_wb.Worksheets("가이드(Guide)").Name = "Guide"
                    except:
                        pass
                    sheet_names = [s.Name for s in self.target_wb.Worksheets]

                if guide_sheet_name not in sheet_names:
                    guide_ws = self.target_wb.Worksheets.Add(After=self.target_wb.Worksheets(self.target_wb.Worksheets.Count))
                    guide_ws.Name = guide_sheet_name
                    
                    guide_text_rows = [
                        ["[ Chemical Manager 프로그램 사용 방법 및 주의사항 ]"],
                        [""],
                        ["1. 기본 동작 원리 및 저장 기준"],
                        ["- 프로그램은 원본 Orderbook.xlsx (주문대장) 파일의 변경을 감지하여 데이터를 읽어옵니다."],
                        ["- 폴더 내에서 이름에 적힌 날짜/시간이 가장 최신인 ChemicalList 파일을 기준으로 삼습니다."],
                        ["- 업데이트 완료 시 덮어쓰기 대신 새로운 날짜/시간이 적힌 새 파일(ChemicalList_YYYYMMDD_HHMMSS)을 생성합니다."],
                        ["- 기존의 낡은 파일들은 안전하게 'Old Chemical List' 폴더로 자동 이동됩니다."],
                        [""],
                        ["2. 사용자 작성 시 유의사항 (Orderbook 및 Chemical List)"],
                        ["- 빈 줄 금지: 데이터 중간에 완전히 비어있는 줄이 있으면 데이터를 읽다가 중단될 수 있으므로 차례대로 기입하세요."],
                        ["- 동기화 기준: 시약명과 카탈로그 번호 등을 기준으로 기존에 등록된 시약인지 새 시약인지 판단합니다."],
                        ["- 수령 확인 (Status): 수령 후 Status 열에 'O' 또는 'ㅇ'을 입력하면 수령으로 간주되며, PDF 출력 필터링에 사용됩니다."],
                        ["- 드롭다운 선택: 캐비넷(Cabinet)은 Room과 Storage Temp.에 따라 동적으로 변하므로 잘못된 값을 억지로 쓰지 마세요."],
                        [""],
                        ["3. 프로그램 사용 주의사항"],
                        ["- 열린 파일 처리: Chemical List 엑셀 파일이 켜져 있어도 프로그램이 알아서 우회하여 새 파일을 생성합니다."],
                        ["- 단, 원본인 Orderbook 파일은 엑셀에서 저장을 완료해야만 프로그램이 변경 사항을 정확히 감지할 수 있습니다."],
                        ["- 자동 동기화 켜짐 상태에서는 Orderbook을 저장할 때마다 병합이 진행되므로, 수동 제어를 원하시면 정지 버튼을 누르세요."]
                    ]
                    
                    for r, row_data in enumerate(guide_text_rows, 1):
                        guide_ws.Cells(r, 1).Value = row_data[0]
                        if r in [3, 9, 15]: # Subtitles
                            guide_ws.Cells(r, 1).Font.Bold = True
                            
                    title_rng = guide_ws.Range("A1:G1")
                    title_rng.Merge()
                    title_rng.Font.Size = 14
                    title_rng.Font.Bold = True
                    title_rng.Interior.Color = 14211288
                    title_rng.HorizontalAlignment = -4108
                    
                    guide_ws.Columns(1).ColumnWidth = 100
            except Exception as e:
                self.log(f"가이드 시트 생성 중 오류 (무시됨): {e}")

            # 엑셀 파일 열 때 항상 ChemicalList 시트가 보이도록 활성화
            tgt_ws.Activate()

            # 백업 로직 (실제 저장할 때만)
            if os.path.exists(target_path):
                backup_dir = os.path.join(src_folder, "backup")
                if not os.path.exists(backup_dir):
                    os.makedirs(backup_dir)
                now_str = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
                backup_path = os.path.join(backup_dir, f"ChemicalList_backup_{now_str}.xlsx")
                try:
                    shutil.copy2(target_path, backup_path)
                    self.log(f"안전한 진행을 위해 기존 파일을 백업했습니다: {os.path.basename(backup_path)}")
                except Exception as e:
                    self.log(f"백업 실패 (무시됨): {e}")

            # Save and Close Target
            if self.target_wb.ReadOnly:
                now_str = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
                target_path_new = os.path.join(src_folder, f"ChemicalList_{now_str}.xlsx")
                self.log(f"파일이 사용 중이므로 새 파일로 저장합니다: {os.path.basename(target_path_new)}")
                self.target_wb.SaveAs(target_path_new)
                saved_path = target_path_new
            else:
                self.log(f"저장 중: {os.path.basename(target_path)}")
                self.target_wb.SaveAs(target_path)
                saved_path = target_path

            self.cleanup()

            return {
                "success": True,
                "new": cnt_new,
                "updated": cnt_upd,
                "duplicate": cnt_dup
            }

        except Exception as e:
            import traceback
            traceback.print_exc()
            return {"success": False, "error": str(e)}
        finally:
            self.cleanup()

    def apply_formatting_and_validation(self, ws, tgt_col_map, max_row):
        self.log("조건부 서식 및 데이터 유효성 검사 주입 중...")
        c_dict = self.config.get("colors", {})
        
        # 1. Conditional Formatting for missing or '-' values
        try:
            prod_col = tgt_col_map.get("Product Name", 0)
            prod_letter = get_col_letter(prod_col) if prod_col > 0 else "D"
            
            def apply_cf(col_name, bg_key, fg_key):
                col_idx = tgt_col_map.get(col_name, 0)
                if col_idx > 0:
                    try:
                        letter = get_col_letter(col_idx)
                        rng = ws.Range(f"{letter}2:{letter}{max_row}")
                        rng.FormatConditions.Delete()
                        fc = rng.FormatConditions.Add(2, None, f'=AND(OR(TRIM({letter}2)="", TRIM({letter}2)="-"), TRIM({prod_letter}2)<>"")')
                        fc.Interior.Color = hex_to_bgr(c_dict.get(bg_key, "#ffffff"))
                        fc.Font.Color = hex_to_bgr(c_dict.get(fg_key, "#000000"))
                    except Exception as e:
                        pass

            apply_cf("CAS No.", "warn_cas_bg", "warn_cas_font")
            apply_cf("Catalog No.", "warn_num_bg", "warn_num_font")
            apply_cf("Room", "warn_loc_bg", "warn_loc_font")
            apply_cf("Storage Temp.", "warn_loc_bg", "warn_loc_font")
            apply_cf("Cabinet", "warn_loc_bg", "warn_loc_font")

            # Conditional Formatting for Search Failed & Manual Input Required
            try:
                max_c = max(tgt_col_map.values()) if tgt_col_map else 20
                max_c_let = get_col_letter(max_c)
                rng_all = ws.Range(f"A2:{max_c_let}{max_row}")

                sf_bg = hex_to_bgr(c_dict.get("search_failed_bg", "#ffff00"))
                sf_fg = hex_to_bgr(c_dict.get("search_failed_font", "#ff0000"))
                mi_bg = hex_to_bgr(c_dict.get("manual_input_bg", "#ffff00"))
                mi_fg = hex_to_bgr(c_dict.get("manual_input_font", "#ff0000"))

                fc_sf = rng_all.FormatConditions.Add(1, 3, "=\"Search Failed\"")
                fc_sf.Interior.Color = sf_bg
                fc_sf.Font.Color = sf_fg
                fc_sf.Font.Bold = True

                fc_pnf = rng_all.FormatConditions.Add(1, 3, "=\"Product Not Found\"")
                fc_pnf.Interior.Color = sf_bg
                fc_pnf.Font.Color = sf_fg
                fc_pnf.Font.Bold = True

                fc_mi = rng_all.FormatConditions.Add(1, 3, "=\"Manual Input Required\"")
                fc_mi.Interior.Color = mi_bg
                fc_mi.Font.Color = mi_fg
                fc_mi.Font.Bold = True

                fc_me = rng_all.FormatConditions.Add(1, 3, "=\"Manual Entry Required\"")
                fc_me.Interior.Color = mi_bg
                fc_me.Font.Color = mi_fg
                fc_me.Font.Bold = True
            except Exception as e:
                pass
        except Exception as e:
            self.log(f"조건부 서식 주입 중 경고 (무시됨): {e}")

        # 2. Data Validation for Room, Storage Temp., and Cabinet
        try:
            room_col = tgt_col_map.get("Room", 0)
            stemp_col = tgt_col_map.get("Storage Temp.", 0)
            cab_col = tgt_col_map.get("Cabinet", 0)
            mfr_col = tgt_col_map.get("Manufacturer", 0)
            cat_col = tgt_col_map.get("Catalog No.", 0)
            
            if room_col > 0:
                room_l = get_col_letter(room_col)
                rng_r = ws.Range(f"{room_l}2:{room_l}{max_row}")
                try:
                    rng_r.Validation.Delete()
                    rng_r.Validation.Add(3, 1, 1, "='index'!$A$2:$A$4")
                except: pass
                
            if stemp_col > 0:
                stemp_l = get_col_letter(stemp_col)
                rng_s = ws.Range(f"{stemp_l}2:{stemp_l}{max_row}")
                try:
                    rng_s.Validation.Delete()
                    rng_s.Validation.Add(3, 1, 1, "='index'!$C$2:$C$5")
                except: pass

                if mfr_col > 0 and cat_col > 0:
                    mfr_l = get_col_letter(mfr_col)
                    cat_l = get_col_letter(cat_col)
                    # Dynamically determine Storage Temp column in DB sheet from headers
                    try:
                        from utils.excel_utils import get_col_letter as gcl
                        db_ws_tmp = None
                        for _s in ws.Parent.Worksheets:
                            if _s.Name == "DB":
                                db_ws_tmp = _s
                                break
                        if db_ws_tmp:
                            db_stemp_c = next(
                                (c for c in range(1, 20)
                                 if str(db_ws_tmp.Cells(1, c).Value or "").strip() == "Storage Temp."),
                                6
                            )
                        else:
                            db_stemp_c = 6
                    except:
                        db_stemp_c = 6
                    try:
                        vals_mfr = ws.Range(f"{mfr_l}2:{mfr_l}{max_row}").Value
                        vals_cat = ws.Range(f"{cat_l}2:{cat_l}{max_row}").Value
                        existing = rng_s.Formula
                        
                        if not isinstance(vals_mfr, tuple):
                            vals_mfr = ((vals_mfr,),)
                            vals_cat = ((vals_cat,),)
                            existing = ((existing,),)
                            
                        new_formulas = []
                        for i in range(len(vals_mfr)):
                            m = vals_mfr[i][0]
                            c = vals_cat[i][0]
                            e = existing[i][0]
                            if m and str(m).strip() and c and str(c).strip():
                                r = i + 2
                                f = f'=IFERROR(IF(OR(VLOOKUP({mfr_l}{r}&"|"&{cat_l}{r}&"", DB!$A:$N, {db_stemp_c}, FALSE)="", VLOOKUP({mfr_l}{r}&"|"&{cat_l}{r}&"", DB!$A:$N, {db_stemp_c}, FALSE)="-"), "", VLOOKUP({mfr_l}{r}&"|"&{cat_l}{r}&"", DB!$A:$N, {db_stemp_c}, FALSE)), "")'
                                new_formulas.append((f,))
                            else:
                                new_formulas.append((e,))
                                
                        if len(new_formulas) == 1:
                            rng_s.Formula = new_formulas[0][0]
                        elif len(new_formulas) > 1:
                            rng_s.Formula = new_formulas
                    except: pass
            if cab_col > 0 and room_col > 0 and stemp_col > 0:
                cab_l = get_col_letter(cab_col)
                room_l = get_col_letter(room_col)
                stemp_l = get_col_letter(stemp_col)
                rng_c = ws.Range(f"{cab_l}2:{cab_l}{max_row}")
                try:
                    rng_c.Validation.Delete()
                    formula = f'=IF(COUNTIF(index!$H:$H, {room_l}2&"_"&{stemp_l}2)=0, index!$G$100, OFFSET(index!$G$1, MATCH({room_l}2&"_"&{stemp_l}2, index!$H:$H, 0)-1, 0, COUNTIF(index!$H:$H, {room_l}2&"_"&{stemp_l}2), 1))'
                    rng_c.Validation.Add(3, 1, 1, formula)
                except: pass
        except Exception as e:
            self.log(f"유효성 검사 주입 중 경고 (무시됨): {e}")

        # 2. Apply Status Formula & Coloring
        try:
            stat_col = tgt_col_map.get("Status", 0)
            qty_col = tgt_col_map.get("Quantity", 0)
            used_col = tgt_col_map.get("Used", 0)
            
            if stat_col > 0 and qty_col > 0 and used_col > 0:
                stat_l = get_col_letter(stat_col)
                qty_l = get_col_letter(qty_col)
                used_l = get_col_letter(used_col)
                prod_col = tgt_col_map.get("Product Name", 0)
                if prod_col > 0:
                    prod_l = get_col_letter(prod_col)
                    formula = f'=IF(TRIM({prod_l}2)="", "", IF(N({qty_l}2)-N({used_l}2)<=0, "X", N({qty_l}2)-N({used_l}2)))'
                else:
                    formula = f'=IF(N({qty_l}2)-N({used_l}2)<=0, "X", N({qty_l}2)-N({used_l}2))'
                
                rng = ws.Range(f"{stat_l}2:{stat_l}{max_row}")
                rng.Formula = formula
                rng.HorizontalAlignment = -4108 # xlCenter
                
                # Status "X" coloring
                try:
                    rng.FormatConditions.Delete()
                    fc_x = rng.FormatConditions.Add(1, 3, "=\"X\"") # xlCellValue, xlEqual
                    fc_x.Interior.Color = hex_to_bgr(c_dict.get("status_x_bg", "#ffe6e6"))
                    fc_x.Font.Color = hex_to_bgr(c_dict.get("status_x_font", "#ff0000"))
                    fc_x.Font.Bold = True
                except Exception as e:
                    pass
        except Exception as e:
            self.log(f"Status 수식 적용 중 경고 (무시됨): {e}")
            
        # 3. Symbol Colors for Signal Word and Key Hazards
        try:
            sig_col = SyncEngine.get_col_idx(["Signal Word", "신호어"], tgt_col_map)
            haz_col = SyncEngine.get_col_idx(["Key Hazards", "주요위험", "Major Hazard"], tgt_col_map)
            self.apply_symbol_colors(ws, max_row, sig_col)
            self.apply_symbol_colors(ws, max_row, haz_col)
        except Exception as e:
            self.log(f"기호 색상 적용 중 경고 (무시됨): {e}")

        # 4. Vertical & Horizontal Alignments
        try:
            center_cols = {
                "Order No.", "Order Date", "Ordered By", "Package Size", 
                "CAS No.", "Catalog No.", "Room", "Storage Temp.", "Cabinet", 
                "Quantity", "Used", "Status"
            }
            for col_name, c_idx in tgt_col_map.items():
                if c_idx > 0:
                    letter = get_col_letter(c_idx)
                    col_rng = ws.Range(f"{letter}2:{letter}{max_row}")
                    col_rng.VerticalAlignment = -4108 # xlVCenter (상하 가운데)
                    if col_name in center_cols:
                        col_rng.HorizontalAlignment = -4108 # xlHAlignCenter (좌우 가운데)
                    else:
                        col_rng.HorizontalAlignment = -4131 # xlHAlignLeft (좌우 왼쪽)

            max_c_idx = max(tgt_col_map.values()) if tgt_col_map else 20
            hdr_let = get_col_letter(max_c_idx)
            header_rng = ws.Range(f"A1:{hdr_let}1")
            header_rng.VerticalAlignment = -4108 # xlVCenter
            header_rng.HorizontalAlignment = -4108 # xlHAlignCenter
        except Exception as e:
            self.log(f"정렬 설정 중 경고 (무시됨): {e}")

    def apply_symbol_colors(self, ws, max_row, col_idx):
        if col_idx <= 0: return
        symbol_colors = {
            '●': 255,             # Red (FF0000)
            '▲': 26367,           # Orange (FF6600)
            '■': 39168,           # Green (009900)
            '◆': 16711680,        # Blue (0000FF)
            '★': 255,             # Red
            '▼': 8388736,         # Purple (800080)
        }
        for r in range(2, max_row + 1):
            cell = ws.Cells(r, col_idx)
            val = cell.Value
            if val and isinstance(val, str):
                for i, char in enumerate(val):
                    if char in symbol_colors:
                        try:
                            cell.GetCharacters(i+1, 1).Font.Color = symbol_colors[char]
                        except:
                            pass

    def apply_db_formatting(self, db_ws, max_row):
        if not db_ws: return
        c_dict = self.config.get("colors", {})
        try:
            max_c = db_ws.Cells(1, db_ws.Columns.Count).End(-4159).Column
            max_c_let = get_col_letter(max(max_c, 14))

            # 1. Header Row (Row 1): Bold, RGB(226, 239, 218), Center Alignment
            hdr_rng = db_ws.Range(f"A1:{max_c_let}1")
            hdr_rng.Font.Bold = True
            hdr_rng.Interior.Color = hex_to_bgr("#E2EFDA")
            hdr_rng.HorizontalAlignment = -4108 # xlCenter
            hdr_rng.VerticalAlignment = -4108   # xlVCenter

            # 2. Data Rows (Rows 2+): Cols 1..6 Center, Cols 7+ Left Alignment
            if max_row >= 2:
                center_rng = db_ws.Range(f"A2:F{max_row}")
                center_rng.HorizontalAlignment = -4108 # xlCenter
                center_rng.VerticalAlignment = -4108   # xlVCenter

                if max(max_c, 14) >= 7:
                    left_rng = db_ws.Range(f"G2:{max_c_let}{max_row}")
                    left_rng.HorizontalAlignment = -4131 # xlLeft
                    left_rng.VerticalAlignment = -4108   # xlVCenter

                rng = db_ws.Range(f"A2:{max_c_let}{max_row}")

                sf_bg = hex_to_bgr(c_dict.get("search_failed_bg", "#ffff00"))
                sf_fg = hex_to_bgr(c_dict.get("search_failed_font", "#ff0000"))
                mi_bg = hex_to_bgr(c_dict.get("manual_input_bg", "#ffff00"))
                mi_fg = hex_to_bgr(c_dict.get("manual_input_font", "#ff0000"))

                fc_sf = rng.FormatConditions.Add(1, 3, "=\"Search Failed\"")
                fc_sf.Interior.Color = sf_bg
                fc_sf.Font.Color = sf_fg
                fc_sf.Font.Bold = True

                fc_pnf = rng.FormatConditions.Add(1, 3, "=\"Product Not Found\"")
                fc_pnf.Interior.Color = sf_bg
                fc_pnf.Font.Color = sf_fg
                fc_pnf.Font.Bold = True

                fc_mi = rng.FormatConditions.Add(1, 3, "=\"Manual Input Required\"")
                fc_mi.Interior.Color = mi_bg
                fc_mi.Font.Color = mi_fg
                fc_mi.Font.Bold = True

                fc_me = rng.FormatConditions.Add(1, 3, "=\"Manual Entry Required\"")
                fc_me.Interior.Color = mi_bg
                fc_me.Font.Color = mi_fg
                fc_me.Font.Bold = True
        except Exception as e:
            self.log(f"DB 시트 서식 적용 중 경고 (무시됨): {e}")

    def cleanup(self):
        if self.source_wb:
            try:
                self.source_wb.Close(SaveChanges=False)
            except: pass
            self.source_wb = None
        if self.target_wb:
            if self.target_wb != self.source_wb:
                try:
                    self.target_wb.Close(SaveChanges=False)
                except: pass
            self.target_wb = None
        if self.excel:
            try:
                self.excel.Quit()
            except: pass
            
            if hasattr(self, 'excel_pid') and self.excel_pid:
                try:
                    import subprocess
                    subprocess.run(f"taskkill /F /PID {self.excel_pid}", shell=True, capture_output=True)
                except: pass
                
            self.excel = None
        if hasattr(self, 'sb_context_manager') and self.sb_context_manager:
            try:
                self.sb_context_manager.__exit__(None, None, None)
            except: pass
            self.sb_context_manager = None
            self.sb = None

def refresh_chemical_list_formatting(target_path, config):
    """
    Opens target_path (ChemicalList.xlsx) in Win32COM Excel, refreshes formulas,
    Data Validation for Room/Storage Temp/Cabinet, and Conditional Formatting.
    """
    if not target_path or not os.path.exists(target_path):
        return
    import win32com.client
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        wb = excel.Workbooks.Open(target_path, UpdateLinks=False)
        
        tgt_ws = None
        for s in wb.Worksheets:
            if s.Name not in ["DB", "가이드(Guide)", "Guide", "index", "Old Chemical List"]:
                tgt_ws = s
                break
        if not tgt_ws:
            tgt_ws = wb.Sheets(1)
            
        tgt_headers = config.get("target_headers", [])
        tgt_col_map = {}
        tgt_last_col = tgt_ws.Cells(1, tgt_ws.Columns.Count).End(-4159).Column
        for c in range(1, max(tgt_last_col + 1, 30)):
            val = tgt_ws.Cells(1, c).Value
            if val:
                tgt_col_map[str(val).strip()] = c
                
        for idx, name in enumerate(tgt_headers):
            if name not in tgt_col_map:
                tgt_col_map[name] = idx + 1
                
        tgt_last_row = tgt_ws.Cells(tgt_ws.Rows.Count, 1).End(-4162).Row
        max_r = max(tgt_last_row, 1000)
        
        engine = SyncEngine.__new__(SyncEngine)
        engine.config = config
        engine.log = lambda msg: print(f"[RefreshFormatting] {msg}")
        
        engine.apply_formatting_and_validation(tgt_ws, tgt_col_map, max_r)
        
        sheet_names = [s.Name for s in wb.Worksheets]
        if "DB" in sheet_names:
            db_ws = wb.Worksheets("DB")
            db_last_r = db_ws.Cells(db_ws.Rows.Count, 1).End(-4162).Row
            if db_last_r >= 2:
                engine.apply_db_formatting(db_ws, db_last_r)
                
        wb.Save()
        wb.Close(False)
        excel.Quit()
    except Exception as e:
        print(f"[RefreshFormatting Error] {e}")
