import os
import time
import shutil
import datetime
import win32com.client
from utils.color_utils import hex_to_bgr
from utils.excel_utils import get_col_letter

class SyncEngine:
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
        
        import glob
        import re
        search_pattern = os.path.join(src_folder, "ChemicalList*.xlsx")
        all_files = glob.glob(search_pattern)
        
        # Filter strictly for ChemicalList.xlsx or ChemicalList_YYYYMMDD_HHMMSS.xlsx
        # Ignore ~$ files and ChemicalList_old_ files
        valid_files = []
        for f in all_files:
            basename = os.path.basename(f)
            if basename.startswith("~$"): continue
            if re.match(r'^ChemicalList(?:_\d{8}_\d{6})?\.xlsx$', basename):
                valid_files.append(f)
        
        target_path = None
        if valid_files:
            # Sort by filename descending (this automatically puts newest YYYYMMDD_HHMMSS first)
            # ChemicalList_20260728... > ChemicalList_20260727... > ChemicalList.xlsx
            valid_files.sort(key=lambda x: os.path.basename(x), reverse=True)
            target_path = valid_files[0]

        try:
            self.excel = win32com.client.DispatchEx("Excel.Application")
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

            # Build Target Indices
            tgt_headers = self.config["target_headers"]
            tgt_col_map = {name: idx + 1 for idx, name in enumerate(tgt_headers)}
            
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
            src_done_rows = []

            self.log("데이터 동기화 진행 중...")
            
            # Find specific cols for logic
            qty_col = tgt_col_map.get("Quantity", 0)
            used_col = tgt_col_map.get("Used", 0)
            stat_col = tgt_col_map.get("Status", 0)
            
            receipt_src_col = src_col_map.get("수령확인", 0)
            chem_src_col = src_col_map.get("시약", 0)

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
                    continue # "수령확인" 열을 찾을 수 없으면 아무것도 가져오지 않음
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
                        if t_header == "Status": continue # Status is handled by formula
                        t_c = tgt_col_map.get(t_header, 0)
                        s_c = src_col_map.get(s_header, 0)
                        if t_c > 0 and s_c > 0 and s_c - 1 < len(row_data):
                            new_val = row_data[s_c - 1]
                            if new_val is not None:
                                tgt_ws.Cells(tgt_r, t_c).Value = new_val
                    
                    tgt_ws.Range(tgt_ws.Cells(tgt_r, 1), tgt_ws.Cells(tgt_r, len(tgt_headers))).HorizontalAlignment = -4108 # xlCenter
                    cnt_new += 1

                src_done_rows.append(header_row + 1 + i)

            # Apply Status Formula
            if stat_col > 0 and qty_col > 0 and used_col > 0 and tgt_last_row >= 2:
                self.log("Status 수식 적용 중...")
                tgt_ws.Range(tgt_ws.Cells(2, stat_col), tgt_ws.Cells(tgt_last_row, stat_col)).FormulaR1C1 = f"=IF(RC{qty_col}>RC{used_col},\"O\",\"X\")"

            # Color Source Rows
            if len(src_done_rows) > 0 and not self.source_is_readonly:
                self.log("원본 시트 완료 표시 중...")
                c_dict = self.config.get("colors", {})
                done_bg = hex_to_bgr(c_dict.get("done_bg", "#e6f7e6"))
                done_fg = hex_to_bgr(c_dict.get("done_font", "#000000"))
                for r in src_done_rows:
                    if chem_src_col > 0:
                        c_cell = src_ws.Cells(r, chem_src_col)
                        c_cell.Interior.Color = done_bg
                        c_cell.Font.Color = done_fg
                    if receipt_src_col > 0:
                        r_cell = src_ws.Cells(r, receipt_src_col)
                        r_cell.Interior.Color = done_bg
                        r_cell.Font.Color = done_fg
                try:
                    self.source_wb.Save()
                except Exception as e:
                    self.log(f"원본 시트 저장 중 오류 (무시됨): {e}")

            # Apply Conditional Formatting & Validation on Target
            if cnt_new > 0 or cnt_upd > 0:
                self.apply_formatting_and_validation(tgt_ws, tgt_col_map, max(tgt_last_row, 1000))
                
                # --- Insert Guide Sheet ---
                try:
                    guide_sheet_name = "가이드(Guide)"
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

                # Save and Close Target
                now_str = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
                new_target_path = os.path.join(src_folder, f"ChemicalList_{now_str}.xlsx")
                self.log(f"새 파일로 저장 중: {os.path.basename(new_target_path)}")
                self.target_wb.SaveAs(new_target_path)
                saved_path = new_target_path
            else:
                self.log("업데이트된 항목이 없습니다. (저장 생략)")
                saved_path = target_path
                
            # Cleanup Old Files
            old_dir = os.path.join(src_folder, "Old Chemical List")
            if not os.path.exists(old_dir):
                os.makedirs(old_dir)
                
            all_chem_files = glob.glob(search_pattern)
            
            for f in all_chem_files:
                basename = os.path.basename(f)
                if basename.startswith("~$"): continue
                # We only move valid target files (like ChemicalList.xlsx or ChemicalList_YYYYMMDD_HHMMSS.xlsx)
                # We don't move random other files starting with ChemicalList just in case
                if not re.match(r'^ChemicalList(?:_\d{8}_\d{6})?\.xlsx$', basename):
                    continue
                    
                if saved_path and os.path.abspath(f) == os.path.abspath(saved_path):
                    continue
                try:
                    dest = os.path.join(old_dir, os.path.basename(f))
                    if os.path.exists(dest):
                        os.remove(dest) # Remove if already exists in Old to overwrite
                    shutil.move(f, dest)
                except Exception as e:
                    self.log(f"백업 파일 이동 실패 (다음 동기화 시 재시도): {os.path.basename(f)}")
            
            return {
                "success": True,
                "new": cnt_new,
                "updated": cnt_upd,
                "duplicate": cnt_dup
            }

        except Exception as e:
            return {"success": False, "error": str(e)}
        finally:
            self.cleanup()

    def apply_formatting_and_validation(self, ws, tgt_col_map, max_row):
        self.log("조건부 서식 및 데이터 유효성 검사 주입 중...")
        c_dict = self.config.get("colors", {})
        xlExpression = 2
        xlValidateList = 3
        xlValidAlertStop = 1
        
        # 1. Conditional Formatting
        prod_col = tgt_col_map.get("Product Name", 0)
        prod_letter = get_col_letter(prod_col) if prod_col > 0 else "D"
        
        def apply_cf(col_name, bg_key, fg_key):
            col_idx = tgt_col_map.get(col_name, 0)
            if col_idx > 0:
                letter = get_col_letter(col_idx)
                rng = ws.Range(f"{letter}2:{letter}{max_row}")
                rng.FormatConditions.Delete()
                # Type=2, Operator=None, Formula1=...
                fc = rng.FormatConditions.Add(2, None, f'=AND(TRIM({letter}2)="", TRIM({prod_letter}2)<>"")')
                fc.Interior.Color = hex_to_bgr(c_dict.get(bg_key, "#ffffff"))
                fc.Font.Color = hex_to_bgr(c_dict.get(fg_key, "#000000"))

        apply_cf("CAS No.", "warn_cas_bg", "warn_cas_font")
        apply_cf("Catalog No.", "warn_num_bg", "warn_num_font")
        apply_cf("Room", "warn_loc_bg", "warn_loc_font")
        apply_cf("Storage Temp.", "warn_loc_bg", "warn_loc_font")
        apply_cf("Cabinet", "warn_loc_bg", "warn_loc_font")
        
        # --- Apply Status Formula ---
        stat_col = tgt_col_map.get("Status", 0)
        qty_col = tgt_col_map.get("Quantity", 0)
        used_col = tgt_col_map.get("Used", 0)
        
        if stat_col > 0 and qty_col > 0 and used_col > 0:
            stat_l = get_col_letter(stat_col)
            qty_l = get_col_letter(qty_col)
            used_l = get_col_letter(used_col)
            if prod_col > 0:
                prod_l = get_col_letter(prod_col)
                formula = f'=IF(TRIM({prod_l}2)="", "", IF(N({qty_l}2)-N({used_l}2)<=0, "X", N({qty_l}2)-N({used_l}2)))'
            else:
                formula = f'=IF(N({qty_l}2)-N({used_l}2)<=0, "X", N({qty_l}2)-N({used_l}2))'
            
            rng = ws.Range(f"{stat_l}2:{stat_l}{max_row}")
            rng.Formula = formula
            rng.HorizontalAlignment = -4108 # xlCenter
        
        # 2. Data Validation
        sheet_names = [sheet.Name for sheet in self.target_wb.Worksheets]
        if "index" in sheet_names:
            idx_ws = self.target_wb.Worksheets("index")
            # Create Helper column in index sheet (Column H = E & "_" & F)
            idx_ws.Range("H2:H1000").Formula = '=E2&"_"&F2'
            
            room_idx = tgt_col_map.get("Room", 0)
            temp_idx = tgt_col_map.get("Storage Temp.", 0)
            cab_idx = tgt_col_map.get("Cabinet", 0)
            
            if room_idx > 0:
                letter = get_col_letter(room_idx)
                rng = ws.Range(f"{letter}2:{letter}{max_row}")
                rng.Validation.Delete()
                # Type=3, AlertStyle=1, Operator=1 (xlBetween)
                room_formula = "=OFFSET('index'!$A$2, 0, 0, MAX(1, COUNTA('index'!$A:$A)-1), 1)"
                rng.Validation.Add(3, 1, 1, room_formula)
                rng.Validation.InCellDropdown = True
            
            if temp_idx > 0:
                letter = get_col_letter(temp_idx)
                rng = ws.Range(f"{letter}2:{letter}{max_row}")
                rng.Validation.Delete()
                temp_formula = "=OFFSET('index'!$C$2, 0, 0, MAX(1, COUNTA('index'!$C:$C)-1), 1)"
                rng.Validation.Add(3, 1, 1, temp_formula)
                rng.Validation.InCellDropdown = True
                
            if cab_idx > 0 and room_idx > 0 and temp_idx > 0:
                cab_letter = get_col_letter(cab_idx)
                room_letter = get_col_letter(room_idx)
                temp_letter = get_col_letter(temp_idx)
                
                # 확실한 빈 셀 하나를 만들어 둡니다 (조건 불일치 시 참조용)
                idx_ws.Range("ZZ1").ClearContents()
                
                rng = ws.Range(f"{cab_letter}2:{cab_letter}{max_row}")
                rng.Validation.Delete()
                # 일치하는 항목이 0개일 경우 강제로 비어있는 ZZ1 셀을 참조하게 하여 쓰레기값이 안 뜨게 합니다.
                formula = f"=IF(COUNTIF('index'!$H:$H, {room_letter}2&\"_\"&{temp_letter}2)=0, 'index'!$ZZ$1, OFFSET('index'!$G$1, MATCH({room_letter}2&\"_\"&{temp_letter}2, 'index'!$H:$H, 0)-1, 0, COUNTIF('index'!$H:$H, {room_letter}2&\"_\"&{temp_letter}2), 1))"
                rng.Validation.Add(3, 1, 1, formula)
                rng.Validation.InCellDropdown = True

    def cleanup(self):
        if self.source_wb:
            try:
                self.source_wb.Close(SaveChanges=False)
            except: pass
            self.source_wb = None
        if self.target_wb:
            try:
                self.target_wb.Close(SaveChanges=False)
            except: pass
            self.target_wb = None
        if self.excel:
            try:
                self.excel.Quit()
            except: pass
            self.excel = None
