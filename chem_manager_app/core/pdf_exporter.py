import os
import win32com.client
from utils.excel_utils import get_col_letter

class PDFExporter:
    def __init__(self, config_data, callback_progress=None):
        self.config = config_data
        self.callback = callback_progress

    def log(self, message):
        if self.callback:
            self.callback(message)
        else:
            print(message)

    def export(self, mode, selected_cols, only_status_o, output_path):
        """
        mode: 1 (By Room), 2 (By Room/Temp/Cabinet)
        selected_cols: list of column headers to include
        only_status_o: bool, if True, only include rows where Status == 'O'
        output_path: Absolute path to save the .pdf file
        """
        import urllib.parse
        import glob
        import re
        src_path_raw = self.config.get("source_file", "")
        if src_path_raw.startswith("file:///"):
            src_path_raw = urllib.parse.unquote(src_path_raw[8:])
        
        if not src_path_raw:
            return {"success": False, "error": "원본 파일 경로가 설정되지 않았습니다."}
            
        src_folder = os.path.dirname(os.path.abspath(src_path_raw))
        
        # Find latest ChemicalList file
        search_pattern = os.path.join(src_folder, "ChemicalList*.xlsx")
        all_files = glob.glob(search_pattern)
        
        valid_files = []
        for f in all_files:
            basename = os.path.basename(f)
            if basename.startswith("~$"): continue
            if re.match(r'^ChemicalList(?:_\d{8}_\d{6})?\.xlsx$', basename):
                valid_files.append(f)
                
        if not valid_files:
            return {"success": False, "error": f"동기화된 ChemicalList 파일을 찾을 수 없습니다. 먼저 동기화를 진행해주세요."}
        
        valid_files.sort(key=lambda x: os.path.basename(x), reverse=True)
        target_path = valid_files[0]
        self.log(f"최신 ChemicalList 파일 읽기: {os.path.basename(target_path)}")

        excel = None
        source_wb = None
        temp_wb = None

        try:
            excel = win32com.client.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.ScreenUpdating = False
            excel.DisplayAlerts = False
            
            self.log("ChemicalList 데이터를 읽어오는 중...")
            source_wb = excel.Workbooks.Open(target_path, ReadOnly=True, UpdateLinks=False)
            ws = source_wb.Sheets(1)
            
            # Read Headers first
            last_col = ws.Cells(1, ws.Columns.Count).End(-4159).Column
            headers = [str(ws.Cells(1, c).Value).strip() for c in range(1, last_col + 1)]
            col_map = {name: c for c, name in enumerate(headers, 1)}
            
            # Determine last row based on Product Name column
            prod_col = col_map.get("Product Name", 0)
            if prod_col > 0:
                last_row = ws.Cells(ws.Rows.Count, prod_col).End(-4162).Row
            else:
                last_row = ws.Cells(ws.Rows.Count, 1).End(-4162).Row
            
            if last_row < 2:
                raise Exception("ChemicalList에 내보낼 데이터가 없습니다.")
            
            # Get grouping column indices
            room_col = col_map.get("Room", 0)
            temp_col = col_map.get("Storage Temp.", 0)
            cab_col = col_map.get("Cabinet", 0)
            status_col = col_map.get("Status", 0)
            prod_col = col_map.get("Product Name", 0)
            
            if mode == 1 and room_col == 0:
                raise Exception("ChemicalList에 'Room' 열을 찾을 수 없습니다.")
            if mode == 2 and (room_col == 0 or temp_col == 0 or cab_col == 0):
                raise Exception("ChemicalList에 'Room', 'Storage Temp.', 'Cabinet' 열을 모두 찾을 수 없습니다.")

            # Filter selected columns
            sel_col_indices = [col_map[c] for c in selected_cols if c in col_map]
            if not sel_col_indices:
                raise Exception("선택된 열이 유효하지 않습니다.")
                
            filtered_headers = [headers[i-1] for i in sel_col_indices]

            # Read all data
            data_range = ws.Range(ws.Cells(2, 1), ws.Cells(last_row, last_col)).Value
            
            if data_range and not isinstance(data_range[0], (tuple, list)):
                data_range = [data_range]
            
            self.log("데이터 그룹화 및 필터링 중...")
            groups = {}
            for row in data_range:
                if not row:
                    continue
                
                if prod_col > 0:
                    prod_val = str(row[prod_col - 1]).strip() if row[prod_col - 1] is not None else ""
                    if not prod_val:
                        continue
                else:
                    if all(c is None or str(c).strip() == "" for c in row):
                        continue
                    
                # Status filter
                if only_status_o and status_col > 0:
                    status_val = str(row[status_col - 1]).strip() if row[status_col - 1] is not None else ""
                    if status_val.upper() == "X":
                        continue
                
                # Extract group key
                if mode == 1:
                    r_val = str(row[room_col - 1]).strip() if room_col > 0 and row[room_col - 1] is not None else "Unknown"
                    key = r_val
                    title_str = f"Chemical List in Room {r_val}"
                    sheet_name = f"Room_{r_val}"[:31] # Excel sheet name max 31 chars
                else:
                    r_val = str(row[room_col - 1]).strip() if room_col > 0 and row[room_col - 1] is not None else "Unknown"
                    t_val = str(row[temp_col - 1]).strip() if temp_col > 0 and row[temp_col - 1] is not None else "Unknown"
                    c_val = str(row[cab_col - 1]).strip() if cab_col > 0 and row[cab_col - 1] is not None else "Unknown"
                    
                    import re
                    c_base = re.sub(r'\d+$', '', c_val).strip()
                    
                    key = f"{r_val}_{t_val}_{c_base}"
                    title_str = f"Chemical List in Room {r_val}, Storage Temp. {t_val}, Cabinet {c_base}"
                    
                    # Make valid sheet name
                    sheet_name = f"{r_val}_{t_val}_{c_base}"
                    # Replace invalid characters
                    for ch in ['\\', '/', '?', '*', '[', ']']:
                        sheet_name = sheet_name.replace(ch, '')
                    sheet_name = sheet_name[:31]

                if sheet_name not in groups:
                    groups[sheet_name] = {"rows": [], "title": title_str}
                groups[sheet_name]["rows"].append(row)
                
            source_wb.Close(False)
            source_wb = None
            
            if not groups:
                return {"success": False, "error": "출력할 데이터가 없습니다 (필터 조건에 맞는 데이터가 없음)."}
            
            self.log(f"임시 워크북 생성 중 (총 {len(groups)}개 그룹)...")
            temp_wb = excel.Workbooks.Add()
            
            # Remove default extra sheets if any
            while temp_wb.Sheets.Count > 1:
                temp_wb.Sheets(temp_wb.Sheets.Count).Delete()
                
            first_sheet = True
            for sheet_name, group_data in groups.items():
                rows = group_data["rows"]
                title_str = group_data["title"]
                if first_sheet:
                    new_ws = temp_wb.Sheets(1)
                    first_sheet = False
                else:
                    new_ws = temp_wb.Sheets.Add(After=temp_wb.Sheets(temp_wb.Sheets.Count))
                
                new_ws.Name = sheet_name
                
                # Write Title on Row 1
                new_ws.Cells(1, 1).Value = title_str
                title_rng = new_ws.Range(new_ws.Cells(1, 1), new_ws.Cells(1, len(filtered_headers)))
                title_rng.Merge()
                title_rng.Font.Size = 14
                title_rng.Font.Bold = True
                title_rng.HorizontalAlignment = -4108 # xlCenter
                title_rng.VerticalAlignment = -4108 # xlCenter
                
                # Write Header on Row 2
                for c, header in enumerate(filtered_headers, 1):
                    new_ws.Cells(2, c).Value = header
                    
                hrange = new_ws.Range(new_ws.Cells(2, 1), new_ws.Cells(2, len(filtered_headers)))
                hrange.Font.Bold = True
                hrange.Interior.Color = 14211288 # #D8D8D8 in BGR (or similar grey)
                hrange.Borders.LineStyle = 1 # xlContinuous
                hrange.HorizontalAlignment = -4108 # xlCenter
                hrange.VerticalAlignment = -4108 # xlCenter
                
                # Write Data on Row 3
                if rows:
                    filtered_rows = [[r[i-1] for i in sel_col_indices] for r in rows]
                    data_rng = new_ws.Range(new_ws.Cells(3, 1), new_ws.Cells(len(rows) + 2, len(filtered_headers)))
                    data_rng.Value = filtered_rows
                    data_rng.Borders.LineStyle = 1
                    data_rng.HorizontalAlignment = -4108 # xlCenter
                    data_rng.VerticalAlignment = -4108 # xlCenter
                
                # Format Columns
                new_ws.Columns.AutoFit()
                
                # PageSetup for PDF
                ps = new_ws.PageSetup
                ps.Orientation = 2 # xlLandscape
                ps.Zoom = False
                ps.FitToPagesWide = 1
                ps.FitToPagesTall = False
                ps.PrintTitleRows = "$2:$2" # Repeat headers only (Row 2), not title (Row 1)
                
            self.log("PDF 파일로 내보내는 중...")
            # Export to PDF (Type 0 = xlTypePDF)
            # Ensure output_path uses Windows backslashes because Excel COM fails with forward slashes
            safe_out_path = os.path.abspath(output_path)
            temp_wb.ExportAsFixedFormat(0, safe_out_path)
            
            self.log("PDF 생성 완료!")
            return {"success": True, "path": output_path}

        except Exception as e:
            import traceback
            traceback.print_exc()
            return {"success": False, "error": str(e)}
        finally:
            if source_wb:
                try: source_wb.Close(False)
                except: pass
            if temp_wb:
                try: temp_wb.Close(False)
                except: pass
            if excel:
                try: excel.Quit()
                except: pass
