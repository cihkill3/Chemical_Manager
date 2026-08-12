import os
from PyQt6.QtWidgets import (QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
                             QPushButton, QLabel, QLineEdit, QFileDialog, QGroupBox, 
                             QTextEdit, QMessageBox, QSpinBox, QCheckBox, QComboBox, QFormLayout,
                             QDialog, QTableWidget, QTableWidgetItem, QHeaderView, QRadioButton, QButtonGroup)
from PyQt6.QtCore import pyqtSignal, Qt, QThread
from gui.mapping_dialog import MappingDialog
from gui.color_dialog import ColorDialog
from gui.pdf_dialog import PdfDialog
from gui.db_update_dialog import DbUpdateDialog
from core.sync_engine import SyncEngine
from core.pdf_exporter import PDFExporter

class PdfWorker(QThread):
    progress = pyqtSignal(str)
    finished = pyqtSignal(dict)
    
    def __init__(self, config, mode, selected_cols, only_status_o, out_path):
        super().__init__()
        self.config = config
        self.mode = mode
        self.selected_cols = selected_cols
        self.only_status_o = only_status_o
        self.out_path = out_path
        
    def run(self):
        try:
            exporter = PDFExporter(self.config, callback_progress=self.progress.emit)
            result = exporter.export(self.mode, self.selected_cols, self.only_status_o, self.out_path)
            self.finished.emit(result)
        except Exception as e:
            self.finished.emit({"success": False, "error": str(e)})

class SdsWorker(QThread):
    progress = pyqtSignal(str)
    finished = pyqtSignal(dict)
    
    def __init__(self, config, start_date, out_path):
        super().__init__()
        self.config = config
        self.start_date = start_date
        self.out_path = out_path
        
    def run(self):
        try:
            import urllib.parse
            src_path_raw = self.config.get("source_file", "")
            if src_path_raw.startswith("file://"):
                src_path_raw = urllib.parse.unquote(src_path_raw.replace("file:///", "").replace("file://", ""))
            src_folder = os.path.dirname(os.path.abspath(src_path_raw))
            db_path = os.path.join(src_folder, "ChemicalList.xlsx")
            
            exporter = PDFExporter(self.config, callback_progress=self.progress.emit)
            result = exporter.export_sds_batch(db_path, self.start_date, self.out_path)
            self.finished.emit(result)
        except Exception as e:
            self.finished.emit({"success": False, "error": str(e)})

class DBConflictResolverDialog(QDialog):
    def __init__(self, conflicts, parent=None):
        super().__init__(parent)
        self.setWindowTitle("DB 정보 충돌 해결 (수동 업데이트)")
        self.resize(850, 600)
        self.conflicts = conflicts
        self.radio_groups = {} 
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout(self)
        
        header_lbl = QLabel(
            f"<h3>⚠️ DB 수동 업데이트 중 기존 DB 정보와 충돌하는 항목 발견</h3>"
            f"<p>총 <b>{len(self.conflicts)}개 제품</b>에서 기존 DB 정보와 새로운 크롤링 정보가 서로 다릅니다.<br>"
            f"적용할 버전을 선택하신 후 하단의 <b>[선택 항목 반영 및 저장]</b> 버튼을 클릭해 주세요.<br>"
            f"<i>(취소 버튼 클릭 시 엑셀 파일에는 아무것도 변경되지 않습니다.)</i></p>"
        )
        header_lbl.setWordWrap(True)
        layout.addWidget(header_lbl)
        
        h_batch = QHBoxLayout()
        btn_all_new = QPushButton("✨ 모든 항목 [새 크롤링 정보] 적용")
        btn_all_old = QPushButton("🛡️ 모든 항목 [기존 DB 정보] 유지")
        
        btn_all_new.clicked.connect(self.select_all_new)
        btn_all_old.clicked.connect(self.select_all_old)
        
        h_batch.addWidget(btn_all_new)
        h_batch.addWidget(btn_all_old)
        layout.addLayout(h_batch)
        
        self.table = QTableWidget()
        self.table.setColumnCount(5)
        self.table.setHorizontalHeaderLabels([
            "제품 (제조사 & 제품번호)", "항목 (Field)", "기존 DB 정보", "새 크롤링 정보", "적용 선택"
        ])
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Interactive)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Interactive)
        self.table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeMode.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeMode.Fixed)
        self.table.setColumnWidth(4, 160)
        self.table.setColumnWidth(0, 180)
        self.table.setColumnWidth(1, 130)
        
        total_rows = sum(len(c["conflicting_fields"]) for c in self.conflicts)
        self.table.setRowCount(total_rows)
        
        row_idx = 0
        for item_idx, conflict in enumerate(self.conflicts):
            item_title = f"{conflict['manufacturer']} ({conflict['catalog_no']})"
            for f_info in conflict["conflicting_fields"]:
                field_name = f_info["field"]
                old_val = f_info["old"]
                new_val = f_info["new"]
                
                item_cell = QTableWidgetItem(item_title)
                field_cell = QTableWidgetItem(field_name)
                old_cell = QTableWidgetItem(str(old_val))
                new_cell = QTableWidgetItem(str(new_val))
                
                item_cell.setFlags(item_cell.flags() ^ Qt.ItemFlag.ItemIsEditable)
                field_cell.setFlags(field_cell.flags() ^ Qt.ItemFlag.ItemIsEditable)
                old_cell.setFlags(old_cell.flags() ^ Qt.ItemFlag.ItemIsEditable)
                new_cell.setFlags(new_cell.flags() ^ Qt.ItemFlag.ItemIsEditable)
                
                self.table.setItem(row_idx, 0, item_cell)
                self.table.setItem(row_idx, 1, field_cell)
                self.table.setItem(row_idx, 2, old_cell)
                self.table.setItem(row_idx, 3, new_cell)
                
                choice_widget = QWidget()
                h_choice = QHBoxLayout(choice_widget)
                h_choice.setContentsMargins(4, 2, 4, 2)
                
                bg = QButtonGroup(choice_widget)
                rb_new = QRadioButton("신규")
                rb_old = QRadioButton("기존")
                rb_new.setChecked(True)
                
                bg.addButton(rb_new, 1)
                bg.addButton(rb_old, 2)
                
                h_choice.addWidget(rb_new)
                h_choice.addWidget(rb_old)
                
                self.table.setCellWidget(row_idx, 4, choice_widget)
                self.radio_groups[(item_idx, field_name)] = (bg, rb_new, rb_old)
                
                row_idx += 1
                
        layout.addWidget(self.table)
        
        h_bottom = QHBoxLayout()
        btn_apply = QPushButton("선택한 항목 반영 및 DB 저장")
        btn_cancel = QPushButton("취소 (변경 사항 버림)")
        
        btn_apply.setStyleSheet("font-weight: bold; background-color: #2b579a; color: white; padding: 6px 15px;")
        btn_cancel.setStyleSheet("padding: 6px 15px;")
        
        btn_apply.clicked.connect(self.accept)
        btn_cancel.clicked.connect(self.reject)
        
        h_bottom.addStretch()
        h_bottom.addWidget(btn_apply)
        h_bottom.addWidget(btn_cancel)
        layout.addLayout(h_bottom)

    def select_all_new(self):
        for (item_idx, field_name), (bg, rb_new, rb_old) in self.radio_groups.items():
            rb_new.setChecked(True)

    def select_all_old(self):
        for (item_idx, field_name), (bg, rb_new, rb_old) in self.radio_groups.items():
            rb_old.setChecked(True)

    def get_resolved_records(self):
        resolved = []
        for item_idx, conflict in enumerate(self.conflicts):
            rec = dict(conflict["crawled"])
            for f_info in conflict["conflicting_fields"]:
                field_name = f_info["field"]
                if (item_idx, field_name) in self.radio_groups:
                    bg, rb_new, rb_old = self.radio_groups[(item_idx, field_name)]
                    if rb_old.isChecked():
                        rec[field_name] = f_info["old"]
                    else:
                        rec[field_name] = f_info["new"]
            resolved.append(rec)
        return resolved

class DbUpdateWorker(QThread):
    progress = pyqtSignal(str)
    finished = pyqtSignal(dict)
    
    def __init__(self, config, opts):
        super().__init__()
        self.config = config
        self.opts = opts
        
    def run(self):
        try:
            import os, urllib.parse, datetime
            src_path_raw = self.config.get("source_file", "")
            if src_path_raw.startswith("file://"):
                src_path_raw = urllib.parse.unquote(src_path_raw.replace("file:///", "").replace("file://", ""))
            src_folder = os.path.dirname(os.path.abspath(src_path_raw))
            target_path = os.path.join(src_folder, "ChemicalList.xlsx")
            
            from core.db_manager import DBManager
            from scrapers.thermofisher import ThermofisherScraper
            from scrapers.tci import TciScraper
            from scrapers.aldrich import AldrichScraper
            from scrapers.abcam import AbcamScraper
            import pandas as pd
            
            from seleniumbase import SB
            
            self.progress.emit("DB 수동 업데이트 시작...")
            
            if not os.path.exists(target_path):
                raise Exception("ChemicalList.xlsx 파일이 없습니다.")
                
            db_manager = DBManager(target_path)
            df = db_manager.load_db()
            
            if df.empty:
                raise Exception("DB가 비어있습니다. 먼저 오더북 동기화를 실행하세요.")
                
            updated = 0
            is_headless = self.config.get("headless", True)
            
            items_to_process = {}
            
            def norm_man(m):
                return DBManager.normalize_manufacturer(m).lower()
            
            # 1. Load DB items
            for idx, row in df.iterrows():
                man = str(row.get("Manufacturer", row.get("제조사", ""))).strip()
                cat = str(row.get("Catalog No.", row.get("제품번호", ""))).strip()
                if cat.endswith(".0"):
                    cat = cat[:-2]
                link = str(row.get("Detail_Link", row.get("상세정보_링크", ""))).strip()
                chem_name = str(row.get("Product Name", row.get("시약명", ""))).strip()
                if man and cat and man != "nan" and cat != "nan":
                    items_to_process[(norm_man(man), cat)] = {"link": link, "name": chem_name}
                    
            # 2. Load ChemicalList sheet items
            with pd.ExcelFile(target_path) as xl:
                target_sheet = None
                for sn in xl.sheet_names:
                    if sn not in ["DB", "가이드(Guide)", "Guide", "index", "Old Chemical List"]:
                        target_sheet = sn
                        break
                        
                if target_sheet:
                    chem_df = xl.parse(target_sheet, header=0)
                    man_col, cat_col, name_col = None, None, None
                    for col in chem_df.columns:
                        col_s = str(col).lower()
                        if any(x in col_s for x in ["company", "제조사", "manufacturer", "maker"]): 
                            man_col = col
                        if any(x in col_s for x in ["catalog", "제품번호", "품번"]): 
                            cat_col = col
                        if any(x in col_s for x in ["name", "품명", "시약명", "제품명"]):
                            name_col = col
                    if man_col and cat_col:
                        for idx, row in chem_df.iterrows():
                            man = str(row.get(man_col, "")).strip()
                            cat = str(row.get(cat_col, "")).strip()
                            if cat.endswith(".0"):
                                cat = cat[:-2]
                            chem_name = str(row.get(name_col, "")).strip() if name_col else ""
                            if man and cat and man != "nan" and cat != "nan":
                                n_man = norm_man(man)
                                if (n_man, cat) not in items_to_process:
                                    items_to_process[(n_man, cat)] = {"link": "", "name": chem_name}
                                
            # Calculate how many items will be updated
            total_updates = 0
            for (man, cat), data in items_to_process.items():
                link = data["link"]
                should_update = False
                if self.opts["mode"] == "all":
                    should_update = True
                elif self.opts["mode"] == "missing":
                    if not link or "해당 제품 없음" in link or "Product Not Found" in link or "오류" in link or "요망" in link or str(link) == "nan":
                        should_update = True
                elif self.opts["mode"] == "specific":
                    if cat in self.opts["specific_products"]:
                        should_update = True
                if should_update:
                    total_updates += 1
            
            fast_mode = total_updates <= 5
            crawled_results_batch = []

            with SB(uc=True, headless=is_headless) as sb:
                for (man, cat), data in items_to_process.items():
                    link = data["link"]
                    fallback_name = data["name"]
                    should_update = False
                    if self.opts["mode"] == "all":
                        should_update = True
                    elif self.opts["mode"] == "missing":
                        if not link or "해당 제품 없음" in link or "Product Not Found" in link or "오류" in link or "요망" in link or str(link) == "nan":
                            should_update = True
                    elif self.opts["mode"] == "specific":
                        if cat in self.opts["specific_products"]:
                            should_update = True
                            
                    if should_update:
                        self.progress.emit(f"[{man}] {cat} 크롤링 중...")
                        man_lower = man.lower()
                        crawled_data = None
                        try:
                            if "thermo" in man_lower or "alfa" in man_lower or "fisher" in man_lower or "invitrogen" in man_lower or "acros" in man_lower:
                                crawled_data = ThermofisherScraper(browser_context=sb, fast_mode=fast_mode, base_dir=src_folder).scrape(cat)
                            elif "tci" in man_lower or "tokyo" in man_lower:
                                crawled_data = TciScraper(browser_context=sb, fast_mode=fast_mode, base_dir=src_folder).scrape(cat)
                            elif "sigma" in man_lower or "aldrich" in man_lower or "millipore" in man_lower:
                                crawled_data = AldrichScraper(browser_context=sb, fast_mode=fast_mode, base_dir=src_folder).scrape(cat)
                            elif "abcam" in man_lower:
                                crawled_data = AbcamScraper(browser_context=sb, fast_mode=fast_mode, base_dir=src_folder).scrape(cat)
                            else:
                                crawled_data = {"error": "Manual Input Required"}
                        except Exception as e:
                            crawled_data = {"error": f"크롤링 오류: {e}"}
                            
                        norm_m = DBManager.normalize_manufacturer(man)
                        db_result = {
                            "Manufacturer": norm_m,
                            "Catalog No.": cat,
                            "Product Name": fallback_name,
                            "CAS No.": "-",
                            "Storage Temp.": "-",
                            "Signal Word": "-",
                            "Key Hazards": "-",
                            "Detailed Hazard Classification": "-",
                            "Sensitivity": "-",
                            "Detail_Link": "-",
                            "SDS_Link": "-",
                            "SDS_Local_Path": "-",
                            "Revision Date": "-"
                        }
                        if crawled_data:
                            if "error" in crawled_data:
                                is_man_req = crawled_data["error"] in ["Manual Input Required", "Manual Entry Required"]
                                db_result["Detail_Link"] = "-" if is_man_req else "Product Not Found"
                                db_result["CAS No."] = "Manual Input Required" if is_man_req else "Search Failed"
                                db_result["Product Name"] = fallback_name
                                db_result["Storage Temp."] = "-"
                                db_result["Sensitivity"] = "-"
                                db_result["Revision Date"] = "-"
                            else:
                                p_name = crawled_data.get("Product Name") or crawled_data.get("시약명")
                                cas_no = crawled_data.get("CAS No.") or crawled_data.get("CAS Number")
                                s_temp = crawled_data.get("Storage Temp.") or crawled_data.get("보관온도")
                                sig_word = crawled_data.get("Signal Word") or crawled_data.get("신호어")
                                key_haz = crawled_data.get("Key Hazards") or crawled_data.get("주요위험")
                                det_haz = crawled_data.get("Detailed Hazard Classification") or crawled_data.get("상세 위험분류")
                                sens = crawled_data.get("Sensitivity") or crawled_data.get("민감성")
                                det_link = crawled_data.get("Detail_Link") or crawled_data.get("상세정보_링크")
                                sds_link = crawled_data.get("SDS_Link")
                                sds_path = crawled_data.get("SDS_Local_Path")

                                if p_name in ["제조사 홈페이지에서 검색 실패", "검색 실패", "Product Not Found", "", None] or str(p_name).strip() == "":
                                    db_result["Product Name"] = fallback_name
                                    db_result["CAS No."] = "Search Failed"
                                    db_result["Detail_Link"] = "Product Not Found"
                                    db_result["Storage Temp."] = "-"
                                    db_result["Sensitivity"] = "-"
                                    db_result["Revision Date"] = "-"
                                else:
                                    db_result["Product Name"] = p_name
                                    if cas_no in ["정보 없음", "", "nan", "None", None, "-"]:
                                        db_result["CAS No."] = "N/A"
                                    else:
                                        db_result["CAS No."] = cas_no
                                        
                                    db_result["Detail_Link"] = det_link or "-"
                                    db_result["SDS_Link"] = sds_link or "-"
                                    db_result["SDS_Local_Path"] = sds_path or "-"
                                    db_result["Signal Word"] = sig_word if sig_word and sig_word not in ["", "None", "nan"] else "-"
                                    db_result["Key Hazards"] = key_haz if key_haz and key_haz not in ["", "None", "nan"] else "-"
                                    db_result["Detailed Hazard Classification"] = det_haz if det_haz and det_haz not in ["", "None", "nan"] else "-"

                                    db_result["Storage Temp."] = DBManager.normalize_temperature(s_temp)
                                    extracted_sens = DBManager.extract_sensitivity(det_haz or "")
                                    db_result["Sensitivity"] = extracted_sens if extracted_sens != "-" else (sens if sens and sens not in ["", "None", "nan"] else "-")
                                    db_result["Revision Date"] = datetime.datetime.now().strftime("%Y-%m-%d")
                        else:
                            db_result["Detail_Link"] = "Product Not Found"
                            db_result["CAS No."] = "Search Failed"
                            db_result["Product Name"] = fallback_name
                            db_result["Storage Temp."] = "-"
                            db_result["Sensitivity"] = "-"
                            db_result["Revision Date"] = "-"
                            
                        crawled_results_batch.append(db_result)
                        updated += 1

            # Detect conflicts in-memory (DO NOT WRITE TO FILE YET)
            conflicts = []
            non_conflicts = []
            
            fields_to_check = [
                "Product Name", "CAS No.", "Storage Temp.", "Sensitivity", 
                "Signal Word", "Key Hazards", "Detailed Hazard Classification", 
                "Detail_Link", "SDS_Link", "SDS_Local_Path"
            ]
            
            for db_result in crawled_results_batch:
                man = db_result.get("Manufacturer", "")
                cat = db_result.get("Catalog No.", "")
                
                existing = db_manager.get_product(man, cat)
                if not existing:
                    non_conflicts.append(db_result)
                else:
                    conflicting_fields = []
                    for f in fields_to_check:
                        old_val = str(existing.get(f, "")).strip() if existing.get(f) is not None else ""
                        new_val = str(db_result.get(f, "")).strip() if db_result.get(f) is not None else ""
                        
                        if old_val in ["", "-", "nan", "None"]:
                            continue
                        if new_val in ["", "-", "nan", "None", "Search Failed", "Product Not Found", "Manual Input Required", "Manual Entry Required"]:
                            continue
                            
                        if old_val.lower() != new_val.lower():
                            conflicting_fields.append({
                                "field": f,
                                "old": old_val,
                                "new": new_val
                            })
                            
                    if conflicting_fields:
                        conflicts.append({
                            "manufacturer": man,
                            "catalog_no": cat,
                            "product_name": db_result.get("Product Name", ""),
                            "crawled": db_result,
                            "existing": existing,
                            "conflicting_fields": conflicting_fields
                        })
                    else:
                        merged = dict(existing)
                        for k, v in db_result.items():
                            if v not in ["", "-", "nan", "None", None, "Search Failed", "Product Not Found", "Manual Input Required"]:
                                merged[k] = v
                        non_conflicts.append(merged)

            self.finished.emit({
                "success": True, 
                "conflicts": conflicts, 
                "non_conflicts": non_conflicts, 
                "target_path": target_path
            })
        except Exception as e:
            self.finished.emit({"success": False, "error": str(e)})

class SyncWorker(QThread):
    progress = pyqtSignal(str)
    finished = pyqtSignal(dict)
    
    def __init__(self, config):
        super().__init__()
        self.config = config
        
    def run(self):
        try:
            engine = SyncEngine(self.config, callback_progress=self.progress.emit)
            result = engine.run_sync()
            self.finished.emit(result)
        except Exception as e:
            self.finished.emit({"success": False, "error": str(e)})

class MainWindow(QMainWindow):
    config_updated = pyqtSignal(dict)
    tray_requested = pyqtSignal()
    close_requested = pyqtSignal()
    manual_sync_started = pyqtSignal()
    manual_sync_finished = pyqtSignal()

    def __init__(self, config_data):
        super().__init__()
        self.setWindowTitle("연구실 시약 주문 관리 시스템")
        self.resize(800, 600)
        self.config = config_data.copy()
        
        self.init_ui()

    def init_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)

        # 1. Config Group
        grp_config = QGroupBox("기본 설정")
        l_config = QVBoxLayout()
        
        # Source File
        h_file = QHBoxLayout()
        h_file.addWidget(QLabel("원본 파일:"))
        self.inp_file = QLineEdit(self.config.get("source_file", ""))
        self.inp_file.setReadOnly(True)
        btn_file = QPushButton("파일 찾기")
        btn_file.clicked.connect(self.select_file)
        h_file.addWidget(self.inp_file)
        h_file.addWidget(btn_file)
        l_config.addLayout(h_file)
        
        # Source Sheet
        h_sheet = QHBoxLayout()
        h_sheet.addWidget(QLabel("원본 시트:"))
        self.inp_sheet = QLineEdit(self.config.get("source_sheet", ""))
        btn_sheet = QPushButton("시트 선택")
        btn_sheet.clicked.connect(self.select_sheet)
        h_sheet.addWidget(self.inp_sheet)
        h_sheet.addWidget(btn_sheet)
        
        h_sheet.addWidget(QLabel("헤더 행 번호:"))
        self.spin_header = QSpinBox()
        self.spin_header.setRange(1, 20)
        self.spin_header.setValue(int(self.config.get("header_row", 1)))
        h_sheet.addWidget(self.spin_header)
        l_config.addLayout(h_sheet)
        
        # Mapping Button
        btn_map = QPushButton("헤더 매핑 설정 (항목 추가/삭제)")
        btn_map.clicked.connect(self.open_mapping)
        l_config.addWidget(btn_map)
        
        grp_config.setLayout(l_config)
        main_layout.addWidget(grp_config)

        # 2. Automation Configuration
        g_automation = QGroupBox("자동화 설정")
        form_auto = QFormLayout()
        
        # Add Enable/Disable Toggle
        self.btn_toggle_auto = QPushButton("자동 동기화 정지")
        self.btn_toggle_auto.setCheckable(True)
        self.btn_toggle_auto.clicked.connect(self.toggle_automation)
        
        self.chk_watch = QCheckBox("원본 파일 변경 시 자동 업데이트")
        self.chk_watch.setChecked(self.config.get("watch_enabled", False))
        self.spin_interval = QSpinBox()
        self.spin_interval.setRange(0, 1440) # 0 means disabled, max 24 hours
        self.spin_interval.setValue(self.config.get("sync_interval_minutes", 0))
        self.spin_interval.setSuffix(" 분 (0: 사용안함)")
        
        self.chk_startup = QCheckBox("윈도우 시작 시 자동 실행 (시스템 트레이)")
        self.chk_startup.setChecked(self.config.get("run_on_startup", False))
        
        self.chk_headless = QCheckBox("DB 크롤링 시 브라우저 숨기기 (Headless 모드)")
        self.chk_headless.setChecked(self.config.get("headless", True))
        
        form_auto.addRow(self.btn_toggle_auto)
        form_auto.addRow(self.chk_watch)
        form_auto.addRow("주기적 업데이트 간격:", self.spin_interval)
        form_auto.addRow(self.chk_startup)
        form_auto.addRow(self.chk_headless)
        g_automation.setLayout(form_auto)
        main_layout.addWidget(g_automation)
        
        btn_apply_auto = QPushButton("자동화 설정 적용")
        btn_apply_auto.clicked.connect(self.save_config)
        main_layout.addWidget(btn_apply_auto)

        # 3. Actions Group
        h_actions = QHBoxLayout()
        
        self.btn_sync = QPushButton("수동 동기화 실행")
        self.btn_sync.setMinimumHeight(40)
        self.btn_sync.clicked.connect(self.run_manual_sync)
        
        self.btn_pdf = QPushButton("PDF 자동 인쇄파일 생성")
        self.btn_pdf.setMinimumHeight(40)
        self.btn_pdf.clicked.connect(self.run_pdf_export)
        
        self.btn_db_update = QPushButton("DB 수동 업데이트")
        self.btn_db_update.setMinimumHeight(40)
        self.btn_db_update.clicked.connect(self.run_db_update)
        
        self.btn_sds = QPushButton("SDS 일괄 병합 생성")
        self.btn_sds.setMinimumHeight(40)
        self.btn_sds.clicked.connect(self.run_sds_batch)
        
        self.btn_tray = QPushButton("백그라운드 실행 (트레이로 최소화)")
        self.btn_tray.setMinimumHeight(40)
        self.btn_tray.clicked.connect(self.minimize_to_tray)
        
        self.btn_guide = QPushButton("도움말 및 주의사항")
        self.btn_guide.setMinimumHeight(40)
        self.btn_guide.clicked.connect(self.show_guide)
        
        h_actions.addWidget(self.btn_sync)
        h_actions.addWidget(self.btn_db_update)
        h_actions.addWidget(self.btn_pdf)
        h_actions.addWidget(self.btn_sds)
        h_actions.addWidget(self.btn_tray)
        h_actions.addWidget(self.btn_guide)
        main_layout.addLayout(h_actions)
        
        # Status Bar
        self.statusBar().showMessage("대기 중...")
        
        self.apply_colors(main_layout)

    def toggle_automation(self):
        if self.btn_toggle_auto.isChecked():
            # Paused
            self.btn_toggle_auto.setText("자동 동기화 시작")
            self.chk_watch.setEnabled(False)
            self.spin_interval.setEnabled(False)
            self.statusBar().showMessage("자동 동기화 정지됨")
            # We need to stop the watcher via config
            temp_config = self.config.copy()
            temp_config["watch_enabled"] = False
            temp_config["sync_interval_minutes"] = 0
            self.config_updated.emit(temp_config)
        else:
            # Resumed
            self.btn_toggle_auto.setText("자동 동기화 정지")
            self.chk_watch.setEnabled(True)
            self.spin_interval.setEnabled(True)
            self.statusBar().showMessage("자동 동기화 활성화됨")
            self.save_config()

    def apply_colors(self, main_layout):
        h_bottom = QHBoxLayout()
        
        # Legend
        grp_legend = QGroupBox("색상 범례 (ChemicalList 경고)")
        l_legend = QVBoxLayout()
        
        btn_color = QPushButton("색상 설정 변경")
        btn_color.clicked.connect(self.open_color_dialog)
        l_legend.addWidget(btn_color)
        
        lbl_legend = QLabel(
            "이곳에서 지정한 색상 규칙에 따라<br>"
            "오더북과 ChemicalList 엑셀 파일의<br>"
            "특정 빈 셀에 조건부 서식으로 색상이 칠해집니다."
        )
        l_legend.addWidget(lbl_legend)
        l_legend.addStretch()
        grp_legend.setLayout(l_legend)
        
        # Log
        grp_log = QGroupBox("로그")
        l_log = QVBoxLayout()
        self.txt_log = QTextEdit()
        self.txt_log.setReadOnly(True)
        l_log.addWidget(self.txt_log)
        grp_log.setLayout(l_log)
        
        h_bottom.addWidget(grp_legend, 1)
        h_bottom.addWidget(grp_log, 2)
        main_layout.addLayout(h_bottom)

    def select_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "원본 오더북 파일 선택", "", "Excel Files (*.xlsx *.xlsm)")
        if file_path:
            import urllib.parse
            if file_path.startswith("file:///"):
                file_path = urllib.parse.unquote(file_path[8:])
            file_path = os.path.abspath(file_path)
            self.inp_file.setText(file_path)
            self.save_config()

    def select_sheet(self):
        src_path = self.inp_file.text().strip()
        
        import urllib.parse
        if src_path.startswith("file:///"):
            src_path = urllib.parse.unquote(src_path[8:])
        src_path = os.path.abspath(src_path)

        if not src_path or not os.path.exists(src_path):
            QMessageBox.warning(self, "경고", "먼저 유효한 원본 파일을 선택하세요.")
            return
            
        import win32com.client
        from PyQt6.QtWidgets import QInputDialog
        try:
            self.log("엑셀 시트 목록을 불러오는 중...")
            excel = win32com.client.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            wb = excel.Workbooks.Open(src_path, ReadOnly=True, UpdateLinks=False)
            sheets = [sheet.Name for sheet in wb.Worksheets]
            wb.Close(False)
            excel.Quit()
            
            if sheets:
                sheet, ok = QInputDialog.getItem(self, "시트 선택", "동기화할 시트를 선택하세요:", sheets, 0, False)
                if ok and sheet:
                    self.inp_sheet.setText(sheet)
                    self.save_config()
        except Exception as e:
            self.log(f"시트 목록을 불러오는 중 오류가 발생했습니다: {e}")
            QMessageBox.critical(self, "오류", f"엑셀 오류: {e}")

    def open_mapping(self):
        dlg = MappingDialog(self.config, self)
        if dlg.exec():
            self.config = dlg.get_config()
            self.save_config()
            self.log("헤더 매핑이 업데이트 되었습니다.")

    def open_color_dialog(self):
        dlg = ColorDialog(self.config, self)
        if dlg.exec():
            self.config = dlg.get_config()
            self.save_config()
            self.log("색상 설정이 업데이트 되었습니다.")

    def show_guide(self):
        guide_text = """<h3>[프로그램 및 Chemical List 사용 가이드]</h3>
<b>1. 기본 동작 원리 및 저장 기준</b><br>
- 프로그램은 원본 Orderbook 파일의 변경을 감지하여 데이터를 읽어옵니다.<br>
- 폴더 내에서 이름에 적힌 날짜/시간이 가장 최신인 ChemicalList 파일을 기준으로 삼습니다.<br>
- 업데이트 완료 시 덮어쓰기 대신 새로운 날짜/시간이 적힌 새 파일(ChemicalList_YYYYMMDD_HHMMSS)을 생성합니다.<br>
- 기존의 낡은 파일들은 안전하게 'Old Chemical List' 폴더로 자동 이동됩니다.<br>
- 로그 파일에는 변경되는 내용을 월별로 기입합니다.<br><br>

<b>2. 사용자 작성 시 유의사항 (Orderbook 및 Chemical List)</b><br>
- 오더북의 "번호, 날짜, 주문자, 품목명, 시약, 수령확인, 회사, 수량, 용량, CAS 번호, 품번, 보관온도" 열의 내용을 자동으로 동기화합니다.<br>
- 오더북을 작성할 때 없는 열은 새로 생성해주고 해당 셀의 내용을 충실하게 적어주세요<br>
- 오더북의 시약 및 수령확인 셀에 모두 "O"표시가 되어있는 행만 사용합니다.<br>
- 동기화 기준: 오더북 '번호'를 기준으로 기존에 등록된 시약인지 새 시약인지 판단합니다.<br>
- 오더북에 Order No.가 있고 시약 및 수령확인에 "O"가 표시된 행은 모두 동기화합니다.<br>
- 시약리스트에 한번 업데이트가 되었어도 다음 동기화 때 오더북에 있는 데이터를 덮어씁니다. 오더북에서 불러온 데이터는 수정하지 마세요<br>
- Aliquat하는 경우 시약리스트의 항목을 삭제하지 말고 used에 숫자를 기입하고 새로운 행에 내용을 기입하고 개수를 적어주세요.<br>
- 시약리스트에 Order No.가 없는 행은 동기화에서 제외됩니다.<br>
- 수동으로 데이터를 추가할 때에는 새로운 행에 Order No.없이 내용을 적어 놓으면 됩니다 (product name은 필수입니다).<br>
- 수령 확인 : 수령 후 수령확인 열에 'O' 또는 'ㅇ'을 입력하면 수령으로 간주되며, 시약리스트 업데이트에 사용합니다.<br>
- 빈 줄 금지: 데이터 중간에 완전히 비어있는 줄이 있으면 데이터를 읽다가 중단될 수 있으므로 차례대로 기입하세요.<br>
- 드롭다운 선택: 캐비넷(Cabinet)은 Room과 Storage Temp.에 따라 동적으로 변하므로 잘못된 값을 억지로 쓰지 마세요.<br><br>

<b>3. 프로그램 사용 주의사항</b><br>
- 열린 파일 처리: Chemical List 엑셀 파일이 켜져 있어도 프로그램이 알아서 우회하여 새 파일을 생성합니다.<br>
- 단, 원본인 Orderbook 파일은 엑셀에서 저장을 완료해야만 프로그램이 변경 사항을 정확히 감지할 수 있습니다.<br>
- 자동 동기화 켜짐 상태에서는 Orderbook을 저장할 때마다 병합이 진행되므로, 수동 제어를 원하시면 정지 버튼을 누르세요.<br><br>

<hr>
<b>제작자:</b> Jeonghun Lee<br>
<b>문의:</b> jhl22@hanyang.ac.kr
"""
        msg = QMessageBox(self)
        msg.setWindowTitle("도움말 및 주의사항")
        msg.setText(guide_text)
        msg.exec()

    def save_config(self):
        self.config["source_file"] = self.inp_file.text()
        self.config["source_sheet"] = self.inp_sheet.text()
        self.config["header_row"] = self.spin_header.value()
        
        if self.btn_toggle_auto.isChecked():
            self.config["watch_enabled"] = False
            self.config["sync_interval_minutes"] = 0
        else:
            self.config["watch_enabled"] = self.chk_watch.isChecked()
            self.config["sync_interval_minutes"] = self.spin_interval.value()
        
        
        self.config["run_on_startup"] = self.chk_startup.isChecked()
        self.config["headless"] = self.chk_headless.isChecked()
        
        import utils.startup_manager as sm
        sm.set_run_on_startup(self.config["run_on_startup"])
        
        self.config_updated.emit(self.config)

    def log(self, message):
        import datetime
        import os
        now = datetime.datetime.now()
        ts = now.strftime("[%H:%M:%S]")
        full_msg = f"{ts} {message}"
        self.txt_log.append(full_msg)
        
        try:
            import sys
            if getattr(sys, 'frozen', False):
                app_root = os.path.dirname(sys.executable)
            else:
                app_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
            log_dir = os.path.join(app_root, "logs")
            if not os.path.exists(log_dir):
                os.makedirs(log_dir)
            file_name = f"app_log_{now.strftime('%Y_%m')}.txt"
            file_path = os.path.join(log_dir, file_name)
            with open(file_path, "a", encoding="utf-8") as f:
                f.write(f"[{now.strftime('%Y-%m-%d %H:%M:%S')}] {message}\n")
        except Exception:
            pass

    def run_manual_sync(self):
        self.save_config()
        self.log("--- 수동 동기화 시작 ---")
        self.statusBar().showMessage("수동 동기화 진행 중...")
        self.btn_sync.setEnabled(False)
        self.manual_sync_started.emit()
        
        self.sync_worker = SyncWorker(self.config)
        self.sync_worker.progress.connect(self.log)
        self.sync_worker.finished.connect(self.on_sync_finished)
        self.sync_worker.start()

    def on_sync_finished(self, result):
        self.btn_sync.setEnabled(True)
        self.manual_sync_finished.emit()
        self.statusBar().showMessage("동기화 완료", 5000)
        
        if result.get("success"):
            new = result.get('new', 0)
            upd = result.get('updated', 0)
            self.log(f"동기화 완료! (신규: {new}건, 업데이트: {upd}건)")
            QMessageBox.information(self, "성공", f"동기화가 완료되었습니다.\n신규: {new}건\n업데이트: {upd}건")
        else:
            self.log(f"오류 발생: {result.get('error')}")
            QMessageBox.critical(self, "오류", f"동기화 중 오류가 발생했습니다:\n{result.get('error')}")

    def run_pdf_export(self):
        dialog = PdfDialog(self.config["target_headers"], self)
        if dialog.exec():
            mode = dialog.get_mode()
            selected_cols = dialog.get_selected_columns()
            only_status_o = dialog.get_only_status_o()
            
            out_path, _ = QFileDialog.getSaveFileName(self, "PDF 저장 위치 선택", "ChemicalList_Print.pdf", "PDF Files (*.pdf)")
            if out_path:
                self.save_config()
                self.log(f"--- PDF 생성 시작 (모드 {mode}) ---")
                self.btn_pdf.setEnabled(False)
                self.statusBar().showMessage("PDF 생성 중...")
                self.pdf_worker = PdfWorker(self.config, mode, selected_cols, only_status_o, out_path)
                self.pdf_worker.progress.connect(self.log)
                self.pdf_worker.finished.connect(self.on_pdf_finished)
                self.pdf_worker.start()
                
    def on_pdf_finished(self, result):
        self.btn_pdf.setEnabled(True)
        self.statusBar().showMessage("PDF 생성 완료", 5000)
        if result.get("success"):
            self.log("PDF 생성 완료!")
            QMessageBox.information(self, "성공", f"PDF 파일이 성공적으로 생성되었습니다.\n{result['path']}")
        else:
            self.log(f"PDF 생성 오류: {result.get('error')}")
            QMessageBox.critical(self, "오류", f"PDF 생성 중 오류가 발생했습니다:\n{result.get('error')}")

    def run_sds_batch(self):
        from PyQt6.QtWidgets import QInputDialog
        import datetime
        today_str = datetime.datetime.now().strftime("%Y-%m-%d")
        date_str, ok = QInputDialog.getText(self, "SDS 일괄 병합", "기준 갱신일을 입력하세요 (YYYY-MM-DD):", QLineEdit.EchoMode.Normal, today_str)
        if ok and date_str:
            out_path, _ = QFileDialog.getSaveFileName(self, "SDS 병합 파일 저장 위치 선택", f"SDS_Batch_{date_str.replace('-','')}.pdf", "PDF Files (*.pdf)")
            if out_path:
                self.save_config()
                self.btn_sds.setEnabled(False)
                self.statusBar().showMessage("SDS 병합 중...")
                self.sds_worker = SdsWorker(self.config, date_str, out_path)
                self.sds_worker.progress.connect(self.log)
                self.sds_worker.finished.connect(self.on_sds_finished)
                self.sds_worker.start()
                
    def on_sds_finished(self, result):
        self.btn_sds.setEnabled(True)
        self.statusBar().showMessage("SDS 병합 완료", 5000)
        if result.get("success"):
            self.log("SDS 병합 완료!")
            QMessageBox.information(self, "성공", f"SDS 병합 파일이 생성되었습니다.\n{result['path']}")
        else:
            self.log(f"SDS 병합 오류: {result.get('error')}")
            QMessageBox.critical(self, "오류", f"SDS 병합 중 오류가 발생했습니다:\n{result.get('error')}")

    def run_db_update(self):
        dlg = DbUpdateDialog(self.config, self)
        if dlg.exec():
            opts = dlg.get_options()
            self.save_config()
            self.btn_db_update.setEnabled(False)
            self.statusBar().showMessage("DB 수동 업데이트 중...")
            self.db_worker = DbUpdateWorker(self.config, opts)
            self.db_worker.progress.connect(self.log)
            self.db_worker.finished.connect(self.on_db_update_finished)
            self.db_worker.start()

    def on_db_update_finished(self, result):
        self.btn_db_update.setEnabled(True)
        self.statusBar().showMessage("크롤링 완료", 5000)
        
        if not result.get("success"):
            self.log(f"DB 업데이트 오류: {result.get('error')}")
            QMessageBox.critical(self, "오류", f"DB 업데이트 중 오류가 발생했습니다:\n{result.get('error')}")
            return

        conflicts = result.get("conflicts", [])
        non_conflicts = result.get("non_conflicts", [])
        target_path = result.get("target_path", "")
        
        records_to_save = list(non_conflicts)
        
        if conflicts:
            dialog = DBConflictResolverDialog(conflicts, parent=self)
            if dialog.exec() == QDialog.DialogCode.Accepted:
                resolved = dialog.get_resolved_records()
                records_to_save.extend(resolved)
            else:
                self.log("DB 수동 업데이트가 사용자에 의해 취소되었습니다. (파일 변경 없음)")
                self.statusBar().showMessage("DB 수동 업데이트 취소됨")
                QMessageBox.information(self, "취소", "DB 수동 업데이트가 취소되었습니다.\n엑셀 파일에는 아무 내용도 변경되지 않았습니다.")
                return

        if records_to_save:
            try:
                from core.db_manager import DBManager
                from core.sync_engine import refresh_chemical_list_formatting
                db_manager = DBManager(target_path)
                db_manager.add_products_batch(records_to_save)
                
                self.log("ChemicalList 시트 서식 및 유효성 검사 갱신 중...")
                refresh_chemical_list_formatting(target_path, self.config)
                
                self.log(f"DB 수동 업데이트 저장 완료! (총 {len(records_to_save)}개 반영)")
                QMessageBox.information(self, "성공", f"DB 수동 업데이트가 성공적으로 저장되었습니다.\n(총 {len(records_to_save)}개 반영)")
            except Exception as e:
                self.log(f"DB 저장 중 오류 발생: {e}")
                QMessageBox.critical(self, "오류", f"DB 저장 중 오류가 발생했습니다:\n{e}")
        else:
            self.log("업데이트할 새로운 DB 정보가 없습니다.")
            QMessageBox.information(self, "안내", "업데이트할 새로운 DB 정보가 없습니다.")

    def minimize_to_tray(self):
        self.save_config()
        self.tray_requested.emit()
        
    def closeEvent(self, event):
        self.close_requested.emit()
        event.accept()
