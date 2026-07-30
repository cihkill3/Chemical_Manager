import os
from PyQt6.QtWidgets import (QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
                             QPushButton, QLabel, QLineEdit, QFileDialog, QGroupBox, 
                             QTextEdit, QMessageBox, QSpinBox, QCheckBox, QComboBox, QFormLayout)
from PyQt6.QtCore import pyqtSignal, Qt, QThread
from gui.mapping_dialog import MappingDialog
from gui.color_dialog import ColorDialog
from gui.pdf_dialog import PdfDialog
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
        
        form_auto.addRow(self.btn_toggle_auto)
        form_auto.addRow(self.chk_watch)
        form_auto.addRow("주기적 업데이트 간격:", self.spin_interval)
        form_auto.addRow(self.chk_startup)
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
        
        self.btn_tray = QPushButton("백그라운드 실행 (트레이로 최소화)")
        self.btn_tray.setMinimumHeight(40)
        self.btn_tray.clicked.connect(self.minimize_to_tray)
        
        self.btn_guide = QPushButton("도움말 및 주의사항")
        self.btn_guide.setMinimumHeight(40)
        self.btn_guide.clicked.connect(self.show_guide)
        
        h_actions.addWidget(self.btn_sync)
        h_actions.addWidget(self.btn_pdf)
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

    def minimize_to_tray(self):
        self.save_config()
        self.tray_requested.emit()
        
    def closeEvent(self, event):
        self.close_requested.emit()
        event.accept()
