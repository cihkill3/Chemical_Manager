from PyQt6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QPushButton, 
                             QLabel, QRadioButton, QLineEdit, QMessageBox, QGroupBox, QListWidget, QListWidgetItem)
from PyQt6.QtCore import Qt
import os
import urllib.parse
import tempfile
import shutil
import pandas as pd
from gui.styles import MODERN_STYLE
from core.db_manager import DBManager
from core.config_manager import resolve_target_file

class DbUpdateDialog(QDialog):
    def __init__(self, config, parent=None):
        super().__init__(parent)
        self.setWindowTitle("DB 수동 업데이트")
        self.resize(520, 420)
        self.setStyleSheet(MODERN_STYLE)
        self.config = config
        self.mode = "missing"
        self.specific_products = []
        self.db_df = self.load_db_df()
        self.init_ui()

    def load_db_df(self):
        try:
            target_path = resolve_target_file(self.config)
            
            if os.path.exists(target_path):
                try:
                    df = pd.read_excel(target_path, sheet_name="DB")
                except Exception as ex:
                    print(f"Direct pd.read_excel error ({ex}), trying temp copy...")
                    fd, temp_path = tempfile.mkstemp(prefix="chemical_db_read_", suffix=".xlsx")
                    os.close(fd)
                    try:
                        shutil.copy2(target_path, temp_path)
                        df = pd.read_excel(temp_path, sheet_name="DB")
                    finally:
                        try:
                            os.remove(temp_path)
                        except OSError as cleanup_error:
                            print(f"Temporary DB copy cleanup warning: {cleanup_error}")
                df = df.rename(columns=DBManager.COLUMN_MAP)
                df = df.loc[:, ~df.columns.duplicated()]
                return df
        except Exception as e:
            print("load_db_df error:", e)
        return None
        
    def init_ui(self):
        layout = QVBoxLayout(self)

        lbl_desc = QLabel("DB(ChemicalList.xlsx 'DB' 시트) 수동 업데이트 옵션을 선택하세요.\n크롤링에는 시간이 걸릴 수 있습니다.")
        layout.addWidget(lbl_desc)

        grp_opts = QGroupBox("업데이트 모드")
        l_opts = QVBoxLayout()

        self.rb_missing = QRadioButton("DB에 없는 항목(Missing)만 크롤링 추가")
        self.rb_missing.setChecked(True)
        
        self.rb_all = QRadioButton("전체 항목 다시 크롤링 (오래 걸림)")
        
        self.rb_specific = QRadioButton("특정 제품번호(Catalog No.)만 업데이트")
        self.list_specific = QListWidget()
        self.list_specific.setEnabled(False)
        
        # Populate list
        if self.db_df is not None and not self.db_df.empty:
            for idx, row in self.db_df.iterrows():
                man = str(row.get("Manufacturer", row.get("제조사", ""))).strip()
                cat = str(row.get("Catalog No.", row.get("제품번호", ""))).strip()
                if cat.endswith(".0"): cat = cat[:-2]
                
                if not man or not cat or man in ["nan", "None"] or cat in ["nan", "None"]:
                    continue
                name = str(row.get("Product Name", row.get("시약명", ""))).strip()
                if name in ["nan", "None"]: name = ""
                
                item_str = f"[{man}] {cat} - {name}" if name else f"[{man}] {cat}"
                item = QListWidgetItem(item_str)
                item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
                item.setCheckState(Qt.CheckState.Unchecked)
                item.setData(Qt.ItemDataRole.UserRole, DBManager.crawl_key(man, cat))
                self.list_specific.addItem(item)
        else:
            self.list_specific.addItem("DB 데이터를 불러올 수 없습니다.")

        def on_rb_toggled():
            self.list_specific.setEnabled(self.rb_specific.isChecked())

        self.rb_missing.toggled.connect(on_rb_toggled)
        self.rb_all.toggled.connect(on_rb_toggled)
        self.rb_specific.toggled.connect(on_rb_toggled)

        l_opts.addWidget(self.rb_missing)
        l_opts.addWidget(self.rb_all)
        l_opts.addWidget(self.rb_specific)
        l_opts.addWidget(self.list_specific)
        
        grp_opts.setLayout(l_opts)
        layout.addWidget(grp_opts)

        h_btn = QHBoxLayout()
        btn_ok = QPushButton("실행")
        btn_ok.clicked.connect(self.on_ok)
        btn_cancel = QPushButton("취소")
        btn_cancel.clicked.connect(self.reject)
        
        h_btn.addWidget(btn_ok)
        h_btn.addWidget(btn_cancel)
        layout.addLayout(h_btn)

    def on_ok(self):
        if self.rb_missing.isChecked():
            self.mode = "missing"
        elif self.rb_all.isChecked():
            self.mode = "all"
        elif self.rb_specific.isChecked():
            self.mode = "specific"
            checked_items = []
            for i in range(self.list_specific.count()):
                item = self.list_specific.item(i)
                if item.checkState() == Qt.CheckState.Checked:
                    product_key = item.data(Qt.ItemDataRole.UserRole)
                    if product_key:
                        checked_items.append(product_key)
            
            if not checked_items:
                QMessageBox.warning(self, "경고", "업데이트할 항목을 하나 이상 체크하세요.")
                return
            self.specific_products = checked_items
        
        self.accept()

    def get_options(self):
        return {
            "mode": self.mode,
            "specific_products": self.specific_products
        }
