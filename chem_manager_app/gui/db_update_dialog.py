from PyQt6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QPushButton, 
                             QLabel, QRadioButton, QLineEdit, QMessageBox, QGroupBox, QListWidget, QListWidgetItem)
from PyQt6.QtCore import Qt
import os
import urllib.parse
from core.db_manager import DBManager

class DbUpdateDialog(QDialog):
    def __init__(self, config, parent=None):
        super().__init__(parent)
        self.setWindowTitle("DB 수동 업데이트")
        self.resize(500, 400)
        self.config = config
        self.mode = "missing"
        self.specific_products = []
        self.db_df = self.load_db_df()
        self.init_ui()

    def load_db_df(self):
        try:
            target_path = self.config.get("target_file", "")
            if not target_path or not os.path.exists(target_path):
                src_path_raw = self.config.get("source_file", "")
                if src_path_raw.startswith("file:///"):
                    src_path_raw = urllib.parse.unquote(src_path_raw[8:])
                src_folder = os.path.dirname(os.path.abspath(src_path_raw))
                
                candidates = [os.path.join(src_folder, f) for f in os.listdir(src_folder) if f.startswith("ChemicalList") and f.endswith(".xlsx") and not f.startswith("~$")]
                if candidates:
                    candidates.sort(key=os.path.getmtime, reverse=True)
                    target_path = candidates[0]
                else:
                    target_path = os.path.join(src_folder, "ChemicalList.xlsx")
            
            if os.path.exists(target_path):
                db_manager = DBManager(target_path)
                df = db_manager.load_db()
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
                item.setData(Qt.ItemDataRole.UserRole, cat) # Save cat in UserRole
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
                    cat = item.data(Qt.ItemDataRole.UserRole)
                    if cat:
                        checked_items.append(cat)
            
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
