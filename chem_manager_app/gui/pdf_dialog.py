from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, 
    QRadioButton, QButtonGroup, QListWidget, QListWidgetItem, 
    QCheckBox, QPushButton, QDialogButtonBox, QMessageBox
)
from PyQt6.QtCore import Qt

class PdfDialog(QDialog):
    def __init__(self, target_headers, parent=None):
        super().__init__(parent)
        self.setWindowTitle("PDF 고급 인쇄 옵션")
        self.setMinimumWidth(400)
        self.target_headers = target_headers
        
        layout = QVBoxLayout(self)
        
        # 1. Grouping Mode
        layout.addWidget(QLabel("1. 그룹화 방식 선택"))
        self.mode_group = QButtonGroup(self)
        self.radio_room = QRadioButton("룸(Room) 별로 묶어서 페이지 나누기")
        self.radio_detail = QRadioButton("룸 + 보관온도 + 캐비넷 별로 묶어서 페이지 나누기")
        self.radio_room.setChecked(True)
        self.mode_group.addButton(self.radio_room, 1)
        self.mode_group.addButton(self.radio_detail, 2)
        layout.addWidget(self.radio_room)
        layout.addWidget(self.radio_detail)
        
        layout.addSpacing(10)
        
        # 2. Status Filter
        layout.addWidget(QLabel("2. 필터 옵션"))
        self.chk_status_o = QCheckBox("Status가 'X'가 아닌 항목만 인쇄 (잔량이 있는 시약만)")
        self.chk_status_o.setChecked(True)
        layout.addWidget(self.chk_status_o)
        
        layout.addSpacing(10)
        
        # 3. Column Selection
        layout.addWidget(QLabel("3. 인쇄할 열 선택"))
        self.list_cols = QListWidget()
        for header in self.target_headers:
            item = QListWidgetItem(header)
            item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
            item.setCheckState(Qt.CheckState.Checked)
            self.list_cols.addItem(item)
        layout.addWidget(self.list_cols)
        
        # Select All / Deselect All
        btn_layout = QHBoxLayout()
        btn_select_all = QPushButton("모두 선택")
        btn_deselect_all = QPushButton("모두 해제")
        btn_select_all.clicked.connect(self.select_all)
        btn_deselect_all.clicked.connect(self.deselect_all)
        btn_layout.addWidget(btn_select_all)
        btn_layout.addWidget(btn_deselect_all)
        layout.addLayout(btn_layout)
        
        layout.addSpacing(10)
        
        # Dialog Buttons
        self.buttonBox = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        self.buttonBox.accepted.connect(self.validate_and_accept)
        self.buttonBox.rejected.connect(self.reject)
        layout.addWidget(self.buttonBox)
        
    def select_all(self):
        for i in range(self.list_cols.count()):
            self.list_cols.item(i).setCheckState(Qt.CheckState.Checked)
            
    def deselect_all(self):
        for i in range(self.list_cols.count()):
            self.list_cols.item(i).setCheckState(Qt.CheckState.Unchecked)
            
    def validate_and_accept(self):
        if not self.get_selected_columns():
            QMessageBox.warning(self, "경고", "최소 1개 이상의 열을 선택해야 합니다.")
            return
        self.accept()
        
    def get_mode(self):
        return self.mode_group.checkedId()
        
    def get_only_status_o(self):
        return self.chk_status_o.isChecked()
        
    def get_selected_columns(self):
        cols = []
        for i in range(self.list_cols.count()):
            item = self.list_cols.item(i)
            if item.checkState() == Qt.CheckState.Checked:
                cols.append(item.text())
        return cols
