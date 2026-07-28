from PyQt6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QLabel, 
                             QPushButton, QColorDialog, QGroupBox, QGridLayout)
from PyQt6.QtGui import QColor, QPalette

class ColorDialog(QDialog):
    def __init__(self, config_data, parent=None):
        super().__init__(parent)
        self.setWindowTitle("색상 범례 설정")
        self.resize(500, 400)
        
        self.config = config_data.copy()
        if "colors" not in self.config:
            self.config["colors"] = {}
        
        self.colors = self.config["colors"]
        
        self.color_map = [
            ("원본 시트 완료 (Source)", "done_bg", "done_font"),
            ("CAS 번호 없음 (경고)", "warn_cas_bg", "warn_cas_font"),
            ("품번 없음 (경고)", "warn_num_bg", "warn_num_font"),
            ("위치/온도 정보 없음", "warn_loc_bg", "warn_loc_font"),
            ("일반 항목", "normal_bg", "normal_font"),
            ("헤더 스타일", "header_bg", "header_font")
        ]
        
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout(self)
        
        grp = QGroupBox("색상 사용자 지정")
        grid = QGridLayout()
        
        grid.addWidget(QLabel("<b>항목</b>"), 0, 0)
        grid.addWidget(QLabel("<b>미리보기</b>"), 0, 1)
        grid.addWidget(QLabel("<b>배경색</b>"), 0, 2)
        grid.addWidget(QLabel("<b>글자색</b>"), 0, 3)
        
        self.preview_labels = {}
        
        for idx, (label_text, bg_key, fg_key) in enumerate(self.color_map):
            row = idx + 1
            
            grid.addWidget(QLabel(label_text), row, 0)
            
            preview = QLabel(" 샘플 텍스트 ")
            preview.setAutoFillBackground(True)
            self.preview_labels[bg_key] = preview
            grid.addWidget(preview, row, 1)
            
            btn_bg = QPushButton("배경색 변경")
            btn_bg.clicked.connect(lambda checked, b=bg_key, f=fg_key: self.choose_color(b, f, True))
            grid.addWidget(btn_bg, row, 2)
            
            btn_fg = QPushButton("글자색 변경")
            btn_fg.clicked.connect(lambda checked, b=bg_key, f=fg_key: self.choose_color(b, f, False))
            grid.addWidget(btn_fg, row, 3)
            
            self.update_preview(bg_key, fg_key)
            
        grp.setLayout(grid)
        layout.addWidget(grp)
        
        # Buttons
        h_btns = QHBoxLayout()
        btn_save = QPushButton("저장")
        btn_cancel = QPushButton("취소")
        btn_save.clicked.connect(self.save)
        btn_cancel.clicked.connect(self.reject)
        h_btns.addStretch()
        h_btns.addWidget(btn_save)
        h_btns.addWidget(btn_cancel)
        layout.addLayout(h_btns)

    def choose_color(self, bg_key, fg_key, is_bg):
        key = bg_key if is_bg else fg_key
        current_hex = self.colors.get(key, "#ffffff" if is_bg else "#000000")
        
        color = QColorDialog.getColor(QColor(current_hex), self, "색상 선택")
        if color.isValid():
            self.colors[key] = color.name() # Returns #RRGGBB
            self.update_preview(bg_key, fg_key)

    def update_preview(self, bg_key, fg_key):
        preview = self.preview_labels.get(bg_key)
        if not preview: return
        
        bg_hex = self.colors.get(bg_key, "#ffffff")
        fg_hex = self.colors.get(fg_key, "#000000")
        
        style = f"background-color: {bg_hex}; color: {fg_hex}; border: 1px solid #ccc;"
        preview.setStyleSheet(style)

    def save(self):
        self.config["colors"] = self.colors
        self.accept()

    def get_config(self):
        return self.config
