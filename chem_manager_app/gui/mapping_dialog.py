from PyQt6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit,
                             QPushButton, QListWidget, QComboBox, QFormLayout, QGroupBox,
                             QMessageBox, QListWidgetItem, QWidget)

class MappingDialog(QDialog):
    def __init__(self, config_data, parent=None):
        super().__init__(parent)
        self.setWindowTitle("헤더 매핑 설정")
        self.resize(600, 500)
        self.config = config_data.copy()
        
        self.src_headers = list(self.config.get("source_headers", []))
        self.tgt_headers = list(self.config.get("target_headers", []))
        self.mapping = dict(self.config.get("mapping", {}))
        
        self.init_ui()

    def init_ui(self):
        main_layout = QHBoxLayout(self)

        # Left Column: Manage Headers
        left_layout = QVBoxLayout()
        
        # Source Headers Group
        grp_src = QGroupBox("오더북 (Source) 헤더 관리")
        l_src = QVBoxLayout()
        self.lst_src = QListWidget()
        self.lst_src.addItems(self.src_headers)
        
        h_src_btn = QHBoxLayout()
        self.inp_src = QLineEdit()
        btn_add_src = QPushButton("추가")
        btn_del_src = QPushButton("삭제")
        btn_add_src.clicked.connect(self.add_src)
        btn_del_src.clicked.connect(self.del_src)
        h_src_btn.addWidget(self.inp_src)
        h_src_btn.addWidget(btn_add_src)
        h_src_btn.addWidget(btn_del_src)
        
        l_src.addWidget(self.lst_src)
        l_src.addLayout(h_src_btn)
        grp_src.setLayout(l_src)
        left_layout.addWidget(grp_src)

        # Target Headers Group
        grp_tgt = QGroupBox("ChemicalList (Target) 헤더 관리")
        l_tgt = QVBoxLayout()
        self.lst_tgt = QListWidget()
        self.lst_tgt.addItems(self.tgt_headers)
        
        h_tgt_btn = QHBoxLayout()
        self.inp_tgt = QLineEdit()
        btn_add_tgt = QPushButton("추가")
        btn_del_tgt = QPushButton("삭제")
        btn_add_tgt.clicked.connect(self.add_tgt)
        btn_del_tgt.clicked.connect(self.del_tgt)
        h_tgt_btn.addWidget(self.inp_tgt)
        h_tgt_btn.addWidget(btn_add_tgt)
        h_tgt_btn.addWidget(btn_del_tgt)
        
        l_tgt.addWidget(self.lst_tgt)
        l_tgt.addLayout(h_tgt_btn)
        grp_tgt.setLayout(l_tgt)
        left_layout.addWidget(grp_tgt)

        main_layout.addLayout(left_layout)

        # Right Column: Mapping
        right_layout = QVBoxLayout()
        grp_map = QGroupBox("헤더 매핑 (Target <- Source)")
        self.form_map = QFormLayout()
        self.map_combos = {}
        
        self.map_widget = QWidget()
        self.map_widget.setLayout(self.form_map)
        right_layout.addWidget(grp_map)
        
        l_map = QVBoxLayout()
        l_map.addWidget(self.map_widget)
        grp_map.setLayout(l_map)
        
        # Buttons
        h_btns = QHBoxLayout()
        btn_save = QPushButton("저장")
        btn_cancel = QPushButton("취소")
        btn_save.clicked.connect(self.save)
        btn_cancel.clicked.connect(self.reject)
        h_btns.addStretch()
        h_btns.addWidget(btn_save)
        h_btns.addWidget(btn_cancel)
        right_layout.addLayout(h_btns)
        
        main_layout.addLayout(right_layout)
        
        self.refresh_mapping_ui()

    def add_src(self):
        text = self.inp_src.text().strip()
        if text and text not in self.src_headers:
            self.src_headers.append(text)
            self.lst_src.addItem(text)
            self.inp_src.clear()
            self.refresh_mapping_ui()

    def del_src(self):
        row = self.lst_src.currentRow()
        if row >= 0:
            item = self.lst_src.takeItem(row)
            self.src_headers.remove(item.text())
            self.refresh_mapping_ui()

    def add_tgt(self):
        text = self.inp_tgt.text().strip()
        if text and text not in self.tgt_headers:
            self.tgt_headers.append(text)
            self.lst_tgt.addItem(text)
            self.inp_tgt.clear()
            self.refresh_mapping_ui()

    def del_tgt(self):
        row = self.lst_tgt.currentRow()
        if row >= 0:
            item = self.lst_tgt.takeItem(row)
            self.tgt_headers.remove(item.text())
            if item.text() in self.mapping:
                del self.mapping[item.text()]
            self.refresh_mapping_ui()

    def refresh_mapping_ui(self):
        # Save current selections
        for t_head, combo in self.map_combos.items():
            self.mapping[t_head] = combo.currentText()
        
        # Clear form
        while self.form_map.count():
            child = self.form_map.takeAt(0)
            if child.widget():
                child.widget().deleteLater()
        
        self.map_combos.clear()
        
        # Rebuild
        for t_head in self.tgt_headers:
            cb = QComboBox()
            cb.addItem("")
            cb.addItems(self.src_headers)
            
            # Select mapped item if exists
            mapped = self.mapping.get(t_head, "")
            if mapped in self.src_headers:
                cb.setCurrentText(mapped)
            
            self.form_map.addRow(t_head + " : ", cb)
            self.map_combos[t_head] = cb

    def save(self):
        # Finalize mapping
        for t_head, combo in self.map_combos.items():
            self.mapping[t_head] = combo.currentText()
            
        self.config["source_headers"] = self.src_headers
        self.config["target_headers"] = self.tgt_headers
        self.config["mapping"] = self.mapping
        
        self.accept()

    def get_config(self):
        return self.config
