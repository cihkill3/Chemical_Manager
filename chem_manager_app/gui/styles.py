# Modern Bright Studio Theme QSS for Chemical Manager

MODERN_STYLE = """
/* Global Application Styling */
QWidget {
    font-family: 'Segoe UI', 'Pretendard', 'Malgun Gothic', sans-serif;
    font-size: 13px;
    color: #0F172A;
    background-color: #F8FAFC;
}

QMainWindow {
    background-color: #F8FAFC;
}

/* Status Header Bar */
#statusHeader {
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #EFF6FF, stop:1 #DBEAFE);
    border: 1px solid #BFDBFE;
    border-radius: 8px;
    padding: 6px 14px;
    margin-bottom: 2px;
}

#statusHeader QLabel {
    font-weight: 700;
    font-size: 14px;
    color: #1E40AF;
    background: transparent;
    padding: 1px 0px;
}

.statusBadge {
    background-color: #FFFFFF;
    border: 1px solid #0284C7;
    border-radius: 10px;
    padding: 3px 10px;
    font-size: 12px;
    font-weight: 600;
    color: #0284C7;
}

/* Modern Card GroupBoxes with Compact Vertical Margins */
QGroupBox {
    background-color: #FFFFFF;
    border: 1px solid #E2E8F0;
    border-radius: 8px;
    margin-top: 10px;
    padding: 10px 12px 8px 12px;
    font-weight: 700;
    font-size: 13px;
    color: #0284C7;
}

QGroupBox::title {
    subcontrol-origin: margin;
    subcontrol-position: top left;
    left: 10px;
    top: 2px;
    padding: 0 4px;
    background: transparent;
    color: #0284C7;
    font-weight: 700;
    font-size: 13px;
}

/* LineEdits and SpinBoxes */
QLineEdit {
    background-color: #FFFFFF;
    border: 1px solid #CBD5E1;
    border-radius: 6px;
    padding: 5px 8px;
    color: #0F172A;
    selection-background-color: #0284C7;
}

QLineEdit:focus {
    border: 1.5px solid #0284C7;
    background-color: #FFFFFF;
}

QLineEdit[readOnly="true"] {
    background-color: #F1F5F9;
    color: #475569;
}

/* SpinBox Arrows */
QSpinBox {
    background-color: #FFFFFF;
    border: 1px solid #CBD5E1;
    border-radius: 6px;
    padding: 4px 26px 4px 8px;
    color: #0F172A;
    font-weight: bold;
    selection-background-color: #0284C7;
}

QSpinBox:focus {
    border: 1.5px solid #0284C7;
}

QSpinBox::up-button {
    subcontrol-origin: border;
    subcontrol-position: top right;
    width: 22px;
    height: 13px;
    border-left: 1px solid #CBD5E1;
    border-bottom: 1px solid #CBD5E1;
    background: #E2E8F0;
    border-top-right-radius: 5px;
}

QSpinBox::down-button {
    subcontrol-origin: border;
    subcontrol-position: bottom right;
    width: 22px;
    height: 13px;
    border-left: 1px solid #CBD5E1;
    background: #E2E8F0;
    border-bottom-right-radius: 5px;
}

QSpinBox::up-button:hover, QSpinBox::down-button:hover {
    background: #0284C7;
}

QSpinBox::up-arrow {
    width: 0px;
    height: 0px;
    border-left: 4px solid transparent;
    border-right: 4px solid transparent;
    border-bottom: 5px solid #0F172A;
}

QSpinBox::up-button:hover QSpinBox::up-arrow {
    border-bottom-color: #FFFFFF;
}

QSpinBox::down-arrow {
    width: 0px;
    height: 0px;
    border-left: 4px solid transparent;
    border-right: 4px solid transparent;
    border-top: 5px solid #0F172A;
}

QSpinBox::down-button:hover QSpinBox::down-arrow {
    border-top-color: #FFFFFF;
}

/* CheckBoxes and RadioButtons with Checkmark Icon */
QCheckBox, QRadioButton {
    spacing: 6px;
    color: #334155;
    font-size: 13px;
    background: transparent;
}

QCheckBox::indicator, QRadioButton::indicator {
    width: 16px;
    height: 16px;
    border-radius: 4px;
    border: 1.5px solid #94A3B8;
    background-color: #FFFFFF;
}

QRadioButton::indicator {
    border-radius: 8px;
}

QCheckBox::indicator:hover, QRadioButton::indicator:hover {
    border-color: #0284C7;
}

QCheckBox::indicator:checked {
    background-color: #0284C7;
    border: 1.5px solid #0284C7;
}

QRadioButton::indicator:checked {
    background-color: #0284C7;
    border: 4px solid #FFFFFF;
    outline: 1.5px solid #0284C7;
}

/* Buttons */
QPushButton {
    background-color: #F1F5F9;
    color: #1E293B;
    border: 1px solid #CBD5E1;
    border-radius: 6px;
    padding: 6px 14px;
    font-weight: 600;
    font-size: 13px;
}

QPushButton:hover {
    background-color: #E2E8F0;
    border-color: #94A3B8;
}

QPushButton:pressed {
    background-color: #CBD5E1;
}

QPushButton:disabled {
    background-color: #F8FAFC;
    color: #94A3B8;
    border-color: #E2E8F0;
}

/* Primary Action Buttons */
#btn_sync {
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #2563EB, stop:1 #1D4ED8);
    border: none;
    color: #FFFFFF;
    font-weight: 700;
}
#btn_sync:hover {
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #3B82F6, stop:1 #2563EB);
}

#btn_db_update {
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #059669, stop:1 #047857);
    border: none;
    color: #FFFFFF;
    font-weight: 700;
}
#btn_db_update:hover {
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #10B981, stop:1 #059669);
}

#btn_pdf {
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #7C3AED, stop:1 #6D28D9);
    border: none;
    color: #FFFFFF;
    font-weight: 700;
}
#btn_pdf:hover {
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #8B5CF6, stop:1 #7C3AED);
}

#btn_sds {
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #D97706, stop:1 #B45309);
    border: none;
    color: #FFFFFF;
    font-weight: 700;
}
#btn_sds:hover {
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #F59E0B, stop:1 #D97706);
}

/* Automation Toggle Button */
#btn_toggle_auto {
    background-color: #F1F5F9;
    border: 1px solid #CBD5E1;
    color: #1E293B;
}
#btn_toggle_auto:checked {
    background-color: #FEE2E2;
    border-color: #FCA5A5;
    color: #991B1B;
}

/* Console Log Window */
QTextEdit {
    background-color: #F8FAFC;
    border: 1px solid #E2E8F0;
    border-radius: 6px;
    font-family: 'Consolas', 'Cascadia Code', 'Courier New', monospace;
    font-size: 12px;
    color: #0F172A;
    padding: 6px;
}

/* List Widgets */
QListWidget {
    background-color: #FFFFFF;
    border: 1px solid #CBD5E1;
    border-radius: 6px;
    padding: 4px;
    color: #0F172A;
}

QListWidget::item {
    padding: 5px 8px;
    border-radius: 4px;
}

QListWidget::item:hover {
    background-color: #F1F5F9;
}

QListWidget::item:selected {
    background-color: #0284C7;
    color: #FFFFFF;
}

/* ScrollBars */
QScrollBar:vertical {
    background: #F8FAFC;
    width: 8px;
    margin: 0;
}
QScrollBar::handle:vertical {
    background: #CBD5E1;
    min-height: 20px;
    border-radius: 4px;
}
QScrollBar::handle:vertical:hover {
    background: #94A3B8;
}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
    height: 0px;
}

/* Dialogs */
QDialog {
    background-color: #F8FAFC;
}

/* StatusBar */
QStatusBar {
    background-color: #FFFFFF;
    color: #475569;
    border-top: 1px solid #E2E8F0;
}
"""
