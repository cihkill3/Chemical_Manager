from PyQt6.QtWidgets import QSystemTrayIcon, QMenu
from PyQt6.QtGui import QIcon, QAction
from PyQt6.QtCore import pyqtSignal

class TrayIcon(QSystemTrayIcon):
    show_window_requested = pyqtSignal()
    exit_requested = pyqtSignal()
    
    def __init__(self, icon, parent=None):
        super().__init__(icon, parent)
        self.setToolTip("연구실 시약 자동 동기화")
        
        menu = QMenu()
        
        show_action = QAction("설정 화면 열기", self)
        show_action.triggered.connect(self.show_window_requested.emit)
        menu.addAction(show_action)
        
        menu.addSeparator()
        
        exit_action = QAction("종료", self)
        exit_action.triggered.connect(self.exit_requested.emit)
        menu.addAction(exit_action)
        
        self.setContextMenu(menu)
        self.activated.connect(self.on_activated)

    def on_activated(self, reason):
        if reason == QSystemTrayIcon.ActivationReason.DoubleClick:
            self.show_window_requested.emit()

    def show_message(self, title, msg):
        self.showMessage(title, msg, QSystemTrayIcon.MessageIcon.Information, 3000)
